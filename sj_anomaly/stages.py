"""stages.py — replika stage-by-stage Power Query M code '04.QUERRY SJ.xlsb'.

Mapping stage → M code (query ``3  42 15`` kecuali disebut lain):
    stage_01  → Cleaned Headers + Renamed QTY Column          (step 1-2)
    stage_02  → Filtered Rows                                  (step 3 + filter junk report '3 1 42 15 Kirim')
    stage_03  → Changed Type2                                  (step 4)
    stage_04  → Added Custom KTG / Concet / Extract GT         (step 5)
    stage_05  → Extracted Text Between Delimiters              (step 6)
    stage_06  → Replaced Value Depo → DCSMGMUR                 (step 7)
    stage_07  → Added Selisih Tgl-Tgl Create                   (step 8)
    stage_08  → Cleaned Text                                   (step 9)
    stageA    → Calculated Standar Satuan + Status Anomali Satuan
    stageB    → Calculated QTY Stats (Rata2/Jumlah per item+satuan+tujuan)
    stageC    → Added Status Anomali QTY (5 rule)
    stageD    → Added QTY Terkecil 2 (rasio konversi)
    stageE    → Calculated QTY Terkecil 2 Stats (mode + frekuensi)
    stageF    → Merged QTY Terkecil 2 Stats (join back)
    stageGK   → Selisih / Deviasi % / Dominasi % / Rasio / Arah
    stageLM   → Status Anomali QTY Terkecil 2 + narasi Analisa
    merge_all  → query 'Merge4205&4215' (join Kirim×Terima×Kategori)
    count_groups → query '4215TerimaCount' / '4215KirimCount'

Semua konstanta ambang diambil dari config['thresholds'] (bisa di-override
tanpa ubah kode — mirror Waste).
"""
from __future__ import annotations

import math
from typing import Any

import pandas as pd

from .loader import (
    _find_column,
    between_delimiters,
    diff_days,
    parse_date_value,
    parse_number_value,
    text_trim,
)
from .utils import fmt_num, fmt_pct, fmt_ratio


class StageError(RuntimeError):
    """Kesalahan stage — pesan ramah untuk UI."""


# ============================================================
# KONSTANTA RULE (dari M code — bisa di-override via config)
# ============================================================

def _get_thresholds(config: dict) -> dict:
    """Ambil konstanta ambang dari config (fallback default M code)."""
    t = config.get("thresholds", {})
    return {
        "QTY_EXTREME_FACTOR":  float(t.get("QTY_EXTREME_FACTOR", 3.5)),
        "QTY_TINY_FACTOR":    float(t.get("QTY_TINY_FACTOR", 0.1)),
        "MIN_TRANSAKSI":      int(t.get("MIN_TRANSAKSI", 3)),
        "DEVIASI_ANOMALI":    float(t.get("DEVIASI_ANOMALI", 0.10)),
        "DEVIASI_PERLU_CEK":  float(t.get("DEVIASI_PERLU_CEK", 0.05)),
        "DOMINASI_MIN":       float(t.get("DOMINASI_MIN", 0.50)),
        "TERIMA_LAMBAT_HARI": int(t.get("TERIMA_LAMBAT_HARI", 2)),
    }


def _get_whole_units(config: dict) -> set[str]:
    """Satuan utuh: QTY desimal pada satuan ini = anomali (rule C.3)."""
    return {s.upper() for s in config.get("whole_units", [])} or {
        "PCS", "KARTON", "KARUNG", "JURIGEN", "PACK", "BOTTLE", "BOX", "CUP",
    }


def _get_diterima_status(config: dict) -> set[str]:
    """Status yang dihitung sebagai 'diterima' untuk statistik QTY (rule B)."""
    return set(config.get("diterima_status", [])) or {
        "Diterima Seluruh nya", "Diterima Seluruhnya",
    }


def _get_depo_replacement(config: dict) -> str:
    """Replace 'Depo' → nilai ini pada Gudang Pengirim (step 7 M code)."""
    return config.get("depo_replacement", "DCSMGMUR")


def is_junk_no_pemindahan(v: Any) -> bool:
    """Judul report ACCURATE yang nyangkut di kolom No. Pemindahan # (filter 3.1)."""
    if v is None:
        return True
    s = text_trim(v)
    if not s:
        return True
    return (
        s == "ACCURATE Accounting System Report"
        or s == "Halaman 1"
        or s.startswith("Tercetak pada")
    )


# ============================================================
# STAGE 1-3 — LOAD KIRIM: header, filter, tipe
# ============================================================

def stage_01_03_load_kirim(
    df_raw: pd.DataFrame,
    config: dict,
) -> pd.DataFrame:
    """Stage 1-3: mapping sheet mentah → baris kirim bertipe.

    Step 1: Cleaned Headers (Text.Trim pada nama kolom)
    Step 2: Rename 'Kts Diterima' → 'QTY Terkecil'
    Step 3: Filter Status Pemindahan Barang ≠ null + junk report
    Step 4: Type cast Tanggal (date), Tgl/Jam Pembuatan (datetime),
            QTY Terkecil & Kuantitas (number)
    """
    kirim_cfg = config.get("kirim_columns", {})

    # Resolve nama kolom asli dari kandidat
    col = {
        "no_pemindahan": _find_column(df_raw.columns, kirim_cfg.get("no_pemindahan", ["No. Pemindahan #"])),
        "tanggal":       _find_column(df_raw.columns, kirim_cfg.get("tanggal", ["Tanggal"])),
        "tgl_create":    _find_column(df_raw.columns, kirim_cfg.get("tgl_create", ["Tgl/Jam Pembuatan"])),
        "pengirim":      _find_column(df_raw.columns, kirim_cfg.get("pengirim", ["Gudang Pengirim"])),
        "tujuan":        _find_column(df_raw.columns, kirim_cfg.get("tujuan", ["Gudang Tujuan"])),
        "kode":          _find_column(df_raw.columns, kirim_cfg.get("kode", ["Kode Barang"])),
        "nama":          _find_column(df_raw.columns, kirim_cfg.get("nama", ["Nama Barang"])),
        "satuan":        _find_column(df_raw.columns, kirim_cfg.get("satuan", ["Satuan"])),
        "qty_terkecil":  _find_column(df_raw.columns, kirim_cfg.get("qty_terkecil", ["Kts Diterima", "QTY Terkecil"])),
        "kuantitas":     _find_column(df_raw.columns, kirim_cfg.get("kuantitas", ["Kuantitas"])),
        "status":        _find_column(df_raw.columns, kirim_cfg.get("status", ["Status Pemindahan Barang"])),
    }

    missing = [k for k, v in col.items() if not v and k in ("nama", "satuan", "status")]
    if missing:
        raise StageError(
            f"Kolom wajib hilang dari sheet Kirim (42.15): {missing}. "
            f"Header tersedia: {list(df_raw.columns)[:15]}."
        )

    rows: list[dict] = []
    for _, r in df_raw.iterrows():
        # Step 3: filter Status Pemindahan Barang ≠ null (baris judul/subtotal gugur)
        status_val = text_trim(r[col["status"]]) if col["status"] else ""
        if not status_val:
            continue
        # Filter junk report ACCURATE (dari query '3 1 42 15 Kirim')
        no_raw = r[col["no_pemindahan"]] if col["no_pemindahan"] else None
        if is_junk_no_pemindahan(no_raw):
            continue
        # Filter Tanggal ≠ null (query '3 1 42 15 Kirim' step 'Filtered Rows1')
        tanggal = parse_date_value(r[col["tanggal"]]) if col["tanggal"] else None
        if not tanggal:
            continue

        tgl_create = (
            parse_date_value(r[col["tgl_create"]], with_time=True)
            if col["tgl_create"] else None
        )

        rows.append({
            "No. Pemindahan #":      text_trim(no_raw) if no_raw is not None else None,
            "Tanggal":               tanggal,
            "Tgl/Jam Pembuatan":     tgl_create,
            "Gudang Pengirim":       str(r[col["pengirim"]]) if col["pengirim"] else "",
            "Gudang Tujuan":         str(r[col["tujuan"]]) if col["tujuan"] else "",
            "Kode Barang":           text_trim(r[col["kode"]]) if col["kode"] else None,
            "Nama Barang":           text_trim(r[col["nama"]]) if col["nama"] else "",
            "Satuan":                text_trim(r[col["satuan"]]) if col["satuan"] else "",
            "QTY Terkecil":          parse_number_value(r[col["qty_terkecil"]]) if col["qty_terkecil"] else None,
            "Kuantitas":             parse_number_value(r[col["kuantitas"]]) if col["kuantitas"] else None,
            "Status Pemindahan Barang": text_trim(status_val),
            # placeholder kolom turunan (diisi stage berikutnya)
            "KTG [RESTO/DC]":        "",
            "Concet":                "",
            "Extract [Gudang Tujuan]": "",
            "Gudang Pengirim (Kode)": "",
            "Gudang Tujuan (Kode)":  "",
            "Selisih Tgl-Tgl Create": None,
            "Satuan Standar (Auto)": None,
            "Jumlah Variasi Satuan": 0,
            "Status Anomali Satuan": "NORMAL",
            "Rata2 QTY Resto":       None,
            "Jumlah Transaksi Resto": 0,
            "Status Anomali QTY":    "NORMAL",
            "QTY Terkecil 2":        None,
            "Standar QTY Terkecil 2": None,
            "Frekuensi Standar QTY Terkecil 2": None,
            "Jumlah Transaksi Valid Ratio": 0,
            "Jumlah Variasi QTY Terkecil 2": 0,
            "Selisih QTY Terkecil 2": None,
            "Deviasi QTY Terkecil 2 %": None,
            "Dominasi Standar QTY Terkecil 2 %": None,
            "Rasio terhadap Standar QTY Terkecil 2": None,
            "Arah Variasi QTY Terkecil 2": None,
            "Status Anomali QTY Terkecil 2": "NORMAL",
            "Analisa QTY Terkecil 2": "",
        })

    if not rows:
        raise StageError(
            "Tidak ada baris valid di sheet Kirim (42.15) — "
            "cek kolom 'Status Pemindahan Barang' & 'Tanggal'."
        )

    df = pd.DataFrame(rows)
    # Catat jumlah baris junk yang dibuang (untuk log pipeline)
    df.attrs["dropped_junk"] = max(0, len(df_raw) - len(rows))
    return df


# ============================================================
# STAGE 4-8 — KTG, Concet, ekstrak gudang, selisih, trim
# ============================================================

def ktg_from_raw(raw: str) -> str:
    """Step 5: KTG [RESTO/DC] — 'DC' bila kode pengirim mulai '2'/'5'."""
    s = text_trim(raw)
    return "DC" if (s.startswith("2") or s.startswith("5")) else "RESTO"


def stage_04_08_transform(df: pd.DataFrame, config: dict) -> None:
    """Stage 4-8: KTG, Concet, ekstrak gudang, replace Depo, selisih, trim.

    Step 5: KTG [RESTO/DC] — dievaluasi pada nilai MENTAH (sebelum ekstrak)
    Step 6: Concet = Pengirim & Tujuan & Nama Barang (nilai mentah)
    Step 7: Extract [Gudang Tujuan] — teks dalam kurung (kolom terpisah)
    Step 8: ekstrak kode/nama dalam kurung untuk kedua gudang
    Step 9: replace 'Depo' → DCSMGMUR pada Gudang Pengirim hasil ekstrak
    Step 10: Selisih Tgl - Tgl Create (hari)
    Step 11: bersihkan teks (Text.Trim)
    """
    depo_repl = _get_depo_replacement(config)

    for idx, r in df.iterrows():
        pengirim_raw = r["Gudang Pengirim"]
        tujuan_raw = r["Gudang Tujuan"]
        nama = r["Nama Barang"]

        # Step 5: KTG
        ktg = ktg_from_raw(pengirim_raw)

        # Step 6: Concet
        concet = f"{text_trim(pengirim_raw)}{text_trim(tujuan_raw)}{text_trim(nama)}"

        # Step 7: Extract [Gudang Tujuan] (kolom terpisah)
        extract_tujuan = between_delimiters(tujuan_raw)

        # Step 8: ekstrak kode gudang
        pengirim_kode = between_delimiters(pengirim_raw)
        tujuan_kode = between_delimiters(tujuan_raw)

        # Step 9: replace 'Depo' → DCSMGMUR
        if pengirim_kode == "Depo":
            pengirim_kode = depo_repl

        # Step 10: Selisih Tgl - Tgl Create (hari)
        tgl_create_only = r["Tgl/Jam Pembuatan"][:10] if r["Tgl/Jam Pembuatan"] else None
        selisih = diff_days(tgl_create_only, r["Tanggal"])

        # Tulis kembali (in-place safe via df.loc)
        df.at[idx, "KTG [RESTO/DC]"] = ktg
        df.at[idx, "Concet"] = concet
        df.at[idx, "Extract [Gudang Tujuan]"] = extract_tujuan
        df.at[idx, "Gudang Pengirim (Kode)"] = pengirim_kode
        df.at[idx, "Gudang Tujuan (Kode)"] = tujuan_kode
        df.at[idx, "Selisih Tgl-Tgl Create"] = selisih
        # Step 11: bersihkan teks (sudah dilakukan di stage_01_03, tapi pastikan)
        df.at[idx, "Nama Barang"] = text_trim(nama)
        df.at[idx, "Satuan"] = text_trim(r["Satuan"])
        df.at[idx, "Gudang Tujuan (Kode)"] = text_trim(tujuan_kode)
        df.at[idx, "Status Pemindahan Barang"] = text_trim(r["Status Pemindahan Barang"])


# ============================================================
# STAGE A — ANOMALI SATUAN (mode satuan per Nama Barang)
# ============================================================

def stageA_anomali_satuan(df: pd.DataFrame) -> None:
    """Stage A: deteksi anomali satuan (mode per Nama Barang).

    Mirror M code 'Calculated Standar Satuan' + 'Status Anomali Satuan':
      - Group by Nama Barang → Satuan Standar (Auto) = mode (count desc)
      - Jumlah Variasi Satuan = nunique satuan
      - Status = 'ANOMALI (Biasa {std} -> Input {sat})' bila variasi>1 & sat≠std
                 else 'NORMAL'
    """
    # Group by Nama Barang → hitung variasi + mode satuan
    by_item: dict[str, dict[str, int]] = {}
    for nama in df["Nama Barang"]:
        if not nama:
            continue
        if nama not in by_item:
            by_item[nama] = {}
        # satuan untuk item ini akan diisi di loop kedua

    # Bangun count satuan per item
    satuan_counts: dict[str, dict[str, int]] = {}
    for _, r in df.iterrows():
        nama = r["Nama Barang"]
        sat = r["Satuan"]
        if not nama or not sat:
            continue
        satuan_counts.setdefault(nama, {})
        satuan_counts[nama][sat] = satuan_counts[nama].get(sat, 0) + 1

    # Hitung mode + variasi per item, assign ke baris
    for idx, r in df.iterrows():
        nama = r["Nama Barang"]
        sat = r["Satuan"]
        if not nama or not sat:
            continue
        counts = satuan_counts.get(nama, {})
        variasi = len(counts)
        # Mode: count desc lalu nama asc (deterministik — mirror M Sort)
        best = None
        best_cnt = -1
        for s in sorted(counts.keys()):
            if counts[s] > best_cnt:
                best = s
                best_cnt = counts[s]
        df.at[idx, "Satuan Standar (Auto)"] = best
        df.at[idx, "Jumlah Variasi Satuan"] = variasi
        df.at[idx, "Status Anomali Satuan"] = (
            f"ANOMALI (Biasa {best} -> Input {sat})"
            if (variasi > 1 and sat != best)
            else "NORMAL"
        )


# ============================================================
# STAGE B — STATISTIK QTY per (Nama Barang, Satuan, Gudang Tujuan)
# ============================================================

def stageB_qty_stats(df: pd.DataFrame, config: dict) -> None:
    """Stage B: statistik QTY per (Nama Barang, Satuan, Gudang Tujuan).

    Mirror M code 'Calculated QTY Stats':
      - Filter status 'Diterima Seluruh nya' / 'Diterima Seluruhnya'
      - Rata2 QTY Resto = mean(QTY Terkecil)
      - Jumlah Transaksi Resto = count baris terfilter
    """
    diterima = _get_diterima_status(config)

    # Akumulasi per key
    acc: dict[tuple, dict] = {}
    for _, r in df.iterrows():
        if r["Status Pemindahan Barang"] not in diterima:
            continue
        key = (r["Nama Barang"], r["Satuan"], r["Gudang Tujuan (Kode)"])
        if key not in acc:
            acc[key] = {"qty_sum": 0.0, "qty_n": 0, "tx_n": 0}
        a = acc[key]
        a["tx_n"] += 1
        if r["QTY Terkecil"] is not None:
            a["qty_sum"] += float(r["QTY Terkecil"])
            a["qty_n"] += 1

    # Assign ke baris
    for idx, r in df.iterrows():
        key = (r["Nama Barang"], r["Satuan"], r["Gudang Tujuan (Kode)"])
        a = acc.get(key)
        if not a:
            df.at[idx, "Rata2 QTY Resto"] = None
            df.at[idx, "Jumlah Transaksi Resto"] = 0
            continue
        df.at[idx, "Rata2 QTY Resto"] = (a["qty_sum"] / a["qty_n"]) if a["qty_n"] > 0 else None
        df.at[idx, "Jumlah Transaksi Resto"] = a["tx_n"]


# ============================================================
# STAGE C — STATUS ANOMALI QTY (5 rule berurutan)
# ============================================================

def stageC_anomali_qty(df: pd.DataFrame, config: dict) -> None:
    """Stage C: 5 rule anomali QTY (mirror M code 'Added Status Anomali QTY').

    Rule 1: status bukan 'Diterima Seluruh*' → NORMAL
    Rule 2: QTY null / ≤ 0 → 'ANOMALI QTY (Nol / Minus)'
    Rule 3: satuan utuh + QTY desimal → 'ANOMALI QTY (Satuan Utuh ber-Desimal: …)'
    Rule 4: lonjakan ekstrem QTY > 3.5 × rata2 peer → 'ANOMALI QTY (Lonjakan Ekstrem: …)'
    Rule 5: terlalu kecil QTY < 0.1 × rata2 peer → 'ANOMALI QTY (Terlalu Kecil vs Rata2 …)'
    """
    th = _get_thresholds(config)
    whole_units = _get_whole_units(config)
    diterima = _get_diterima_status(config)
    min_tx = th["MIN_TRANSAKSI"]
    extreme = th["QTY_EXTREME_FACTOR"]
    tiny = th["QTY_TINY_FACTOR"]

    for idx, r in df.iterrows():
        v_status = r["Status Pemindahan Barang"]
        v_qty = r["QTY Terkecil"]
        v_avg = r["Rata2 QTY Resto"]
        v_sat = (r["Satuan"] or "").upper()
        v_count = r["Jumlah Transaksi Resto"]

        is_whole = v_sat in whole_units
        is_decimal = (
            v_qty is not None
            and isinstance(v_qty, (int, float))
            and abs(float(v_qty) % 1) > 1e-9
        )

        # Rule 1
        if v_status not in diterima:
            df.at[idx, "Status Anomali QTY"] = "NORMAL"
            continue
        # Rule 2
        if v_qty is None or float(v_qty) <= 0:
            df.at[idx, "Status Anomali QTY"] = "ANOMALI QTY (Nol / Minus)"
            continue
        # Rule 3
        if is_whole and is_decimal:
            df.at[idx, "Status Anomali QTY"] = (
                f"ANOMALI QTY (Satuan Utuh ber-Desimal: {fmt_num(v_qty)})"
            )
            continue
        # Rule 4
        if (
            v_count >= min_tx
            and v_avg is not None
            and v_avg > 0
            and float(v_qty) > extreme * v_avg
        ):
            df.at[idx, "Status Anomali QTY"] = (
                f"ANOMALI QTY (Lonjakan Ekstrem: {fmt_num(v_qty)} vs Rata2 {fmt_num(v_avg)})"
            )
            continue
        # Rule 5
        if (
            v_count >= min_tx
            and v_avg is not None
            and v_avg > 0
            and float(v_qty) < tiny * v_avg
        ):
            df.at[idx, "Status Anomali QTY"] = (
                f"ANOMALI QTY (Terlalu Kecil vs Rata2 {fmt_num(v_avg)})"
            )
            continue
        df.at[idx, "Status Anomali QTY"] = "NORMAL"


# ============================================================
# STAGE D — QTY TERKECIL 2 (rasio konversi)
# ============================================================

def stageD_qty_terkecil_2(df: pd.DataFrame) -> None:
    """Stage D: QTY Terkecil 2 = QTY Terkecil / Kuantitas (mirror M code step D).

    Null bila QTY Terkecil / Kuantitas null atau Kuantitas = 0.
    """
    for idx, r in df.iterrows():
        qty = r["QTY Terkecil"]
        kuantitas = r["Kuantitas"]
        if (
            qty is None
            or kuantitas is None
            or (isinstance(kuantitas, (int, float)) and float(kuantitas) == 0)
        ):
            df.at[idx, "QTY Terkecil 2"] = None
        else:
            df.at[idx, "QTY Terkecil 2"] = float(qty) / float(kuantitas)


# ============================================================
# STAGE E-F — STANDAR RASIO (mode per Nama Barang + Satuan)
# ============================================================

def stageEF_standar_rasio(df: pd.DataFrame) -> None:
    """Stage E-F: standar rasio (mode per Nama Barang + Satuan) + merge back.

    Mirror M code 'Calculated QTY Terkecil 2 Stats':
      - Group by (Nama Barang, Satuan) → frekuensi rasio (round 6)
      - Standar = mode (count desc, ratio desc) — mirror M Sort
      - Frekuensi Standar = count dari mode
      - Jumlah Transaksi Valid Ratio = total baris dengan rasio non-null
      - Jumlah Variasi QTY Terkecil 2 = nunique rasio round 6
    """
    # Akumulasi frekuensi rasio per (Nama Barang, Satuan)
    acc: dict[tuple, dict] = {}
    for _, r in df.iterrows():
        ratio = r["QTY Terkecil 2"]
        if ratio is None or not isinstance(ratio, (int, float)) or not math.isfinite(ratio):
            continue
        rounded = round(float(ratio), 6)
        key = (r["Nama Barang"], r["Satuan"])
        if key not in acc:
            acc[key] = {"freq": {}, "total": 0, "variasi": set()}
        a = acc[key]
        a["freq"][rounded] = a["freq"].get(rounded, 0) + 1
        a["total"] += 1
        a["variasi"].add(rounded)

    # Assign ke baris
    for idx, r in df.iterrows():
        key = (r["Nama Barang"], r["Satuan"])
        a = acc.get(key)
        if not a:
            df.at[idx, "Standar QTY Terkecil 2"] = None
            df.at[idx, "Frekuensi Standar QTY Terkecil 2"] = None
            df.at[idx, "Jumlah Transaksi Valid Ratio"] = 0
            df.at[idx, "Jumlah Variasi QTY Terkecil 2"] = 0
            continue
        df.at[idx, "Jumlah Transaksi Valid Ratio"] = a["total"]
        df.at[idx, "Jumlah Variasi QTY Terkecil 2"] = len(a["variasi"])
        # Mode: count desc lalu rasio desc (mirror M: {"Cnt",desc},{"Ratio",desc})
        mode_val = None
        mode_cnt = -1
        for ratio in sorted(a["freq"].keys(), reverse=True):
            if a["freq"][ratio] > mode_cnt:
                mode_cnt = a["freq"][ratio]
                mode_val = ratio
        df.at[idx, "Standar QTY Terkecil 2"] = mode_val
        df.at[idx, "Frekuensi Standar QTY Terkecil 2"] = mode_cnt


# ============================================================
# STAGE G-K — SELISIH / DEVIASI / DOMINASI / RASIO / ARAH
# ============================================================

def stageGK_selisih_deviasi(df: pd.DataFrame) -> None:
    """Stage G-K: Selisih, Deviasi %, Dominasi %, Rasio, Arah.

    G. Selisih = actual − standar
    H. Deviasi % = (actual−std)/std
    I. Dominasi % = Frekuensi / Total
    J. Rasio = actual/std
    K. Arah = LEBIH BESAR / LEBIH KECIL / SESUAI STANDAR
    """
    for idx, r in df.iterrows():
        actual = r["QTY Terkecil 2"]
        std = r["Standar QTY Terkecil 2"]
        total = r["Jumlah Transaksi Valid Ratio"]
        freq = r["Frekuensi Standar QTY Terkecil 2"]

        # G. Selisih
        if actual is not None and std is not None:
            df.at[idx, "Selisih QTY Terkecil 2"] = float(actual) - float(std)
        else:
            df.at[idx, "Selisih QTY Terkecil 2"] = None

        # H. Deviasi %
        if (
            actual is not None
            and std is not None
            and isinstance(std, (int, float))
            and float(std) != 0
        ):
            df.at[idx, "Deviasi QTY Terkecil 2 %"] = (float(actual) - float(std)) / float(std)
        else:
            df.at[idx, "Deviasi QTY Terkecil 2 %"] = None

        # I. Dominasi Standar %
        if (
            isinstance(total, (int, float)) and total > 0
            and freq is not None
        ):
            df.at[idx, "Dominasi Standar QTY Terkecil 2 %"] = float(freq) / float(total)
        else:
            df.at[idx, "Dominasi Standar QTY Terkecil 2 %"] = None

        # J. Rasio terhadap Standar
        if (
            actual is not None
            and std is not None
            and isinstance(std, (int, float))
            and float(std) != 0
        ):
            df.at[idx, "Rasio terhadap Standar QTY Terkecil 2"] = float(actual) / float(std)
        else:
            df.at[idx, "Rasio terhadap Standar QTY Terkecil 2"] = None

        # K. Arah Variasi
        dev = df.at[idx, "Deviasi QTY Terkecil 2 %"]
        if dev is None:
            df.at[idx, "Arah Variasi QTY Terkecil 2"] = None
        elif dev > 0:
            df.at[idx, "Arah Variasi QTY Terkecil 2"] = "LEBIH BESAR"
        elif dev < 0:
            df.at[idx, "Arah Variasi QTY Terkecil 2"] = "LEBIH KECIL"
        else:
            df.at[idx, "Arah Variasi QTY Terkecil 2"] = "SESUAI STANDAR"


# ============================================================
# STAGE L-M — STATUS ANOMALI KONVERSI + NARASI
# ============================================================

def stageLM_status_konversi(df: pd.DataFrame, config: dict) -> None:
    """Stage L-M: status anomali konversi + narasi (mirror M code).

    L. Status (rule berurutan):
        - actual null → TIDAK DAPAT DIANALISA
        - std null → TIDAK ADA STANDAR
        - count < 3 → PERLU CEK - DATA TERBATAS / DATA TERBATAS
        - dominasi < 50% → KONVERSI TIDAK STABIL
        - |deviasi| > 10% → ANOMALI KONVERSI
        - |deviasi| > 5% → PERLU CEK
        - else → NORMAL
    M. Narasi teks (persis M code).
    """
    th = _get_thresholds(config)
    min_tx = th["MIN_TRANSAKSI"]
    anom = th["DEVIASI_ANOMALI"]
    perlu = th["DEVIASI_PERLU_CEK"]
    dominasi_min = th["DOMINASI_MIN"]

    for idx, r in df.iterrows():
        v_ratio = r["QTY Terkecil 2"]
        v_std = r["Standar QTY Terkecil 2"]
        v_dev = r["Deviasi QTY Terkecil 2 %"]
        v_count = r["Jumlah Transaksi Valid Ratio"]
        v_dom = r["Dominasi Standar QTY Terkecil 2 %"]

        if v_ratio is None:
            status = "TIDAK DAPAT DIANALISA"
        elif v_std is None:
            status = "TIDAK ADA STANDAR"
        elif not isinstance(v_count, (int, float)) or v_count < min_tx:
            if v_dev is not None and abs(v_dev) > anom:
                status = "PERLU CEK - DATA TERBATAS"
            else:
                status = "DATA TERBATAS"
        elif v_dom is not None and v_dom < dominasi_min:
            status = "KONVERSI TIDAK STABIL"
        elif v_dev is not None and abs(v_dev) > anom:
            status = "ANOMALI KONVERSI"
        elif v_dev is not None and abs(v_dev) > perlu:
            status = "PERLU CEK"
        else:
            status = "NORMAL"
        df.at[idx, "Status Anomali QTY Terkecil 2"] = status

        # M. Narasi
        actual_text = "-" if v_ratio is None else fmt_ratio(v_ratio)
        std_text = "-" if v_std is None else fmt_ratio(v_std)
        dev_text = "-" if v_dev is None else fmt_pct(v_dev)

        if status == "ANOMALI KONVERSI":
            narasi = f"QTY Terkecil 2 {actual_text} menyimpang {dev_text} dari standar {std_text}"
        elif status == "PERLU CEK":
            narasi = f"QTY Terkecil 2 {actual_text} berbeda {dev_text} dari standar {std_text}"
        elif status == "KONVERSI TIDAK STABIL":
            narasi = f"Nilai QTY Terkecil 2 memiliki banyak variasi; standar {std_text} tidak dominan"
        elif status == "PERLU CEK - DATA TERBATAS":
            narasi = "QTY Terkecil 2 berbeda signifikan tetapi jumlah data historis masih terbatas"
        elif status == "DATA TERBATAS":
            narasi = "Belum cukup transaksi untuk membentuk standar rasio yang kuat"
        elif status == "TIDAK ADA STANDAR":
            narasi = "Belum tersedia standar QTY Terkecil 2"
        elif status == "TIDAK DAPAT DIANALISA":
            narasi = "QTY Terkecil 2 tidak dapat dihitung karena Kuantitas kosong atau 0"
        else:
            narasi = "QTY Terkecil 2 konsisten dengan standar"
        df.at[idx, "Analisa QTY Terkecil 2"] = narasi


# ============================================================
# LOAD TERIMA (42.05) & KATEGORI (21  DBS Kategori)
# ============================================================

def load_terima_rows(df_raw: pd.DataFrame, config: dict) -> pd.DataFrame:
    """Map sheet Terima mentah → DataFrame terstruktur (mirror loadTerimaRows)."""
    terima_cfg = config.get("terima_columns", {})
    col = {
        "nomor_kirim":    _find_column(df_raw.columns, terima_cfg.get("nomor_kirim", ["Nomor #Kirim"])),
        "nomor_terima":   _find_column(df_raw.columns, terima_cfg.get("nomor_terima", ["Nomor #Terima"])),
        "gudang_kirim":   _find_column(df_raw.columns, terima_cfg.get("gudang_kirim", ["Gudang #Kirim"])),
        "gudang_terima":  _find_column(df_raw.columns, terima_cfg.get("gudang_terima", ["Gudang #Terima"])),
        "status":         _find_column(df_raw.columns, terima_cfg.get("status", ["Status Pengiriman #"])),
        "tgl_kirim":      _find_column(df_raw.columns, terima_cfg.get("tgl_kirim", ["Tanggal #Kirim"])),
        "tgl_terima":     _find_column(df_raw.columns, terima_cfg.get("tgl_terima", ["Tanggal #Terima"])),
        "vs":             _find_column(df_raw.columns, terima_cfg.get("vs", ["#Tgl Kirim vs Tgl Terima"])),
        "kode":           _find_column(df_raw.columns, terima_cfg.get("kode", ["Kode Barang"])),
        "nama":           _find_column(df_raw.columns, terima_cfg.get("nama", ["Nama Barang"])),
        "sat_kirim":      _find_column(df_raw.columns, terima_cfg.get("sat_kirim", ["#Sat. Kirim"])),
        "qty_kirim":      _find_column(df_raw.columns, terima_cfg.get("qty_kirim", ["#Qty Kirim"])),
        "rasio":          _find_column(df_raw.columns, terima_cfg.get("rasio", ["Rasio Satuan"])),
        "sat_terkecil":   _find_column(df_raw.columns, terima_cfg.get("sat_terkecil", ["#Sat. Terkecil"])),
        "qty_terkecil":   _find_column(df_raw.columns, terima_cfg.get("qty_terkecil", ["#Qty. Terkecil"])),
        "tgl_it":         _find_column(df_raw.columns, terima_cfg.get("tgl_it", ["#Tgl/Jam Pembuatan IT"])),
        "dok":            _find_column(df_raw.columns, terima_cfg.get("dok", ["Memiliki dokumen (Terima)"])),
    }
    if not col["nomor_kirim"]:
        raise StageError(
            f"Kolom 'Nomor #Kirim' tidak ditemukan di sheet Terima (42.05). "
            f"Header tersedia: {list(df_raw.columns)[:15]}"
        )

    rows = []
    for _, r in df_raw.iterrows():
        nk = text_trim(r[col["nomor_kirim"]]) if col["nomor_kirim"] else None
        if not nk:
            continue  # baris tanpa Nomor #Kirim = judul/kosong
        rows.append({
            "Nomor #Kirim":              nk,
            "Nomor #Terima":             text_trim(r[col["nomor_terima"]]) if col["nomor_terima"] else None,
            "Gudang #Kirim":             text_trim(r[col["gudang_kirim"]]) if col["gudang_kirim"] else None,
            "Gudang #Terima":            text_trim(r[col["gudang_terima"]]) if col["gudang_terima"] else None,
            "Status Pengiriman #":       text_trim(r[col["status"]]) if col["status"] else None,
            "Tanggal #Kirim":            parse_date_value(r[col["tgl_kirim"]]) if col["tgl_kirim"] else None,
            "Tanggal #Terima":           parse_date_value(r[col["tgl_terima"]]) if col["tgl_terima"] else None,
            "#Tgl Kirim vs Tgl Terima":  text_trim(r[col["vs"]]) if col["vs"] else None,
            "Kode Barang":               text_trim(r[col["kode"]]) if col["kode"] else None,
            "Nama Barang":               text_trim(r[col["nama"]]) if col["nama"] else None,
            "#Sat. Kirim":               text_trim(r[col["sat_kirim"]]) if col["sat_kirim"] else None,
            "#Qty Kirim":                parse_number_value(r[col["qty_kirim"]]) if col["qty_kirim"] else None,
            "Rasio Satuan":              parse_number_value(r[col["rasio"]]) if col["rasio"] else None,
            "#Sat. Terkecil":            text_trim(r[col["sat_terkecil"]]) if col["sat_terkecil"] else None,
            "#Qty. Terkecil":            parse_number_value(r[col["qty_terkecil"]]) if col["qty_terkecil"] else None,
            "#Tgl/Jam Pembuatan IT":     parse_date_value(r[col["tgl_it"]], with_time=True) if col["tgl_it"] else None,
            "Memiliki dokumen (Terima)": text_trim(r[col["dok"]]) if col["dok"] else None,
        })
    return pd.DataFrame(rows)


def load_kategori_rows(df_raw: pd.DataFrame, config: dict) -> pd.DataFrame:
    """Map sheet DBS Kategori mentah → DataFrame terstruktur."""
    kat_cfg = config.get("kategori_columns", {})
    col_nama = _find_column(df_raw.columns, kat_cfg.get("nama", ["NAMA BARANG"]))
    col_kat = _find_column(df_raw.columns, kat_cfg.get("kat", ["KATEGORI"]))
    col_jenis = _find_column(df_raw.columns, kat_cfg.get("jenis", ["JENIS"]))
    if not col_nama or not col_jenis:
        raise StageError(
            f"Kolom 'NAMA BARANG'/'JENIS' tidak ditemukan di sheet DBS Kategori. "
            f"Header tersedia: {list(df_raw.columns)[:10]}"
        )
    rows = []
    for _, r in df_raw.iterrows():
        nb = text_trim(r[col_nama])
        if not nb:
            continue
        rows.append({
            "NAMA BARANG": nb,
            "KATEGORI":    text_trim(r[col_kat]) if col_kat else "",
            "JENIS":       text_trim(r[col_jenis]),
        })
    return pd.DataFrame(rows)


# ============================================================
# MERGE — query 'Merge4205&4215'
# ============================================================

def merge_all(
    df_kirim: pd.DataFrame,
    df_terima: pd.DataFrame,
    df_kategori: pd.DataFrame,
    config: dict,
) -> pd.DataFrame:
    """Merge Kirim × Terima × Kategori (mirror M code 'Merge4205&4215').

    - Left join Kirim ← Terima on No. Pemindahan # == Nomor #Kirim
    - Left join Kirim ← Kategori on Nama Barang (case-insensitive)
    - Kolom turunan:
        * Selisih Hari = Tanggal #Terima − Tanggal (date diff)
        * KTG_Kiriman[DC/RESTO] = dari Gudang Pengirim (mentah)
        * JENIS = kategori (default 'EQP' bila null — persis M code)
    """
    # Index terima by Nomor #Kirim (first match — M NestedJoin bisa multi-match,
    # tapi kami ambil first; duplikat terima di-catat sbg warning).
    terima_idx: dict[str, dict] = {}
    terima_dup: set[str] = set()
    for _, t in df_terima.iterrows():
        k = t["Nomor #Kirim"]
        if k in terima_idx:
            terima_dup.add(k)
        else:
            terima_idx[k] = t.to_dict()

    # Index kategori by NAMA BARANG (uppercase)
    kat_idx: dict[str, dict] = {}
    for _, k in df_kategori.iterrows():
        kat_idx[k["NAMA BARANG"].upper()] = k.to_dict()

    rows = []
    for _, r in df_kirim.iterrows():
        no = r["No. Pemindahan #"]
        t = terima_idx.get(no) if no else None
        kat = kat_idx.get((r["Nama Barang"] or "").upper())

        # KTG_Kiriman[DC/RESTO] — dari Gudang Pengirim mentah
        ktg_kiriman = ktg_from_raw(r["Gudang Pengirim"])

        # Selisih Hari = Tanggal #Terima − Tanggal
        selisih_hari = None
        if t and t.get("Tanggal #Terima"):
            selisih_hari = diff_days(r["Tanggal"], t["Tanggal #Terima"])

        merged = dict(r.to_dict())
        merged.update({
            "tNomorKirim":            t.get("Nomor #Kirim") if t else None,
            "tNomorTerima":           t.get("Nomor #Terima") if t else None,
            "tGudangKirim":           t.get("Gudang #Kirim") if t else None,
            "tGudangTerima":          t.get("Gudang #Terima") if t else None,
            "tStatusPengiriman":      t.get("Status Pengiriman #") if t else None,
            "tTanggalKirim":          t.get("Tanggal #Kirim") if t else None,
            "tTanggalTerima":         t.get("Tanggal #Terima") if t else None,
            "tTglKirimVsTerima":      t.get("#Tgl Kirim vs Tgl Terima") if t else None,
            "tSatKirim":              t.get("#Sat. Kirim") if t else None,
            "tQtyKirim":              t.get("#Qty Kirim") if t else None,
            "tRasioSatuan":           t.get("Rasio Satuan") if t else None,
            "tSatTerkecil":           t.get("#Sat. Terkecil") if t else None,
            "tQtyTerkecil":           t.get("#Qty. Terkecil") if t else None,
            "Memiliki Dokumen [Terima]": t is not None,
            "Selisih Hari":           selisih_hari,
            "KTG_Kiriman[DC/RESTO]":  ktg_kiriman,
            "JENIS (Kategori)":       (kat.get("JENIS") if kat else None) or "EQP",
            "KATEGORI (Master)":      kat.get("KATEGORI") if kat else None,
        })
        rows.append(merged)

    df_merged = pd.DataFrame(rows)

    # Back-fill "Count" per Concet (mirror M code 4215TerimaCount → 4215Terima&Count
    # join). DETAIL_COLS di exporter membaca _g("Count") — tanpa ini, kolom
    # "Count Terima (Concet)" di sheet Detail/Anomali akan kosong.
    if "Concet" in df_merged.columns and len(df_merged):
        counts = df_merged.groupby("Concet")["Concet"].transform("count").astype(int)
        df_merged["Count"] = counts
    else:
        df_merged["Count"] = 0

    df_merged.attrs["terima_dup_count"] = len(terima_dup)
    df_merged.attrs["tanpa_terima"] = sum(1 for r in rows if not r["Memiliki Dokumen [Terima]"])
    return df_merged


# ============================================================
# COUNT — 4215TerimaCount / 4215KirimCount (group by Concet)
# ============================================================

def count_groups(df_kirim: pd.DataFrame, only_dup: bool = False) -> pd.DataFrame:
    """Group by Concet → count (mirror M '4215TerimaCount' / '4215KirimCount').

    Return DataFrame: Concet | Count | Contoh Barang | Contoh Tanggal.
    """
    groups: dict[str, dict] = {}
    for _, r in df_kirim.iterrows():
        c = r["Concet"]
        if c not in groups:
            groups[c] = {
                "Concet": c,
                "Count": 0,
                "Contoh Barang": r["Nama Barang"],
                "Contoh Tanggal": r["Tanggal"],
            }
        groups[c]["Count"] += 1

    rows = list(groups.values())
    rows.sort(key=lambda g: g["Count"], reverse=True)
    if only_dup:
        rows = [g for g in rows if g["Count"] > 1]
    return pd.DataFrame(rows, columns=["Concet", "Count", "Contoh Barang", "Contoh Tanggal"])
