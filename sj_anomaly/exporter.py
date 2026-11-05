"""exporter.py — penulis output Excel multi-sheet (mirror Waste exporter).

Sheet:
  1. Detail             — semua baris merge, kolom mengikuti M code final
  2. Anomali Prioritas   — hanya baris dengan salah satu status ≠ NORMAL / dokumen terima hilang
  3. Ringkasan          — KPI + timestamp + log ringkas
  4. Rekap Count         — grup Concet dengan count > 1 (potensi duplikat)
  5. Log Run             — log pipeline timestamped

Mengikuti pola Waste: WriteOnlyCell streaming utk data besar (> threshold),
mode normal utk data kecil supaya formatting penuh (border, font, autofilter).
"""
from __future__ import annotations

import os
from datetime import datetime
from pathlib import Path

import numpy as np
import pandas as pd
from openpyxl import Workbook
from openpyxl.cell import WriteOnlyCell
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

from .utils import fmt_date_id, severity_of_status


# ============================================================
# STYLE CONSTANTS — object di-share (mirror Waste L15)
# ============================================================
HEADER_FILL = PatternFill(start_color="172033", end_color="172033", fill_type="solid")
HEADER_FONT = Font(name="Segoe UI", size=11, bold=True, color="FFFFFF")
HEADER_ALIGN = Alignment(horizontal="center", vertical="center", wrap_text=True)
TITLE_FONT = Font(name="Segoe UI", size=14, bold=True, color="172033")
BODY_FONT = Font(name="Segoe UI", size=10, color="1F2937")
BODY_ALIGN = Alignment(vertical="top", wrap_text=False)
EMPTY_NOTE_FONT = Font(name="Segoe UI", size=10, italic=True, color="6B7280")

STATUS_FILLS = {
    "anomali":    PatternFill(start_color="DC2626", end_color="DC2626", fill_type="solid"),
    "perlu-cek":  PatternFill(start_color="EA580C", end_color="EA580C", fill_type="solid"),
    "perhatian":  PatternFill(start_color="F59E0B", end_color="F59E0B", fill_type="solid"),
    "normal":     PatternFill(start_color="16A34A", end_color="16A34A", fill_type="solid"),
}
STATUS_FONT = Font(name="Segoe UI", size=10, bold=True, color="FFFFFF")
STATUS_ALIGN = Alignment(horizontal="center", vertical="center")

THIN_BORDER = Border(
    left=Side(style="thin", color="E5E7EB"),
    right=Side(style="thin", color="E5E7EB"),
    top=Side(style="thin", color="E5E7EB"),
    bottom=Side(style="thin", color="E5E7EB"),
)

STREAMING_THRESHOLD_DEFAULT = 35_000


# ============================================================
# MAIN ENTRY
# ============================================================

def export_to_excel(
    df_merged: pd.DataFrame,
    output_path: str | os.PathLike,
    config: dict,
    summary: dict,
    log_lines: list[str],
    extra_sheets: dict[str, pd.DataFrame] | None = None,
) -> dict:
    """Tulis multi-sheet xlsx. Return summary dict.

    Return contract (mirror Waste):
      - sukses: {"success": True, "output_path": <str>, ...}
      - gagal : {"success": False, "error": <pesan>, "output_path": <str>}
    """
    output_path = Path(output_path)
    try:
        output_path.parent.mkdir(parents=True, exist_ok=True)
    except OSError:
        return {
            "success": False,
            "error": f"Tidak bisa membuat folder output '{output_path.parent}' — cek permission",
            "output_path": str(output_path),
        }

    # Pilih mode: streaming (memori konstan) utk data besar, normal utk kecil
    try:
        threshold = int(config.get("output", {}).get(
            "streaming_threshold", STREAMING_THRESHOLD_DEFAULT
        ))
    except (TypeError, ValueError):
        threshold = STREAMING_THRESHOLD_DEFAULT
    use_streaming = len(df_merged) > max(0, threshold)

    try:
        if use_streaming:
            wb = Workbook(write_only=True)
            _write = _write_sheet_streaming
        else:
            wb = Workbook()
            wb.remove(wb.active)
            _write = _write_sheet

        sheet_cfg = config.get("output", {})

        # 1. Detail
        detail_name = sheet_cfg.get("detail_sheet", "Detail")
        cols = DETAIL_COLS
        _write(wb, detail_name, df_merged, cols, freeze="A2", status_col_idx=_status_idx(df_merged, cols))

        # 2. Anomali Prioritas
        anomali_name = sheet_cfg.get("anomali_sheet", "Anomali Prioritas")
        df_anomali = _filter_anomali(df_merged)
        _write(wb, anomali_name, df_anomali, cols, freeze="A2",
               status_col_idx=_status_idx(df_anomali, cols))

        # 3. Ringkasan
        ringkasan_name = sheet_cfg.get("ringkasan_sheet", "Ringkasan")
        df_ringkasan = _build_ringkasan(summary, df_merged)
        _write(wb, ringkasan_name, df_ringkasan, freeze="A2")

        # 4. Rekap Count
        rekap_name = sheet_cfg.get("rekap_count_sheet", "Rekap Count")
        df_count = summary.get("df_count")
        if df_count is None or df_count.empty:
            df_rekap = pd.DataFrame(columns=["Concet", "Count", "Contoh Barang", "Contoh Tanggal"])
        else:
            df_rekap = df_count[df_count["Count"] > 1].copy()
            if "Contoh Tanggal" in df_rekap.columns:
                df_rekap["Contoh Tanggal"] = df_rekap["Contoh Tanggal"].apply(
                    lambda x: fmt_date_id(x) if x else None
                )
        _write(wb, rekap_name, df_rekap, freeze="A2")

        # 5. Extra sheets (opsional)
        n_extra = 0
        for sheet_name, extra_df in (extra_sheets or {}).items():
            if extra_df is None or extra_df.empty:
                continue
            _write(wb, str(sheet_name), extra_df, freeze="A2")
            n_extra += 1

        # 6. Sheet views tambahan — mirror sheet di file asli '04.QUERRY SJ.xlsb'
        # Generate 8 view: Upload, Terima Pivot, Manual, Kirim Pivot,
        # Double Terima, Double Kirim, 4215 Terima, 4215 Kirim
        try:
            from .sheets import generate_all_views
            df_kirim = summary.get("df")
            if df_kirim is not None and df_count is not None:
                views = generate_all_views(df_kirim, df_merged, df_count)
                for vname, vdf in views.items():
                    if vdf is None or vdf.empty:
                        continue
                    _write(wb, vname, vdf, freeze="A2")
                    n_extra += 1
        except Exception as exc:
            # Views tambahan gagal = tidak fatal, lanjutkan ke Log Run
            import logging
            logging.getLogger(__name__).warning("sheets views gagal: %s", exc)

        # 7. Log Run
        log_name = sheet_cfg.get("log_sheet", "Log Run")
        df_log = _build_log(summary, log_lines)
        _write(wb, log_name, df_log, freeze="A2")

        try:
            wb.save(output_path)
        except PermissionError:
            return {
                "success": False,
                "error": "File output sedang dibuka di Excel — tutup dulu lalu ulangi",
                "output_path": str(output_path),
            }
        except MemoryError:
            return {
                "success": False,
                "error": ("Memori tidak cukup saat menulis Excel. Tutup aplikasi lain "
                          "lalu ulangi, atau pecah data menjadi beberapa bagian."),
                "output_path": str(output_path),
            }
    except MemoryError:
        return {
            "success": False,
            "error": ("Memori tidak cukup saat menyiapkan laporan. Tutup aplikasi lain "
                      "lalu ulangi, atau pecah data menjadi beberapa bagian."),
            "output_path": str(output_path),
        }

    return {
        "success": True,
        "output_path": str(output_path),
        "streaming": use_streaming,
        "total_rows": int(len(df_merged)),
        "anomali_count": int(len(df_anomali)),
        "extra_sheets_count": n_extra,
        "anomali_satuan": summary.get("anomali_satuan", 0),
        "anomali_qty": summary.get("anomali_qty", 0),
        "anomali_konversi": summary.get("anomali_konversi", 0),
        "kirim_tanpa_terima": summary.get("kirim_tanpa_terima", 0),
    }


# ============================================================
# KOLOM EXPORT (mirror Waste DETAIL_COLS)
# ============================================================

class ColSpec:
    """Spec kolom export: header + getter + lebar."""
    __slots__ = ("header", "get", "wch", "status_key")
    def __init__(self, header: str, get, wch: int = 14, status_key: str | None = None):
        self.header = header
        self.get = get
        self.wch = wch
        self.status_key = status_key


def _g(field: str):
    """Getter generik: ambil field dari row dict, None bila tidak ada."""
    def _get(r):
        v = r.get(field)
        return v
    return _get


DETAIL_COLS: list[ColSpec] = [
    ColSpec("No. Pemindahan",         _g("No. Pemindahan #"), 16),
    ColSpec("Tanggal",                lambda r: fmt_date_id(r.get("Tanggal")) if r.get("Tanggal") else None, 12),
    ColSpec("Tgl/Jam Buat IT",        lambda r: r.get("Tgl/Jam Pembuatan"), 18),
    ColSpec("Gudang Pengirim",        _g("Gudang Pengirim (Kode)"), 16),
    ColSpec("Gudang Tujuan",          _g("Gudang Tujuan (Kode)"), 20),
    ColSpec("Kode",                   lambda r: int(r.get("Kode Barang")) if r.get("Kode Barang") and str(r.get("Kode Barang")).isdigit() else r.get("Kode Barang"), 8),
    ColSpec("Nama Barang",            _g("Nama Barang"), 26),
    ColSpec("Satuan",                 _g("Satuan"), 9),
    ColSpec("QTY Terkecil",           _g("QTY Terkecil"), 12),
    ColSpec("Kuantitas",              _g("Kuantitas"), 10),
    ColSpec("Status Pemindahan",      _g("Status Pemindahan Barang"), 20),
    ColSpec("KTG [RESTO/DC]",         _g("KTG [RESTO/DC]"), 13),
    ColSpec("Selisih Tgl-Create",    _g("Selisih Tgl-Tgl Create"), 14),
    ColSpec("Satuan Standar (Auto)", _g("Satuan Standar (Auto)"), 14),
    ColSpec("Jml Variasi Satuan",     _g("Jumlah Variasi Satuan"), 12),
    ColSpec("Status Anomali Satuan",  _g("Status Anomali Satuan"), 28, status_key="satuan"),
    ColSpec("Rata2 QTY Resto",        _g("Rata2 QTY Resto"), 13),
    ColSpec("Jml Transaksi Resto",    _g("Jumlah Transaksi Resto"), 12),
    ColSpec("Status Anomali QTY",     _g("Status Anomali QTY"), 38, status_key="qty"),
    ColSpec("QTY Terkecil 2",          _g("QTY Terkecil 2"), 13),
    ColSpec("Standar QTY Terkecil 2", _g("Standar QTY Terkecil 2"), 14),
    ColSpec("Frekuensi Standar",       _g("Frekuensi Standar QTY Terkecil 2"), 12),
    ColSpec("Jml Transaksi Valid",    _g("Jumlah Transaksi Valid Ratio"), 12),
    ColSpec("Jml Variasi QTY2",       _g("Jumlah Variasi QTY Terkecil 2"), 12),
    ColSpec("Selisih QTY2",           _g("Selisih QTY Terkecil 2"), 12),
    ColSpec("Deviasi QTY2 %",         _g("Deviasi QTY Terkecil 2 %"), 12),
    ColSpec("Dominasi Standar %",     _g("Dominasi Standar QTY Terkecil 2 %"), 12),
    ColSpec("Rasio thd Standar",      _g("Rasio terhadap Standar QTY Terkecil 2"), 12),
    ColSpec("Arah Variasi",           _g("Arah Variasi QTY Terkecil 2"), 14),
    ColSpec("Status Anomali QTY2",    _g("Status Anomali QTY Terkecil 2"), 26, status_key="konversi"),
    ColSpec("Analisa QTY2",           _g("Analisa QTY Terkecil 2"), 48),
    ColSpec("Dok Terima (42.05)",     lambda r: "YA" if r.get("Memiliki Dokumen [Terima]") else "TIDAK", 12, status_key="dok"),
    ColSpec("Nomor Terima",           _g("tNomorTerima"), 16),
    ColSpec("Tanggal Terima",         lambda r: fmt_date_id(r.get("tTanggalTerima")) if r.get("tTanggalTerima") else None, 13),
    ColSpec("Selisih Hari",           _g("Selisih Hari"), 11),
    ColSpec("KTG Kiriman",            _g("KTG_Kiriman[DC/RESTO]"), 11),
    ColSpec("JENIS (Kategori)",       _g("JENIS (Kategori)"), 12),
    ColSpec("Count Terima (Concet)",  _g("Count"), 12),
]


def _status_idx(df: pd.DataFrame, cols: list[ColSpec]) -> list[int]:
    """Index kolom yang ber-status (diwarnai berdasarkan severity)."""
    # Karena DETAIL_COLS punya beberapa kolom status (satuan/qty/konversi/dok),
    # kembalikan index untuk semua. Streaming writer pakai ini.
    return [i for i, c in enumerate(cols) if c.status_key is not None]


def _filter_anomali(df: pd.DataFrame) -> pd.DataFrame:
    """Filter baris anomali (salah satu status ≠ NORMAL atau dokumen terima hilang)."""
    if df.empty:
        return df
    mask = (
        (df["Status Anomali Satuan"] != "NORMAL")
        | (df["Status Anomali QTY"] != "NORMAL")
        | df["Status Anomali QTY Terkecil 2"].astype(str).str.startswith("ANOMALI")
        | df["Status Anomali QTY Terkecil 2"].astype(str).str.startswith("PERLU CEK")
        | (~df["Memiliki Dokumen [Terima]"])
    )
    return df[mask].copy()


# ============================================================
# SHEET BUILDERS (Ringkasan / Log)
# ============================================================

def _build_ringkasan(summary: dict, df: pd.DataFrame) -> pd.DataFrame:
    """Sheet Ringkasan — KPI + meta info (mirror Waste _build_log)."""
    rows = [
        ("SJ Insight — Ringkasan Run", ""),
        ("Waktu Run", datetime.now().strftime("%Y-%m-%d %H:%M:%S")),
        ("Sumber Kirim", summary.get("kirim_path", "-")),
        ("Sumber Terima", summary.get("terima_path", "-")),
        ("Sumber Kategori", summary.get("kategori_path", "-")),
        ("", ""),
        ("Total Transaksi", summary.get("total_transaksi", 0)),
        ("Total Barang Unik", summary.get("total_barang", 0)),
        ("Total Gudang Tujuan", summary.get("total_gudang_tujuan", 0)),
        ("Anomali Satuan", summary.get("anomali_satuan", 0)),
        ("Anomali QTY", summary.get("anomali_qty", 0)),
        ("Anomali Konversi", summary.get("anomali_konversi", 0)),
        ("Perlu Cek (konversi)", summary.get("perlu_cek", 0)),
        ("Data Terbatas / Tidak Stabil", summary.get("data_terbatas", 0)),
        ("Baris Normal Bersih", summary.get("normal", 0)),
        ("Kirim tanpa Dokumen Terima", summary.get("kirim_tanpa_terima", 0)),
        ("Terima Lambat (≥2 hari)", summary.get("terima_lambat", 0)),
        ("Grup Concet Duplikat", summary.get("duplikat_concet", 0)),
        ("Rata2 Selisih Hari", summary.get("rata2_selisih_hari") or "-"),
        ("Tanggal Min", fmt_date_id(summary.get("tanggal_min")) if summary.get("tanggal_min") else "-"),
        ("Tanggal Max", fmt_date_id(summary.get("tanggal_max")) if summary.get("tanggal_max") else "-"),
    ]
    return pd.DataFrame(rows, columns=["Key", "Value"])


def _build_log(summary: dict, log_lines: list[str]) -> pd.DataFrame:
    """Sheet Log Run — log pipeline timestamped."""
    rows = [
        ("Waktu", "Persen", "Pesan"),
    ]
    # log_lines sudah berformat "[HH:MM:SS] msg (N%)" — pisahkan utk tampilan
    import re
    for line in log_lines:
        m = re.match(r"^\[(\d{2}:\d{2}:\d{2})\]\s+(.*?)\s*\((\d+)%\)\s*$", line)
        if m:
            rows.append((m.group(1), int(m.group(3)), m.group(2)))
        else:
            # baris tanpa persen
            m2 = re.match(r"^\[(\d{2}:\d{2}:\d{2})\]\s+(.*)$", line)
            if m2:
                rows.append((m2.group(1), "", m2.group(2)))
            else:
                rows.append(("", "", line))
    return pd.DataFrame(rows[1:], columns=rows[0])


# ============================================================
# SHEET WRITER — STREAMING MODE (mirror Waste F10)
# ============================================================

def _dedupe_sheet_name(sheet_name: str, existing: list[str]) -> str:
    """Truncate 31 char + dedupe."""
    sheet_name = sheet_name[:31]
    if sheet_name not in existing:
        return sheet_name
    base = sheet_name[:27]
    i = 2
    while f"{base} ({i})" in existing:
        i += 1
    return f"{base} ({i})"


def _to_cell_value(val):
    """NaN/NA → None; numpy scalar → python scalar; list/dict → str."""
    if val is None or val is pd.NA:
        return None
    if isinstance(val, float) and (val != val):  # NaN
        return None
    if isinstance(val, (np.integer,)):
        return int(val)
    if isinstance(val, (np.floating,)):
        f = float(val)
        return None if (f != f) else f
    if isinstance(val, (np.bool_,)):
        return bool(val)
    if isinstance(val, (list, dict)):
        return str(val)
    return val


def _column_cell_values(series: pd.Series) -> list:
    """Konversi satu kolom → list python scalar siap openpyxl."""
    dtype = series.dtype
    if pd.api.types.is_float_dtype(dtype) or pd.api.types.is_integer_dtype(dtype):
        return series.astype(object).where(series.notna(), None).tolist()
    return [_to_cell_value(v) for v in series.tolist()]


def _write_sheet_streaming(
    wb: Workbook,
    sheet_name: str,
    df: pd.DataFrame,
    cols: list[ColSpec] | None = None,
    freeze: str | None = "A2",
    status_col_idx: list[int] | None = None,
):
    """Tulis DataFrame ke sheet write-only — memori KONSTAN."""
    sheet_name = _dedupe_sheet_name(sheet_name, wb.sheetnames)
    ws = wb.create_sheet(title=sheet_name)

    if df.empty:
        cell = WriteOnlyCell(ws, value=f"(Tidak ada data untuk sheet '{sheet_name}')")
        cell.font = EMPTY_NOTE_FONT
        ws.append([cell])
        return

    # Pakai ColSpec bila diberikan, else pakai df.columns apa adanya
    if cols is not None:
        headers = [c.header for c in cols]
        getters = [c.get for c in cols]
        # Bangun body dari dict-per-baris
        records = df.to_dict(orient="records")
        body_rows = [[_to_cell_value(g(r)) for g in getters] for r in records]
        widths = [c.wch for c in cols]
    else:
        headers = list(df.columns)
        body_rows = [[_to_cell_value(v) for v in row] for row in df.itertuples(index=False, name=None)]
        widths = [14] * len(headers)

    # Dimensi harus di-set SEBELUM baris (kontrak write-only openpyxl)
    for c_idx, (col_name, w) in enumerate(zip(headers, widths), start=1):
        max_len = max(len(str(col_name)), min(
            max((len(str(s)) for s in body_rows[:100] if s and s[c_idx - 1] is not None), default=8), 60
        ))
        ws.column_dimensions[get_column_letter(c_idx)].width = min(max_len + 2, w)
    ws.row_dimensions[1].height = 28
    if freeze:
        ws.freeze_panes = freeze

    # Header (styled)
    header_cells = []
    for h in headers:
        hc = WriteOnlyCell(ws, value=str(h))
        hc.fill = HEADER_FILL
        hc.font = HEADER_FONT
        hc.alignment = HEADER_ALIGN
        hc.border = THIN_BORDER
        header_cells.append(hc)
    ws.append(header_cells)

    # Body
    for row in body_rows:
        ws.append([v if v is not None else None for v in row])

    # Auto filter
    last_col = get_column_letter(len(headers))
    ws.auto_filter.ref = f"A1:{last_col}{len(body_rows) + 1}"


def _write_sheet(
    wb: Workbook,
    sheet_name: str,
    df: pd.DataFrame,
    cols: list[ColSpec] | None = None,
    freeze: str | None = "A2",
    status_col_idx: list[int] | None = None,
):
    """Tulis DataFrame ke sheet baru dengan header styled + auto-width."""
    sheet_name = _dedupe_sheet_name(sheet_name, wb.sheetnames)
    ws = wb.create_sheet(title=sheet_name)

    if df.empty:
        ws["A1"] = f"(Tidak ada data untuk sheet '{sheet_name}')"
        ws["A1"].font = EMPTY_NOTE_FONT
        return

    if cols is not None:
        headers = [c.header for c in cols]
        getters = [c.get for c in cols]
        records = df.to_dict(orient="records")
        body_rows = [[_to_cell_value(g(r)) for g in getters] for r in records]
        widths = [c.wch for c in cols]
        # Mapping index status: cari ColSpec dengan status_key
        status_keys_idx = [(i, c.status_key) for i, c in enumerate(cols) if c.status_key is not None]
    else:
        headers = list(df.columns)
        body_rows = [[_to_cell_value(v) for v in row] for row in df.itertuples(index=False, name=None)]
        widths = [14] * len(headers)
        status_keys_idx = []

    # Header
    ws.append(headers)
    for cell in ws[1]:
        cell.fill = HEADER_FILL
        cell.font = HEADER_FONT
        cell.alignment = HEADER_ALIGN
        cell.border = THIN_BORDER

    # Body
    for row in body_rows:
        ws.append(row)

    # Styling body
    for row_cells in ws.iter_rows(min_row=2, max_row=len(body_rows) + 1):
        for cell in row_cells:
            cell.font = BODY_FONT
            cell.border = THIN_BORDER
            cell.alignment = BODY_ALIGN
        # Status coloring
        for idx, key in status_keys_idx:
            if idx >= len(row_cells):
                continue
            v = row_cells[idx].value
            if key == "dok":
                # Dokumen terima: TIDAK = merah, YA = hijau
                if v == "TIDAK":
                    row_cells[idx].fill = STATUS_FILLS["anomali"]
                    row_cells[idx].font = STATUS_FONT
                    row_cells[idx].alignment = STATUS_ALIGN
                elif v == "YA":
                    row_cells[idx].fill = STATUS_FILLS["normal"]
                    row_cells[idx].font = STATUS_FONT
                    row_cells[idx].alignment = STATUS_ALIGN
            else:
                # Status teks → severity
                if isinstance(v, str):
                    sev = severity_of_status(v)
                    if sev != "normal":
                        row_cells[idx].fill = STATUS_FILLS.get(sev, STATUS_FILLS["perhatian"])
                        row_cells[idx].font = STATUS_FONT
                        row_cells[idx].alignment = STATUS_ALIGN

    # Auto-width
    for c_idx, (col_name, w) in enumerate(zip(headers, widths), start=1):
        max_len = len(str(col_name))
        # Sample 100 rows max
        sample = [body_rows[i][c_idx - 1] for i in range(min(100, len(body_rows))) if body_rows[i][c_idx - 1] is not None]
        if sample:
            sample_max = max(len(str(s)) for s in sample)
            max_len = max(max_len, min(sample_max, 60))
        ws.column_dimensions[get_column_letter(c_idx)].width = min(max_len + 2, w)

    ws.row_dimensions[1].height = 28
    if freeze:
        ws.freeze_panes = freeze

    last_col = get_column_letter(len(headers))
    ws.auto_filter.ref = f"A1:{last_col}{len(body_rows) + 1}"
