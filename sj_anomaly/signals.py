"""signals.py — ringkasan KPI + chart data utk dashboard UI.

Di-port pola ``waste_anomaly.insights`` / ``charts`` (sebagian) —
membangun ringkasan agregat dari DataFrame hasil pipeline utk dikonsumsi
tab Dashboard / Scorecard.
"""
from __future__ import annotations

import pandas as pd

from .utils import severity_of_status


def build_summary(df_kirim: pd.DataFrame, df_merged: pd.DataFrame, df_count: pd.DataFrame, config: dict) -> dict:
    """Bangun ringkasan KPI utk dashboard (mirror buildSummary pipeline).

    Hitung:
      - total_transaksi, total_barang, total_gudang_tujuan
      - anomali_satuan, anomali_qty, anomali_konversi, perlu_cek, data_terbatas, normal
      - kirim_tanpa_terima, terima_lambat (selisih_hari >= threshold)
      - duplikat_concet (count > 1)
      - rata2_selisih_hari
      - tanggal_min, tanggal_max
    """
    th_terima_lambat = int(config.get("thresholds", {}).get("TERIMA_LAMBAT_HARI", 2))

    anomali_satuan = int((df_kirim["Status Anomali Satuan"] != "NORMAL").sum())
    anomali_qty = int((df_kirim["Status Anomali QTY"] != "NORMAL").sum())

    # Status konversi: ANOMALI KONVERSI / PERLU CEK / DATA TERBATAS / KONVERSI TIDAK STABIL
    status_konv = df_kirim["Status Anomali QTY Terkecil 2"]
    anomali_konversi = int((status_konv == "ANOMALI KONVERSI").sum())
    perlu_cek = int(status_konv.astype(str).str.startswith("PERLU CEK").sum())
    data_terbatas = int(
        ((status_konv == "DATA TERBATAS") | (status_konv == "KONVERSI TIDAK STABIL")).sum()
    )

    # Baris normal bersih (ketiga status NORMAL)
    normal_mask = (
        (df_kirim["Status Anomali Satuan"] == "NORMAL")
        & (df_kirim["Status Anomali QTY"] == "NORMAL")
        & (df_kirim["Status Anomali QTY Terkecil 2"] == "NORMAL")
    )
    normal = int(normal_mask.sum())

    # Merge-side stats
    if "Memiliki Dokumen [Terima]" in df_merged.columns:
        kirim_tanpa_terima = int((~df_merged["Memiliki Dokumen [Terima]"]).sum())
    else:
        kirim_tanpa_terima = 0
    if "Selisih Hari" in df_merged.columns:
        selisih_valid = df_merged["Selisih Hari"].dropna()
        selisih_valid = selisih_valid[selisih_valid.apply(lambda x: isinstance(x, (int, float)))]
        terima_lambat = int((selisih_valid >= th_terima_lambat).sum())
        rata2_selisih = float(selisih_valid.mean()) if len(selisih_valid) else None
    else:
        terima_lambat = 0
        rata2_selisih = None
    rata2_selisih = round(rata2_selisih, 2) if rata2_selisih is not None else None

    duplikat_concet = int((df_count["Count"] > 1).sum()) if len(df_count) else 0

    tanggal_all = df_kirim["Tanggal"].dropna().sort_values().tolist()
    tanggal_min = tanggal_all[0] if tanggal_all else None
    tanggal_max = tanggal_all[-1] if tanggal_all else None

    return {
        "total_transaksi":       int(len(df_kirim)),
        "total_barang":          int(df_kirim["Nama Barang"].nunique()),
        "total_gudang_tujuan":   int(df_kirim["Gudang Tujuan (Kode)"].nunique()),
        "anomali_satuan":        anomali_satuan,
        "anomali_qty":           anomali_qty,
        "anomali_konversi":      anomali_konversi,
        "perlu_cek":             perlu_cek,
        "data_terbatas":         data_terbatas,
        "normal":                normal,
        "kirim_tanpa_terima":    kirim_tanpa_terima,
        "terima_lambat":         terima_lambat,
        "duplikat_concet":       duplikat_concet,
        "rata2_selisih_hari":    rata2_selisih,
        "tanggal_min":           tanggal_min,
        "tanggal_max":           tanggal_max,
    }


def build_charts_data(df_kirim: pd.DataFrame, df_merged: pd.DataFrame) -> dict:
    """Bangun data chart utk dashboard (per gudang / per barang / breakdown)."""
    # Index merged by No. Pemindahan # utk lookup cepat
    merged_by_no: dict[str, dict] = {}
    if "No. Pemindahan #" in df_merged.columns:
        for _, m in df_merged.iterrows():
            no = m["No. Pemindahan #"]
            if no and no not in merged_by_no:
                merged_by_no[no] = m.to_dict()

    # --- Per gudang tujuan (top 12 by volume) ---
    by_gudang: dict[str, dict] = {}
    for _, r in df_kirim.iterrows():
        g = r["Gudang Tujuan (Kode)"] or "(tanpa nama)"
        if g not in by_gudang:
            by_gudang[g] = {"gudang": g, "total": 0, "anomali": 0}
        by_gudang[g]["total"] += 1
        no = r["No. Pemindahan #"]
        m = merged_by_no.get(no)
        is_anomali = (
            r["Status Anomali Satuan"] != "NORMAL"
            or r["Status Anomali QTY"] != "NORMAL"
            or severity_of_status(r["Status Anomali QTY Terkecil 2"]) == "anomali"
            or (m is not None and not m.get("Memiliki Dokumen [Terima]", True))
        )
        if is_anomali:
            by_gudang[g]["anomali"] += 1
    per_gudang = sorted(by_gudang.values(), key=lambda x: x["total"], reverse=True)[:12]

    # --- Per barang (top 12 by anomali lalu volume) ---
    by_barang: dict[str, dict] = {}
    for _, r in df_kirim.iterrows():
        g = r["Nama Barang"] or "(tanpa nama)"
        if g not in by_barang:
            by_barang[g] = {"nama_barang": g, "total": 0, "anomali": 0}
        by_barang[g]["total"] += 1
        no = r["No. Pemindahan #"]
        m = merged_by_no.get(no)
        if (
            r["Status Anomali Satuan"] != "NORMAL"
            or r["Status Anomali QTY"] != "NORMAL"
            or severity_of_status(r["Status Anomali QTY Terkecil 2"]) == "anomali"
            or (m is not None and not m.get("Memiliki Dokumen [Terima]", True))
        ):
            by_barang[g]["anomali"] += 1
    per_barang = sorted(
        by_barang.values(),
        key=lambda x: (x["anomali"], x["total"]),
        reverse=True,
    )[:12]

    # --- Donut: klasifikasi baris berdasar severity terburuk ---
    n_anomali = n_perlu_cek = n_perhatian = n_normal = 0
    for _, r in df_kirim.iterrows():
        no = r["No. Pemindahan #"]
        m = merged_by_no.get(no)
        severities = [
            severity_of_status(r["Status Anomali Satuan"]),
            severity_of_status(r["Status Anomali QTY"]),
            severity_of_status(r["Status Anomali QTY Terkecil 2"]),
        ]
        if m is not None and not m.get("Memiliki Dokumen [Terima]", True):
            severities.append("anomali")
        if "anomali" in severities:
            n_anomali += 1
        elif "perlu-cek" in severities:
            n_perlu_cek += 1
        elif "perhatian" in severities:
            n_perhatian += 1
        else:
            n_normal += 1
    breakdown = [
        {"name": "Anomali", "value": n_anomali},
        {"name": "Perlu Cek", "value": n_perlu_cek},
        {"name": "Perhatian", "value": n_perhatian},
        {"name": "Normal", "value": n_normal},
    ]
    breakdown = [d for d in breakdown if d["value"] > 0]

    return {
        "per_gudang": per_gudang,
        "per_barang": per_barang,
        "anomaly_breakdown": breakdown,
    }
