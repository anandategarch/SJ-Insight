"""sheets.py — generator 8 sheet view yang ada di file asli '04.QUERRY SJ.xlsb'.

File asli user punya 17 sheet. 8 di antaranya BISA di-generate dari DataFrame
pipeline Python (5 sudah ada: Detail, Anomali, Ringkasan, Rekap Count, Log).
Modul ini generate 8 sheet tambahan yang mirror struktur sheet di xlsb:

  1. UPLOAD        — view gabungan Kirim×Terima (mirror sheet UPLOAD di xlsb)
  2. TERIMA_PIVOT  — pivot Kirim dgn Status Anomali (mirror sheet TERIMA)
  3. MANUAL        — filter Kirim subset (mirror sheet MANUAL)
  4. KIRIM_PIVOT   — pivot Kirim mentah (mirror sheet KIRIM)
  5. DOUBLE_TERIMA — filter Count>1 Concet, view Terima (mirror DOUBLE TERIMA)
  6. DOUBLE_KIRIM  — filter Count>1 Concet, view Kirim (mirror DOUBLE KIRIM)
  7. TERIMA_FLAT   — flat table = df_kirim (mirror 4215_TERIMA)
  8. KIRIM_FLAT    — flat table subset (mirror 4215_KIRIM)

Sheet yang TIDAK bisa di-generate (pure user input/master, load-only):
  - DOWNLOAD SJ, CEK MANUAL, DATABASE MANUAL, BANTU MANUAL, FILTER MANUAL
  - DATABASE, SATUAN, DATABASE SATUAN
"""
from __future__ import annotations

from typing import Any

import pandas as pd

from .utils import fmt_date_id as _fmt_date


# ============================================================
# SHEET 1: UPLOAD — view gabungan Kirim×Terima
# ============================================================
# Mirror sheet UPLOAD di xlsb (1487 rows). Header asli:
#   Tanggal [Kirim] | Tanggal [Terima] | NO. IT KIRIM | NO. IT TERIMA |
#   Gudang Pengirim | Extract [Gudang Tujuan] | Jenis | Dokumen2 |
#   [Selisih Hari] | Nm. File | ... | DOWNLOAD SJ

UPLOAD_COLS = [
    "Tanggal [Kirim]",
    "Tanggal [Terima]",
    "NO. IT KIRIM",
    "NO. IT TERIMA",
    "Gudang Pengirim",
    "Extract [Gudang Tujuan]",
    "Jenis",
    "Dokumen2",
    "[Selisih Hari]",
    "KTG [RESTO/DC]",
    "Status Pengiriman",
]


def generate_upload_view(df_merged: pd.DataFrame) -> pd.DataFrame:
    """Generate sheet UPLOAD — view gabungan Kirim×Terima.

    Source: df_merged (hasil stage Merge). Ambil kolom utama + rename.
    """
    if df_merged is None or df_merged.empty:
        return pd.DataFrame(columns=UPLOAD_COLS)

    rows: list[dict[str, Any]] = []
    for _, r in df_merged.iterrows():
        # Dokumen2: YA bila ada dokumen terima, TIDAK bila tidak ada
        has_dok = bool(r.get("Memiliki Dokumen [Terima]", False))
        rows.append({
            "Tanggal [Kirim]":         _fmt_date(r.get("Tanggal")),
            "Tanggal [Terima]":        _fmt_date(r.get("tTanggalTerima")) if r.get("tTanggalTerima") else None,
            "NO. IT KIRIM":            r.get("No. Pemindahan #"),
            "NO. IT TERIMA":           r.get("tNomorTerima"),
            "Gudang Pengirim":         r.get("Gudang Pengirim (Kode)"),
            "Extract [Gudang Tujuan]": r.get("Gudang Tujuan (Kode)"),
            "Jenis":                   r.get("JENIS (Kategori)"),
            "Dokumen2":                "YA" if has_dok else "TIDAK",
            "[Selisih Hari]":          r.get("Selisih Hari"),
            "KTG [RESTO/DC]":          r.get("KTG [RESTO/DC]"),
            "Status Pengiriman":       r.get("tStatusPengiriman"),
        })
    return pd.DataFrame(rows, columns=UPLOAD_COLS)


# ============================================================
# SHEET 2: TERIMA (PIVOT) — Kirim dgn Status Anomali
# ============================================================
# Mirror sheet TERIMA di xlsb (336 rows). Header:
#   Tanggal | No. Pemindahan # | Gudang Pengirim | Gudang Tujuan |
#   Tgl/Jam Pembuatan | Nama Barang | Satuan | Status Anomali Satuan |
#   Status Anomali QTY | Keterangan Pemindahan Barang |
#   Status Anomali QTY Terkecil 2 | [QTY] | Sum of QTY Terkecil 2

TERIMA_PIVOT_COLS = [
    "Tanggal",
    "No. Pemindahan #",
    "Gudang Pengirim",
    "Gudang Tujuan",
    "Tgl/Jam Pembuatan",
    "Nama Barang",
    "Satuan",
    "Status Anomali Satuan",
    "Status Anomali QTY",
    "Keterangan Pemindahan Barang",
    "Status Anomali QTY Terkecil 2",
    "[QTY]",
    "Sum of QTY Terkecil 2",
]


def generate_terima_pivot(df_kirim: pd.DataFrame) -> pd.DataFrame:
    """Generate sheet TERIMA (pivot view dari df_kirim).

    Filter: hanya baris dgn Status Pemindahan != null (mirror pivot filter).
    """
    if df_kirim is None or df_kirim.empty:
        return pd.DataFrame(columns=TERIMA_PIVOT_COLS)

    rows: list[dict[str, Any]] = []
    for _, r in df_kirim.iterrows():
        status = r.get("Status Pemindahan Barang")
        if not status:
            continue
        rows.append({
            "Tanggal":                          _fmt_date(r.get("Tanggal")),
            "No. Pemindahan #":                 r.get("No. Pemindahan #"),
            "Gudang Pengirim":                  r.get("Gudang Pengirim (Kode)"),
            "Gudang Tujuan":                    r.get("Gudang Tujuan (Kode)"),
            "Tgl/Jam Pembuatan":                r.get("Tgl/Jam Pembuatan"),
            "Nama Barang":                      r.get("Nama Barang"),
            "Satuan":                           r.get("Satuan"),
            "Status Anomali Satuan":             r.get("Status Anomali Satuan"),
            "Status Anomali QTY":                r.get("Status Anomali QTY"),
            "Keterangan Pemindahan Barang":      "",  # tidak ada di pipeline
            "Status Anomali QTY Terkecil 2":     r.get("Status Anomali QTY Terkecil 2"),
            "[QTY]":                             r.get("Kuantitas"),
            "Sum of QTY Terkecil 2":             r.get("QTY Terkecil 2"),
        })
    return pd.DataFrame(rows, columns=TERIMA_PIVOT_COLS)


# ============================================================
# SHEET 3: MANUAL — filter Kirim subset
# ============================================================
# Mirror sheet MANUAL di xlsb (676 rows). Header:
#   Tanggal | No. Pemindahan # | Gudang Pengirim | Gudang Tujuan |
#   Tgl/Jam Pembuatan | Nama Barang | Status Pemindahan Barang | Satuan |
#   Keterangan Pemindahan Barang

MANUAL_COLS = [
    "Tanggal",
    "No. Pemindahan #",
    "Gudang Pengirim",
    "Gudang Tujuan",
    "Tgl/Jam Pembuatan",
    "Nama Barang",
    "Status Pemindahan Barang",
    "Satuan",
    "Keterangan Pemindahan Barang",
]


def generate_manual_view(df_kirim: pd.DataFrame) -> pd.DataFrame:
    """Generate sheet MANUAL — subset Kirim tanpa kolom analisa."""
    if df_kirim is None or df_kirim.empty:
        return pd.DataFrame(columns=MANUAL_COLS)

    rows: list[dict[str, Any]] = []
    for _, r in df_kirim.iterrows():
        rows.append({
            "Tanggal":                     _fmt_date(r.get("Tanggal")),
            "No. Pemindahan #":            r.get("No. Pemindahan #"),
            "Gudang Pengirim":             r.get("Gudang Pengirim (Kode)"),
            "Gudang Tujuan":               r.get("Gudang Tujuan (Kode)"),
            "Tgl/Jam Pembuatan":           r.get("Tgl/Jam Pembuatan"),
            "Nama Barang":                 r.get("Nama Barang"),
            "Status Pemindahan Barang":    r.get("Status Pemindahan Barang"),
            "Satuan":                      r.get("Satuan"),
            "Keterangan Pemindahan Barang": "",
        })
    return pd.DataFrame(rows, columns=MANUAL_COLS)


# ============================================================
# SHEET 4: KIRIM (PIVOT) — pivot Kirim mentah
# ============================================================
# Mirror sheet KIRIM di xlsb (62 rows). Header:
#   Tanggal | No. Pemindahan # | Gudang Pengirim | Gudang Tujuan |
#   Nama Barang | Tgl/Jam Pembuatan | Status Pemindahan Barang | Satuan |
#   Keterangan Pemindahan Barang | QTY | QTY Terkecil

KIRIM_PIVOT_COLS = [
    "Tanggal",
    "No. Pemindahan #",
    "Gudang Pengirim",
    "Gudang Tujuan",
    "Nama Barang",
    "Tgl/Jam Pembuatan",
    "Status Pemindahan Barang",
    "Satuan",
    "Keterangan Pemindahan Barang",
    "QTY",
    "QTY Terkecil",
]


def generate_kirim_pivot(df_kirim: pd.DataFrame) -> pd.DataFrame:
    """Generate sheet KIRIM — pivot mentah dari df_kirim."""
    if df_kirim is None or df_kirim.empty:
        return pd.DataFrame(columns=KIRIM_PIVOT_COLS)

    rows: list[dict[str, Any]] = []
    for _, r in df_kirim.iterrows():
        rows.append({
            "Tanggal":                     _fmt_date(r.get("Tanggal")),
            "No. Pemindahan #":            r.get("No. Pemindahan #"),
            "Gudang Pengirim":             r.get("Gudang Pengirim (Kode)"),
            "Gudang Tujuan":               r.get("Gudang Tujuan (Kode)"),
            "Nama Barang":                 r.get("Nama Barang"),
            "Tgl/Jam Pembuatan":           r.get("Tgl/Jam Pembuatan"),
            "Status Pemindahan Barang":    r.get("Status Pemindahan Barang"),
            "Satuan":                      r.get("Satuan"),
            "Keterangan Pemindahan Barang": "",
            "QTY":                         r.get("Kuantitas"),
            "QTY Terkecil":                r.get("QTY Terkecil"),
        })
    return pd.DataFrame(rows, columns=KIRIM_PIVOT_COLS)


# ============================================================
# SHEET 5 & 6: DOUBLE TERIMA / DOUBLE KIRIM — Count > 1
# ============================================================
# Mirror sheet DOUBLE TERIMA (43 rows) & DOUBLE KIRIM (46 rows).
# Pivot filter 4215TerimaCount.Count > 1.
# Header DOUBLE TERIMA:
#   Gudang Pengirim | Tanggal | Gudang Tujuan | Tgl/Jam Pembuatan |
#   Status Pemindahan Barang | No. Pemindahan # | Nama Barang | Satuan |
#   Sum of Kuantitas
# Header DOUBLE KIRIM: + Sum of Kts Diterima (kolom 10)

DOUBLE_TERIMA_COLS = [
    "Gudang Pengirim",
    "Tanggal",
    "Gudang Tujuan",
    "Tgl/Jam Pembuatan",
    "Status Pemindahan Barang",
    "No. Pemindahan #",
    "Nama Barang",
    "Satuan",
    "Sum of Kuantitas",
]

DOUBLE_KIRIM_COLS = DOUBLE_TERIMA_COLS + ["Sum of Kts Diterima"]


def generate_double_terima(
    df_kirim: pd.DataFrame,
    df_count: pd.DataFrame,
) -> pd.DataFrame:
    """Generate sheet DOUBLE TERIMA — baris dengan Count Concet > 1.

    Mirror M code 4215TerimaCount + pivot filter Count > 1.
    """
    if df_kirim is None or df_kirim.empty:
        return pd.DataFrame(columns=DOUBLE_TERIMA_COLS)
    if df_count is None or df_count.empty:
        return pd.DataFrame(columns=DOUBLE_TERIMA_COLS)

    # Concet dengan Count > 1
    dup_concets = set(df_count.loc[df_count["Count"] > 1, "Concet"])
    if not dup_concets:
        return pd.DataFrame(columns=DOUBLE_TERIMA_COLS)

    # Filter df_kirim baris yang Concet-nya duplikat
    df_dup = df_kirim[df_kirim["Concet"].isin(dup_concets)].copy()
    rows: list[dict[str, Any]] = []
    for _, r in df_dup.iterrows():
        rows.append({
            "Gudang Pengirim":            r.get("Gudang Pengirim (Kode)"),
            "Tanggal":                    _fmt_date(r.get("Tanggal")),
            "Gudang Tujuan":              r.get("Gudang Tujuan (Kode)"),
            "Tgl/Jam Pembuatan":          r.get("Tgl/Jam Pembuatan"),
            "Status Pemindahan Barang":   r.get("Status Pemindahan Barang"),
            "No. Pemindahan #":           r.get("No. Pemindahan #"),
            "Nama Barang":                r.get("Nama Barang"),
            "Satuan":                     r.get("Satuan"),
            "Sum of Kuantitas":           r.get("Kuantitas"),
        })
    return pd.DataFrame(rows, columns=DOUBLE_TERIMA_COLS)


def generate_double_kirim(
    df_kirim: pd.DataFrame,
    df_count: pd.DataFrame,
) -> pd.DataFrame:
    """Generate sheet DOUBLE KIRIM — baris dengan Count Concet > 1 (+QTY Terkecil)."""
    if df_kirim is None or df_kirim.empty:
        return pd.DataFrame(columns=DOUBLE_KIRIM_COLS)
    if df_count is None or df_count.empty:
        return pd.DataFrame(columns=DOUBLE_KIRIM_COLS)

    dup_concets = set(df_count.loc[df_count["Count"] > 1, "Concet"])
    if not dup_concets:
        return pd.DataFrame(columns=DOUBLE_KIRIM_COLS)

    df_dup = df_kirim[df_kirim["Concet"].isin(dup_concets)].copy()
    rows: list[dict[str, Any]] = []
    for _, r in df_dup.iterrows():
        rows.append({
            "Gudang Pengirim":            r.get("Gudang Pengirim (Kode)"),
            "Tanggal":                    _fmt_date(r.get("Tanggal")),
            "Gudang Tujuan":              r.get("Gudang Tujuan (Kode)"),
            "Tgl/Jam Pembuatan":          r.get("Tgl/Jam Pembuatan"),
            "Status Pemindahan Barang":   r.get("Status Pemindahan Barang"),
            "No. Pemindahan #":           r.get("No. Pemindahan #"),
            "Nama Barang":                r.get("Nama Barang"),
            "Satuan":                     r.get("Satuan"),
            "Sum of Kuantitas":           r.get("Kuantitas"),
            "Sum of Kts Diterima":        r.get("QTY Terkecil"),
        })
    return pd.DataFrame(rows, columns=DOUBLE_KIRIM_COLS)


# ============================================================
# SHEET 7 & 8: TERIMA FLAT / KIRIM FLAT — full table
# ============================================================
# Mirror sheet 4215_TERIMA & 4215_KIRIM (kosong di xlsb, output target M query).
# Ini = df_kirim full dengan semua kolom derived.

def generate_terima_flat(df_kirim: pd.DataFrame) -> pd.DataFrame:
    """Generate sheet 4215_TERIMA — flat table = df_kirim (32 cols)."""
    if df_kirim is None or df_kirim.empty:
        return pd.DataFrame()
    return df_kirim.copy()


def generate_kirim_flat(df_kirim: pd.DataFrame) -> pd.DataFrame:
    """Generate sheet 4215_KIRIM — flat subset (Kolom utama Kirim)."""
    if df_kirim is None or df_kirim.empty:
        return pd.DataFrame()
    cols = [c for c in [
        "No. Pemindahan #", "Tanggal", "Tgl/Jam Pembuatan",
        "Gudang Pengirim (Kode)", "Gudang Tujuan (Kode)",
        "Nama Barang", "Satuan", "QTY Terkecil", "Kuantitas",
        "Status Pemindahan Barang",
    ] if c in df_kirim.columns]
    return df_kirim[cols].copy()


# ============================================================
# DAFTAR SEMUA SHEET VIEW
# ============================================================
SHEET_VIEWS = {
    "Upload":         (UPLOAD_COLS,        generate_upload_view),
    "Terima Pivot":   (TERIMA_PIVOT_COLS,  generate_terima_pivot),
    "Manual":         (MANUAL_COLS,        generate_manual_view),
    "Kirim Pivot":    (KIRIM_PIVOT_COLS,   generate_kirim_pivot),
    "Double Terima":  (DOUBLE_TERIMA_COLS, generate_double_terima),
    "Double Kirim":   (DOUBLE_KIRIM_COLS,  generate_double_kirim),
    "4215 Terima":    (None,               generate_terima_flat),
    "4215 Kirim":     (None,               generate_kirim_flat),
}


def generate_all_views(
    df_kirim: pd.DataFrame,
    df_merged: pd.DataFrame,
    df_count: pd.DataFrame,
) -> dict[str, pd.DataFrame]:
    """Generate SEMUA sheet view dari DataFrame pipeline.

    Return: dict {sheet_name: DataFrame}
    """
    views: dict[str, pd.DataFrame] = {}
    views["Upload"] = generate_upload_view(df_merged)
    views["Terima Pivot"] = generate_terima_pivot(df_kirim)
    views["Manual"] = generate_manual_view(df_kirim)
    views["Kirim Pivot"] = generate_kirim_pivot(df_kirim)
    views["Double Terima"] = generate_double_terima(df_kirim, df_count)
    views["Double Kirim"] = generate_double_kirim(df_kirim, df_count)
    views["4215 Terima"] = generate_terima_flat(df_kirim)
    views["4215 Kirim"] = generate_kirim_flat(df_kirim)
    return views
