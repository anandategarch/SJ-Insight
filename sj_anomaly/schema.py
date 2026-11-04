"""schema.py — validasi kontrak header 3 sumber SJ (anti degradasi senyap).

Mengikuti pola ``waste_anomaly.schema`` (SchemaValidationError + pesan
actionable dengan near-miss suggestions).

Hard error (run gagal):
  - Kolom wajib hilang di salah satu sumber
  - Header duplikat pada kolom wajib
"""
from __future__ import annotations

import pandas as pd

from .loader import _find_column


class SchemaValidationError(ValueError):
    """Input tidak memenuhi kontrak skema — pesan sudah user-friendly."""


# Kolom wajib (mirror M code step 1-3 / Merge query)
KIRIM_REQUIRED = ["Nama Barang", "Satuan", "Status Pemindahan Barang"]
TERIMA_REQUIRED = ["Nomor #Kirim"]
KATEGORI_REQUIRED = ["NAMA BARANG", "JENIS"]


def _suggest_near_miss(missing: str, columns: list) -> str | None:
    """Cari kandidat kolom yang mirip (case/strip beda)."""
    norm = lambda s: str(s).strip().upper().replace("  ", " ")
    target = norm(missing)
    for col in columns:
        if norm(col) == target and str(col) != missing:
            return str(col)
    return None


def _check_duplicates(columns: list[str]) -> list[str]:
    """Return list kolom duplikat."""
    seen: dict[str, int] = {}
    dups: list[str] = []
    for col in columns:
        seen[col] = seen.get(col, 0) + 1
        if seen[col] == 2:
            dups.append(col)
    return dups


def validate_kirim_schema(df: pd.DataFrame, config: dict, logger=None) -> list[str]:
    """Validasi sheet Kirim (42.15). Raise SchemaValidationError bila wajib hilang."""
    warnings: list[str] = []
    columns = [str(c) for c in df.columns]

    dups = _check_duplicates(columns)
    if dups:
        raise SchemaValidationError(
            f"Header DUPLIKAT di sheet Kirim (42.15): {dups}. "
            f"pandas membaca kolom kedua sebagai '{dups[0]}.1' — "
            f"hapus/rename kolom duplikat di Excel lalu save ulang."
        )

    kirim_cfg = config.get("kirim_columns", {})
    all_required_keys = ["nama", "satuan", "status"]
    missing: list[str] = []
    hints: list[str] = []
    found_map: dict[str, str] = {}
    for key in all_required_keys:
        cand = kirim_cfg.get(key, [])
        found = _find_column(columns, cand)
        if found:
            found_map[key] = found
        else:
            missing.append(f"{key} (kandidat: {cand[:2]})")
            for c in cand:
                near = _suggest_near_miss(c, columns)
                if near:
                    hints.append(f"'{c}' tidak ada, tapi ada '{near}' — rename ke '{c}'")
                    break

    if missing:
        # Sertakan info header row + sample headers utk diagnosis
        hdr_row = df.attrs.get("header_row", "?")
        sample_cols = columns[:12]
        raise SchemaValidationError(
            f"Kolom WAJIB hilang dari sheet Kirim (42.15): {missing}. "
            f"Header terdeteksi di baris {hdr_row}. "
            f"Kolom tersedia ({len(columns)} total): {sample_cols}... "
            f"Pastikan file adalah report '42.15 Pemindahan Barang (KIRIM)' "
            f"dengan header 'Nama Barang' / 'Satuan' / 'Status Pemindahan Barang'. "
            + (f" Petunjuk: {'; '.join(hints)}." if hints else "")
        )

    return warnings


def validate_terima_schema(df: pd.DataFrame, config: dict, logger=None) -> list[str]:
    """Validasi sheet Terima (42.05). Raise bila 'Nomor #Kirim' hilang."""
    warnings: list[str] = []
    columns = [str(c) for c in df.columns]

    dups = _check_duplicates(columns)
    if dups:
        raise SchemaValidationError(
            f"Header DUPLIKAT di sheet Terima (42.05): {dups}. "
            f"Hapus/rename kolom duplikat di Excel."
        )

    terima_cfg = config.get("terima_columns", {})
    found = _find_column(columns, terima_cfg.get("nomor_kirim", ["Nomor #Kirim"]))
    if not found:
        near = _suggest_near_miss("Nomor #Kirim", columns)
        hint = f" Petunjuk: ada '{near}' — rename ke 'Nomor #Kirim'." if near else ""
        raise SchemaValidationError(
            f"Kolom 'Nomor #Kirim' tidak ditemukan di sheet Terima (42.05). "
            f"Pastikan file adalah report '42.05 Rincian Pemindahan Barang' "
            f"dengan header di baris ke-4.{hint}"
        )
    return warnings


def validate_kategori_schema(df: pd.DataFrame, config: dict, logger=None) -> list[str]:
    """Validasi sheet DBS Kategori. Raise bila NAMA BARANG / JENIS hilang."""
    warnings: list[str] = []
    columns = [str(c) for c in df.columns]

    dups = _check_duplicates(columns)
    if dups:
        raise SchemaValidationError(
            f"Header DUPLIKAT di sheet DBS Kategori: {dups}. "
            f"Hapus/rename kolom duplikat di Excel."
        )

    kat_cfg = config.get("kategori_columns", {})
    for key in ["nama", "jenis"]:
        cand = kat_cfg.get(key, [])
        found = _find_column(columns, cand)
        if not found:
            near = _suggest_near_miss(cand[0] if cand else "", columns)
            hint = f" Petunjuk: ada '{near}' — rename ke '{cand[0]}'." if near else ""
            raise SchemaValidationError(
                f"Kolom '{cand[0] if cand else key}' tidak ditemukan di sheet "
                f"DBS Kategori (Sheet1). Header tersedia: {columns[:8]}.{hint}"
            )
    return warnings
