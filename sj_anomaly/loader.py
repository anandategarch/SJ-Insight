"""loader.py — pemuat 3 sumber Excel untuk pipeline SJ Insight.

Mengikuti pola ``waste_anomaly.loader`` (calamine → openpyxl fallback,
_find_column multi-kandidat, pesan error jelas) tetapi disesuaikan utk 3
sumber SJ:
  - 42.15 Pemindahan Barang (KIRIM) — sheet "#42.15. Pemindahan Barang (KIRI"
  - 42.05 Rincian Pemindahan (TERIMA) — sheet "#42.05. Rincian Pemindahan Bara"
  - DBS Kategori — sheet "Sheet1"

M code sumber (mapping header):
  - ``Transform File (2)`` : Excel.Workbook → sheet "#42.05. Rincian Pemindahan Bara"
  - ``Transform File (3)`` : sheet "#42.15. Pemindahan Barang (KIRI"
  - ``Transform File (6)`` : Sheet1 (DBS Kategori)

Sheet 42.15 & 42.05 punya 3 baris judul report ACCURATE di atas header
(itulah kenapa M code memakai ``Table.Skip(3)``). DBS Kategori punya
header di baris pertama (skip=0).
"""
from __future__ import annotations

import logging
import math
import os
import re
from pathlib import Path
from typing import Any

import numpy as np
import pandas as pd

# Engine selection: calamine (cepat) → fallback openpyxl (default pandas)
try:
    import python_calamine  # noqa: F401
    _EXCEL_ENGINE = "calamine"
except ImportError:
    _EXCEL_ENGINE = None  # pandas pakai default (openpyxl)

from openpyxl import load_workbook

_log = logging.getLogger(__name__)


# ============================================================
# FLOAT NORMALIZATION (mirror Waste F7 — anti FP dust)
# ============================================================

def _float15(x: Any) -> Any:
    """Bulatkan satu float ke 15 digit signifikan (semantik penyimpanan Excel)."""
    try:
        if pd.isna(x):
            return x
        x_f = float(x)
        if math.isinf(x_f):
            return x
        return float(f"{x_f:.15g}")
    except (TypeError, ValueError):
        return x


def _normalize_float_precision(df: pd.DataFrame) -> pd.DataFrame:
    """Normalisasi kolom float ke 15 digit signifikan (in-place safe)."""
    for col in df.columns:
        s = df[col]
        if not pd.api.types.is_float_dtype(s):
            continue
        arr = s.to_numpy()
        finite = np.isfinite(arr)
        if not finite.any():
            continue
        vals = arr[finite]
        if (np.all(np.abs(vals) < 1e15)
                and np.all(vals == np.trunc(vals))):
            continue  # bulat ≤ 15 digit → "%.15g" identik → no-op
        arr[finite] = np.array([_float15(v) for v in vals])
        df[col] = arr
    return df


def _read_excel_fast(path, sheet_name, **kwargs):
    """Wrapper pd.read_excel dengan engine optimal (calamine jika ada)."""
    if _EXCEL_ENGINE:
        df = pd.read_excel(path, sheet_name=sheet_name, engine=_EXCEL_ENGINE, **kwargs)
    else:
        df = pd.read_excel(path, sheet_name=sheet_name, **kwargs)
    return _normalize_float_precision(df)


def _guard_xls(workbook_path: str | os.PathLike) -> None:
    """Tolak file .xls (Excel lama, format BIFF) dengan pesan jelas."""
    if Path(workbook_path).suffix.lower() == ".xls":
        raise ValueError("File .xls (Excel lama) tidak didukung — save as .xlsx dulu")


# ============================================================
# DETEKSI SHEET (token-based, mirror Waste find_sheet)
# ============================================================

def list_sheets(workbook_path: str | os.PathLike) -> list[str]:
    """List semua sheet di workbook (openpyxl, read-only)."""
    _guard_xls(workbook_path)
    wb = load_workbook(workbook_path, read_only=True, data_only=True)
    sheets = list(wb.sheetnames)
    wb.close()
    return sheets


def _sheet_tokens(name: str) -> frozenset[str]:
    """Token nama sheet: lowercase, pecah pada non-alfanumerik."""
    return frozenset(
        t for t in re.split(r"[^0-9a-zA-Z]+", str(name).strip().lower()) if t
    )


def find_sheet(names_candidates: list[str], available: list[str]) -> str | None:
    """Cari sheet yang cocok dari daftar kandidat (case-insensitive).

    Dua fase:
      1. EXACT (case-insensitive, strip).
      2. TOKEN: semua token kandidat harus ada sebagai TOKEN UTUH di nama
         sheet. Wajib unik per kandidat (anti ambigu).
    """
    avail = [str(s) for s in available]
    avail_norm = [s.strip().lower() for s in avail]
    # 1. Exact
    for cand in names_candidates:
        cl = str(cand).strip().lower()
        if cl and cl in avail_norm:
            return avail[avail_norm.index(cl)]
    # 2. Token containment
    sheet_tokens = [_sheet_tokens(s) for s in avail]
    for cand in names_candidates:
        cand_tokens = _sheet_tokens(str(cand))
        if not cand_tokens:
            continue
        matches = [i for i, toks in enumerate(sheet_tokens) if cand_tokens.issubset(toks)]
        if len(matches) == 1:
            return avail[matches[0]]
    return None


# ============================================================
# KEY NORMALIZATION (join Kirim×Terima via Nomor #Kirim)
# ============================================================

def _key_str(value: Any) -> str | None:
    """Normalisasi nilai sel jadi string key join.

    - None / NaN / NaT / pd.NA → None (bukan string "nan")
    - float integral → str(int)  (1049.0 → "1049")
    - nilai lain → str(value).strip()
    """
    if value is None:
        return None
    try:
        if pd.isna(value):
            return None
    except (TypeError, ValueError):
        pass
    if isinstance(value, float) and math.isfinite(value) and value.is_integer():
        return str(int(value))
    return str(value).strip()


# ============================================================
# NORMALISASI HEADER (mirror _normalize_columns Waste)
# ============================================================

def _normalize_columns(df: pd.DataFrame) -> pd.DataFrame:
    """Strip whitespace + hilangkan BOM char + cast float-integral header ke int."""
    new_cols = []
    for c in df.columns:
        if isinstance(c, str):
            s = c.replace("\ufeff", "").strip()
            new_cols.append(s)
        else:
            if isinstance(c, float) and math.isfinite(c) and c.is_integer():
                c = int(c)
            new_cols.append(str(c))
    df.columns = new_cols
    return df


def _find_column(available: list[str], candidates: list[str]) -> str | None:
    """Cari kolom dari daftar kandidat (exact → contains)."""
    avail_lower = [str(c).lower().strip() for c in available]
    for cand in candidates:
        cl = cand.lower().strip()
        if cl in avail_lower:
            return available[avail_lower.index(cl)]
    for cand in candidates:
        cl = cand.lower().strip()
        for i, s in enumerate(avail_lower):
            if cl and cl in s:
                return available[i]
    return None


# ============================================================
# UTILITAS PARSING TANGGAL & ANGKA (culture id-ID)
# ============================================================

def parse_date_value(v: Any, with_time: bool = False) -> str | None:
    """Parse nilai tanggal fleksibel → ISO 'YYYY-MM-DD' atau 'YYYY-MM-DD HH:MM:SS'.

    Mendukung:
      - datetime.datetime / datetime.date → ISO
      - Excel serial (number) → Date (Excel epoch 1899-12-30)
      - 'dd/mm/yyyy[ HH:MM:SS]' / 'dd-mm-yyyy' (format ACCURATE id-ID)
      - 'yyyy-mm-dd[ HH:MM:SS]' (ISO)
      - Fallback Date.parse (mis. '07 Nov 2025')
    """
    if v is None or v == "" or (isinstance(v, float) and math.isnan(v)):
        return None

    if isinstance(v, pd.Timestamp):
        v = v.to_pydatetime()

    import datetime as _dt
    if isinstance(v, _dt.datetime):
        return v.strftime("%Y-%m-%d %H:%M:%S") if with_time else v.strftime("%Y-%m-%d")
    if isinstance(v, _dt.date):
        return v.strftime("%Y-%m-%d")

    if isinstance(v, (int, float)) and not math.isnan(v) and isfinite(v):
        try:
            serial = int(v)
            if 1 < serial < 200000:  # reasonable Excel date range
                d = _dt.date(1899, 12, 30) + _dt.timedelta(days=serial)
                return d.strftime("%Y-%m-%d")
        except (ValueError, OverflowError):
            pass

    s = str(v).strip()
    if not s:
        return None

    # dd/mm/yyyy atau dd-mm-yyyy (± jam) — format ACCURATE id-ID
    m = re.match(
        r"^(\d{1,2})[/\-.](\d{1,2})[/\-.](\d{2,4})(?:[ T](\d{1,2}):(\d{2})(?::(\d{2}))?)?",
        s,
    )
    if m:
        dd, mm, yy, hh, mi, ss = m.groups()
        year = int(yy)
        if year < 100:
            year += 2000 if year < 50 else 1900
        try:
            d = _dt.datetime(
                year, int(mm), int(dd),
                int(hh or 0), int(mi or 0), int(ss or 0),
            )
            return d.strftime("%Y-%m-%d %H:%M:%S") if with_time else d.strftime("%Y-%m-%d")
        except ValueError:
            pass

    # yyyy-mm-dd (± jam)
    m = re.match(
        r"^(\d{4})-(\d{1,2})-(\d{1,2})(?:[ T](\d{1,2}):(\d{2})(?::(\d{2}))?)?",
        s,
    )
    if m:
        y, mm, dd, hh, mi, ss = m.groups()
        try:
            d = _dt.datetime(
                int(y), int(mm), int(dd),
                int(hh or 0), int(mi or 0), int(ss or 0),
            )
            return d.strftime("%Y-%m-%d %H:%M:%S") if with_time else d.strftime("%Y-%m-%d")
        except ValueError:
            pass

    # Fallback: Date.parse via pandas Timestamp
    try:
        ts = pd.to_datetime(s, errors="raise")
        return ts.strftime("%Y-%m-%d %H:%M:%S") if with_time else ts.strftime("%Y-%m-%d")
    except Exception:
        return None


def diff_days(a_iso: str | None, b_iso: str | None) -> int | None:
    """Selisih hari antar dua tanggal ISO (b - a), bulat."""
    if not a_iso or not b_iso:
        return None
    try:
        a = pd.Timestamp(a_iso[:10])
        b = pd.Timestamp(b_iso[:10])
        if pd.isna(a) or pd.isna(b):
            return None
        return int((b - a).days)
    except Exception:
        return None


def parse_number_value(v: Any) -> float | None:
    """Coerce angka fleksibel: number, '1.234,56' (id-ID), '1,234.56' (en)."""
    if v is None or v == "":
        return None
    if isinstance(v, (int, float)) and not (isinstance(v, float) and math.isnan(v)):
        return float(v)
    s = str(v).strip().replace(" ", "")
    if not s or s == "-":
        return None
    has_comma = "," in s
    has_dot = "." in s
    if has_comma and has_dot:
        # id-ID: '.' ribuan, ',' desimal → en: '.' ribuan, '.' desimal (tergantung posisi)
        if s.rfind(",") > s.rfind("."):
            s = s.replace(".", "").replace(",", ".")
        else:
            s = s.replace(",", "")
    elif has_comma:
        s = s.replace(",", ".")
    try:
        n = float(s)
        if math.isnan(n) or math.isinf(n):
            return None
        return n
    except ValueError:
        return None


def between_delimiters(text: Any, start: str = "(", end: str = ")") -> str:
    """Ambil teks di antara dua delimiter — mirror Text.BetweenDelimiters M code.

    Tanpa kurung → kembalikan apa adanya (trim).
    """
    s = str(text if text is not None else "")
    i = s.find(start)
    if i < 0:
        return s.strip()
    j = s.find(end, i + len(start))
    if j < 0:
        return s[i + len(start):].strip()
    return s[i + len(start):j].strip()


def text_trim(v: Any) -> str:
    """Trim teks aman (null-safe)."""
    return str(v if v is not None else "").strip()


# ============================================================
# LOAD WORKBOOK — skip 3 baris judul + promote headers
# ============================================================

def _promote_headers(matrix: list[list[Any]]) -> tuple[list[str], list[dict]]:
    """Promote baris pertama sebagai header; sisanya → list[dict]."""
    raw_headers = [
        (str(h).strip() if h is not None else "") or f"Column{i + 1}"
        for i, h in enumerate(matrix[0] if matrix else [])
    ]
    # Anti duplikat: "X", "X" → "X", "X.1"
    seen: dict[str, int] = {}
    headers = []
    for h in raw_headers:
        n = seen.get(h, 0)
        seen[h] = n + 1
        headers.append(h if n == 0 else f"{h}.{n}")

    rows: list[dict] = []
    for row in matrix[1:]:
        if all(c is None or c == "" for c in row):
            continue
        obj = {h: row[i] if i < len(row) else None for i, h in enumerate(headers)}
        rows.append(obj)
    return headers, rows


def load_sheet(
    workbook_path: str | os.PathLike,
    candidates: list[str],
    skip_rows: int = 3,
    fallback_first: bool = True,
) -> pd.DataFrame:
    """Load sheet → DataFrame (header = baris setelah ``skip_rows``).

    Mirror M code: ``Table.Skip(3)`` + ``Table.PromoteHeaders``.
    """
    _guard_xls(workbook_path)
    sheets = list_sheets(workbook_path)
    sheet = find_sheet(candidates, sheets)
    if sheet is None:
        if not fallback_first:
            raise ValueError(
                f"Tidak menemukan sheet. Kandidat: {candidates}. "
                f"Sheet tersedia: {sheets}."
            )
        sheet = sheets[0]
        _log.warning("Sheet kandidat tak ketemu — pakai sheet pertama: %s", sheet)

    # Baca sebagai matrix (header=1) supaya kita bisa skip + promote manual
    df_raw = _read_excel_fast(
        workbook_path, sheet_name=sheet, header=None, dtype=object
    )
    matrix = df_raw.values.tolist()
    if not matrix:
        raise ValueError(f"Sheet '{sheet}' kosong.")
    sliced = matrix[skip_rows:]
    if not sliced:
        raise ValueError(
            f"Sheet '{sheet}' hanya berisi {skip_rows} baris judul — tidak ada data."
        )
    headers, rows = _promote_headers(sliced)
    df = pd.DataFrame(rows, columns=headers)
    return _normalize_columns(df)


# ============================================================
# LOAD 1: KIRIM 42.15
# ============================================================

def load_kirim(
    workbook_path: str | os.PathLike,
    config: dict,
    logger=None,
) -> pd.DataFrame:
    """Load sheet Kirim (42.15). Sheet '#42.15. Pemindahan Barang (KIRI'.

    Return DataFrame mentah dengan kolom asli (belum di-cast) — pipeline
    yang menangani tipe. Hanya buang baris judul (skip=3) + promote header.
    """
    candidates = config.get("input_sheets", {}).get(
        "kirim", ["#42.15. Pemindahan Barang (KIRI"]
    )
    df = load_sheet(workbook_path, candidates, skip_rows=3)
    if logger:
        logger(
            f"[loader] Kirim (42.15): sheet='{getattr(df, 'attrs', {}).get('sheet', '?')}', "
            f"{len(df)} baris, {len(df.columns)} kolom (engine={_EXCEL_ENGINE or 'default'})"
        )
    return df


# ============================================================
# LOAD 2: TERIMA 42.05
# ============================================================

def load_terima(
    workbook_path: str | os.PathLike,
    config: dict,
    logger=None,
) -> pd.DataFrame:
    """Load sheet Terima (42.05). Sheet '#42.05. Rincian Pemindahan Bara'."""
    candidates = config.get("input_sheets", {}).get(
        "terima", ["#42.05. Rincian Pemindahan Bara"]
    )
    df = load_sheet(workbook_path, candidates, skip_rows=3)
    if logger:
        logger(f"[loader] Terima (42.05): {len(df)} baris, {len(df.columns)} kolom")
    return df


# ============================================================
# LOAD 3: KATEGORI
# ============================================================

def load_kategori(
    workbook_path: str | os.PathLike,
    config: dict,
    logger=None,
) -> pd.DataFrame:
    """Load sheet DBS Kategori (Sheet1). Header di baris pertama (skip=0)."""
    candidates = config.get("input_sheets", {}).get(
        "kategori", ["Sheet1", "DBS_KATEGORI"]
    )
    df = load_sheet(workbook_path, candidates, skip_rows=0)
    if logger:
        logger(f"[loader] Kategori: {len(df)} baris, {len(df.columns)} kolom")
    return df
