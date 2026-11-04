"""utils.py — helper murni untuk SJ Insight.

Mengikuti pola ``waste_anomaly.utils``:
  - load_config + sidecar frozen-aware
  - clamp, norm_text
  - format angka id-ID (fmt_num / fmt_ratio / fmt_pct / fmt_date_id)
  - severity status (untuk pewarnaan tabel UI)
"""
from __future__ import annotations

import json
import math
import os
import shutil
import warnings
from pathlib import Path

from .frozen import bundle_dir, exe_dir, is_frozen


# ============================================================
# KONFIGURASI
# ============================================================

CONFIG_PATH = Path(__file__).resolve().parent / "config.json"


def _seed_sidecar_config(bundled: Path, sidecar: Path) -> bool:
    """Frozen first-run: salin config bundled → sidecar."""
    try:
        sidecar.parent.mkdir(parents=True, exist_ok=True)
        shutil.copyfile(bundled, sidecar)
        return True
    except OSError:
        return False


def _active_config_path() -> Path:
    """Jalur config aktif (dinamis — frozen-aware)."""
    if not is_frozen():
        return CONFIG_PATH
    sidecar = exe_dir() / "config.json"
    if sidecar.exists():
        return sidecar
    bundled = bundle_dir() / "sj_anomaly" / "config.json"
    if bundled.exists():
        if _seed_sidecar_config(bundled, sidecar) and sidecar.exists():
            return sidecar
        return bundled
    return sidecar


_DEFAULT_CONFIG: dict = {
    "thresholds": {
        "QTY_EXTREME_FACTOR":   3.5,
        "QTY_TINY_FACTOR":      0.1,
        "MIN_TRANSAKSI":        3,
        "DEVIASI_ANOMALI":      0.10,
        "DEVIASI_PERLU_CEK":    0.05,
        "DOMINASI_MIN":         0.50,
        "TERIMA_LAMBAT_HARI":   2,
    },
    "whole_units":     ["PCS", "KARTON", "KARUNG", "JURIGEN", "PACK", "BOTTLE", "BOX", "CUP"],
    "diterima_status": ["Diterima Seluruh nya", "Diterima Seluruhnya"],
    "depo_replacement": "DCSMGMUR",
    "default_key_akhir": 31,
}


def load_config(config_path: str | os.PathLike | None = None) -> dict:
    """Muat config.json. Gagal baca → default + key ``_config_load_error``."""
    p = Path(config_path) if config_path else _active_config_path()
    try:
        with open(p, "r", encoding="utf-8") as f:
            return json.load(f)
    except Exception as e:
        warnings.warn(f"Config load gagal: {p}: {e} — pakai default")
        cfg = {k: (v.copy() if isinstance(v, (dict, list)) else v)
               for k, v in _DEFAULT_CONFIG.items()}
        cfg["_config_load_error"] = str(e)
        return cfg


# ============================================================
# ANGKA / BANTUAN MURNI
# ============================================================

def clamp(value: int, lo: int, hi: int) -> int:
    """Batasi value ke [lo, hi]."""
    return max(lo, min(hi, value))


_INVISIBLE_CHARS = ("\xa0", "\u200b", "\u200c", "\u200d", "\ufeff")


def norm_text(value: object) -> str:
    """Normalisasi kunci teks: NBSP→spasi, zero-width dihapus, strip + collapse."""
    if value is None:
        return ""
    if isinstance(value, float) and math.isnan(value):
        return ""
    text = str(value).replace("\xa0", " ")
    for ch in _INVISIBLE_CHARS:
        text = text.replace(ch, "")
    return " ".join(text.split())


# ============================================================
# FORMAT ANGKA id-ID (mirror Number.ToText M code)
# ============================================================

def _id_num(value: float, decimals: int) -> str:
    """Format angka gaya Indonesia: ribuan '.', desimal ','."""
    s = f"{value:,.{decimals}f}"
    return s.replace(",", "X").replace(".", ",").replace("X", ".")


def fmt_num(value, max_dec: int = 2) -> str:
    """'#,##0.##' id-ID → 1.234,5 (tanpa nol ekor)."""
    if value is None:
        return "-"
    try:
        f = float(value)
    except (TypeError, ValueError):
        return str(value)
    if math.isnan(f) or math.isinf(f):
        return "-"
    # Bulatkan ke max_dec dulu lalu potong nol ekor
    rounded = round(f, max_dec)
    if rounded == int(rounded):
        return _id_num(rounded, 0)
    # Hitung digit desimal efektif
    s = f"{rounded:.{max_dec}f}"
    # Buang trailing zero
    s = s.rstrip("0").rstrip(".")
    decimals = 0 if "." not in s else len(s.split(".")[1])
    return _id_num(rounded, decimals)


def fmt_ratio(value, max_dec: int = 6) -> str:
    """'#,##0.######' id-ID → hingga 6 desimal (utk rasio konversi)."""
    if value is None:
        return "-"
    try:
        f = float(value)
    except (TypeError, ValueError):
        return str(value)
    if math.isnan(f) or math.isinf(f):
        return "-"
    s = f"{f:.{max_dec}f}".rstrip("0").rstrip(".")
    decimals = 0 if "." not in s else len(s.split(".")[1])
    return _id_num(f, decimals)


def fmt_pct(value) -> str:
    """'0.00%' id-ID → 12,34% (input fraksi, dikali 100)."""
    if value is None:
        return "-"
    try:
        f = float(value)
    except (TypeError, ValueError):
        return str(value)
    if math.isnan(f) or math.isinf(f):
        return "-"
    return _id_num(f * 100, 2) + "%"


_BULAN_ID = (
    "Jan", "Feb", "Mar", "Apr", "Mei", "Jun",
    "Jul", "Agu", "Sep", "Okt", "Nov", "Des",
)


def fmt_date_id(iso: str | None) -> str:
    """'dd MMM yyyy' id-ID → 07 Nov 2025. Input ISO 'YYYY-MM-DD[ HH:MM:SS]'."""
    if not iso:
        return "-"
    import re
    m = re.match(r"^(\d{4})-(\d{2})-(\d{2})", str(iso))
    if not m:
        return str(iso)
    y, mo, d = m.group(1), m.group(2), m.group(3)
    return f"{d} {_BULAN_ID[int(mo) - 1]} {y}"


def fmt_datetime_id(iso: str | None) -> str:
    """'dd MMM yyyy HH:MM' id-ID."""
    if not iso:
        return "-"
    base = fmt_date_id(iso)
    import re
    t = re.search(r"(\d{2}:\d{2})", str(iso))
    return f"{base} {t.group(1)}" if t else base


# ============================================================
# SEVERITY STATUS (untuk pewarnaan tabel UI)
# ============================================================

Severity = str  # "anomali" | "perlu-cek" | "perhatian" | "normal"


def severity_of_status(status: str) -> Severity:
    """Klasifikasi status → severity untuk pewarnaan UI."""
    s = (status or "").upper()
    if s.startswith("ANOMALI"):
        return "anomali"
    if s.startswith("PERLU CEK"):
        return "perlu-cek"
    if s in {
        "DATA TERBATAS",
        "KONVERSI TIDAK STABIL",
        "TIDAK ADA STANDAR",
        "TIDAK DAPAT DIANALISA",
    }:
        return "perhatian"
    return "normal"
