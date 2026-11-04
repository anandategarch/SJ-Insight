"""shared.py — konstanta + helper murni lintas UI (mirror Waste ui/shared).

Kontrak publik:
  - APP_DIR / BUNDLE_DIR (frozen-aware)
  - LOG_RUN_LINES (satu ambang lintas penampang log)
  - fmt_id (format angka id-ID)
"""
from __future__ import annotations

import math
from pathlib import Path

from sj_anomaly.frozen import bundle_dir as _bundle_dir
from sj_anomaly.frozen import exe_dir as _exe_dir

APP_DIR = _exe_dir()
BUNDLE_DIR = _bundle_dir()

APP_VERSION = "1.0.0"
APP_TITLE = "SJ Insight — Analisa Pemindahan Barang"

LOG_RUN_LINES = 2000


def fmt_id(value, decimals: int = 0) -> str:
    """Format angka gaya Indonesia: ribuan '.', desimal ','."""
    if value is None:
        return ""
    try:
        f = float(value)
    except (TypeError, ValueError):
        return str(value)
    if math.isnan(f):
        return ""
    s = f"{f:,.{decimals}f}"
    return s.replace(",", "X").replace(".", ",").replace("X", ".")
