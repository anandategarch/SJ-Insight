"""logging_setup.py — logger file (rotasi 1 MB × 3) + console.

Di-port dari ``waste_anomaly.logging_setup`` (sedikit disederhanakan).
"""
from __future__ import annotations

import logging
import sys
from logging.handlers import RotatingFileHandler
from pathlib import Path

from .frozen import exe_dir

LOG_DIR = exe_dir() / "logs"
LOG_FILE = LOG_DIR / "sj_insight.log"


def setup_logging(level: int = logging.INFO) -> logging.Logger:
    """Pasang root logger: file rotasi + console (stderr).

    Aman dipanggil berulang (handler lama dibuang).
    """
    root = logging.getLogger()
    for h in list(root.handlers):
        root.removeHandler(h)

    fmt = logging.Formatter(
        fmt="%(asctime)s [%(levelname)s] %(name)s: %(message)s",
        datefmt="%Y-%m-%d %H:%M:%S",
    )

    try:
        LOG_DIR.mkdir(parents=True, exist_ok=True)
        fh = RotatingFileHandler(
            LOG_FILE, maxBytes=1_000_000, backupCount=3, encoding="utf-8"
        )
        fh.setFormatter(fmt)
        fh.setLevel(level)
        root.addHandler(fh)
    except OSError:
        # Folder read-only (mis. mode frozen di Program Files) — console saja.
        pass

    ch = logging.StreamHandler(sys.stderr)
    ch.setFormatter(fmt)
    ch.setLevel(level)
    root.addHandler(ch)

    root.setLevel(level)
    return logging.getLogger("sj_anomaly")
