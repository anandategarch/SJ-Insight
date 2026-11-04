"""frozen.py — deteksi mode frozen (PyInstaller) + jalur bundle/exe.

Di-port dari ``waste_anomaly.frozen`` supaya config.json / followup.json
bisa disimpan berdampingan dengan EXE pada mode frozen (perilaku sama
seperti Waste v2.4.1 — sidecar writable, bukan _MEIPASS temp yang
dihapus saat exit).
"""
from __future__ import annotations

import sys
from pathlib import Path


def is_frozen() -> bool:
    """True bila berjalan sebagai PyInstaller EXE."""
    return getattr(sys, "frozen", False) and hasattr(sys, "_MEIPASS")


def bundle_dir() -> Path:
    """Folder resource bundled (read-only).

    Dev: repo root (parent dari paket ``sj_anomaly``).
    Frozen: ``sys._MEIPASS`` (temp PyInstaller).
    """
    if is_frozen():
        return Path(sys._MEIPASS)  # type: ignore[attr-defined]
    return Path(__file__).resolve().parent.parent


def exe_dir() -> Path:
    """Folder EXE (writable). Dev: repo root."""
    if is_frozen():
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent.parent
