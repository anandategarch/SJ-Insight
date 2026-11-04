"""
ui_app.py — ENTRY POINT aplikasi SJ Insight (tetap dipakai MULAI.bat).

Struktur:
    sj_anomaly/   — pipeline + stages + loader + exporter + config.json (mirror waste_anomaly/)
    sj_ui/        — app + state + workers + theme + widgets + models + view_* (mirror ui/)

Run:
    python ui_app.py
"""
from __future__ import annotations

import sys
from pathlib import Path

# Defensive: jalan dari cwd mana pun tetap menemukan paket sj_ui & sj_anomaly
sys.path.insert(0, str(Path(__file__).resolve().parent))

from sj_ui.app import main  # noqa: E402

if __name__ == "__main__":
    raise SystemExit(main())
