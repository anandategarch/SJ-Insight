"""sj_ui — paket UI PySide6 untuk SJ Insight.

Mengikuti arsitektur ``ui`` Waste (v2.6.0 rebuild):
  - shared    : konstanta + helper murni
  - theme     : palet warna + QSS global
  - state     : AppContext (state tunggal + signal Qt)
  - models    : PandasTableModel (QAbstractTableModel di atas DataFrame)
  - workers   : QThread worker (pipeline/export)
  - widgets   : blok anti-terpotong (WLabel, DataTable, KPI)
  - view_*    : 4 tab (Dashboard / Anomali / Kirim vs Terima / Pipeline)
  - app       : MainWindow + main()
"""
from __future__ import annotations

__version__ = "1.0.0"
