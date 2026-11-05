"""view_laporan.py — tab Laporan Lengkap: tampilkan 8 view sheet mirror xlsb.

8 sheet view yang di-generate dari pipeline (mirror sheet di file asli
'04.QUERRY SJ.xlsb'):
  - Upload          — view gabungan Kirim×Terima (mirror sheet UPLOAD)
  - Terima Pivot    — pivot Kirim dgn Status Anomali (mirror sheet TERIMA)
  - Manual          — filter Kirim subset (mirror sheet MANUAL)
  - Kirim Pivot     — pivot Kirim mentah (mirror sheet KIRIM)
  - Double Terima   — filter Count>1 Concet, view Terima (mirror DOUBLE TERIMA)
  - Double Kirim    — filter Count>1 Concet, view Kirim (mirror DOUBLE KIRIM)
  - 4215 Terima     — flat table = df_kirim (mirror 4215_TERIMA)
  - 4215 Kirim      — flat table subset (mirror 4215_KIRIM)
"""
from __future__ import annotations

import pandas as pd
from PySide6.QtCore import Qt
from PySide6.QtWidgets import (
    QComboBox,
    QHBoxLayout,
    QLabel,
    QPushButton,
    QVBoxLayout,
    QWidget,
)

from sj_anomaly.sheets import (
    DOUBLE_KIRIM_COLS,
    DOUBLE_TERIMA_COLS,
    KIRIM_PIVOT_COLS,
    MANUAL_COLS,
    TERIMA_PIVOT_COLS,
    UPLOAD_COLS,
    generate_all_views,
)

from .state import AppContext
from .widgets import DataTable


class LaporanView(QWidget):
    """Tab Laporan Lengkap — 8 view sheet mirror file asli xlsb."""

    VIEW_LABELS = [
        ("Upload",        "View gabungan Kirim×Terima (mirror sheet UPLOAD)"),
        ("Terima Pivot",  "Pivot Kirim dgn Status Anomali (mirror sheet TERIMA)"),
        ("Manual",        "Filter Kirim subset (mirror sheet MANUAL)"),
        ("Kirim Pivot",   "Pivot Kirim mentah (mirror sheet KIRIM)"),
        ("Double Terima", "Filter Count>1 Concet, view Terima (mirror DOUBLE TERIMA)"),
        ("Double Kirim",  "Filter Count>1 Concet, view Kirim (mirror DOUBLE KIRIM)"),
        ("4215 Terima",   "Flat table = df_kirim (mirror 4215_TERIMA)"),
        ("4215 Kirim",    "Flat table subset (mirror 4215_KIRIM)"),
    ]

    def __init__(self, ctx: AppContext, parent: QWidget | None = None) -> None:
        super().__init__(parent)
        self.setObjectName("laporanView")
        self.ctx = ctx
        self._current_view: str | None = None

        layout = QVBoxLayout(self)
        layout.setContentsMargins(16, 16, 16, 16)
        layout.setSpacing(10)

        # ---- Header ----
        title = QLabel("📋 Laporan Lengkap — 8 View Sheet (mirror file asli xlsb)")
        title.setStyleSheet("font-size: 16px; font-weight: 700; color: #1c1917;")
        layout.addWidget(title)
        sub = QLabel(
            "8 view yang sama dengan sheet di file '04.QUERRY SJ.xlsb'. "
            "Di-generate dari pipeline. Pilih view utk lihat data."
        )
        sub.setObjectName("muted")
        sub.setStyleSheet("font-size: 12px;")
        layout.addWidget(sub)

        # ---- View selector ----
        selector_row = QHBoxLayout()
        selector_row.setSpacing(8)
        lbl = QLabel("Pilih View:")
        lbl.setStyleSheet("font-weight: 600; color: #1c1917;")
        selector_row.addWidget(lbl)

        self._view_combo = QComboBox()
        for key, desc in self.VIEW_LABELS:
            self._view_combo.addItem(f"{key} — {desc}", key)
        self._view_combo.currentIndexChanged.connect(self._on_view_changed)
        selector_row.addWidget(self._view_combo, 1)

        self._refresh_btn = QPushButton("🔄 Refresh")
        self._refresh_btn.setObjectName("ghost")
        self._refresh_btn.setCursor(Qt.PointingHandCursor)
        self._refresh_btn.setToolTip("Re-generate view dari data terbaru")
        self._refresh_btn.clicked.connect(self._refresh)
        selector_row.addWidget(self._refresh_btn)
        layout.addLayout(selector_row)

        # ---- Info bar ----
        self._info_label = QLabel("Belum ada data — jalankan pipeline dulu di tab Pipeline.")
        self._info_label.setObjectName("muted")
        self._info_label.setStyleSheet("font-size: 12px; color: #78716c;")
        layout.addWidget(self._info_label)

        # ---- Table ----
        self._table = DataTable()
        layout.addWidget(self._table, 1)

        # Connect signals
        self.ctx.data_ready.connect(self._refresh)

    # ------------------------------------------------------------
    # View switching
    # ------------------------------------------------------------
    def _on_view_changed(self, idx: int) -> None:
        if idx < 0 or idx >= len(self.VIEW_LABELS):
            return
        key = self.VIEW_LABELS[idx][0]
        self._current_view = key
        self._render_view(key)

    def _refresh(self) -> None:
        """Re-generate view dari data terbaru."""
        if self._current_view is None and self.VIEW_LABELS:
            self._current_view = self.VIEW_LABELS[0][0]
        if self._current_view:
            self._render_view(self._current_view)

    def _render_view(self, view_key: str) -> None:
        """Render satu view ke table."""
        if self.ctx.df is None or self.ctx.df_merged is None:
            self._info_label.setText("Belum ada data — jalankan pipeline dulu di tab Pipeline.")
            self._table.set_df(pd.DataFrame())
            return

        # Generate all views (cheap — pure transform dari DataFrame)
        try:
            views = generate_all_views(
                self.ctx.df, self.ctx.df_merged, self.ctx.df_count
            )
        except Exception as exc:
            self._info_label.setText(f"⚠ Gagal generate view: {exc}")
            return

        df_view = views.get(view_key)
        if df_view is None or df_view.empty:
            self._info_label.setText(f"View '{view_key}': tidak ada data (kosong).")
            self._table.set_df(pd.DataFrame())
            return

        self._table.set_df(df_view)
        self._info_label.setText(
            f"View '{view_key}': {len(df_view)} baris × {len(df_view.columns)} kolom"
        )
