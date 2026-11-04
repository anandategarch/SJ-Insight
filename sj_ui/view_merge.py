"""view_merge.py — tab Kirim vs Terima: rekonsiliasi dokumen (mirror Waste view_periode).

Tampilan:
  - Tabel merge (Kirim × Terima × Kategori) — semua kolom kirim + kolom terima + Selisih Hari
  - Highlight: kirim tanpa dokumen terima (merah) + terima lambat (selisih hari ≥ 2)
"""
from __future__ import annotations

import pandas as pd
from PySide6.QtWidgets import (
    QHBoxLayout,
    QLabel,
    QPushButton,
    QVBoxLayout,
    QWidget,
)

from .state import AppContext
from .widgets import DataTable


MERGE_COLS = [
    "No. Pemindahan #",
    "Tanggal",
    "Gudang Tujuan (Kode)",
    "Nama Barang",
    "Status Pemindahan Barang",
    "Memiliki Dokumen [Terima]",
    "tNomorTerima",
    "tTanggalTerima",
    "Selisih Hari",
    "tStatusPengiriman",
    "KTG_Kiriman[DC/RESTO]",
    "JENIS (Kategori)",
    "Concet",
]


class MergeView(QWidget):
    """Tab Kirim vs Terima — rekonsiliasi dokumen 42.05."""

    def __init__(self, ctx: AppContext, parent: QWidget | None = None) -> None:
        super().__init__(parent)
        self.setObjectName("mergeView")
        self.ctx = ctx

        layout = QVBoxLayout(self)
        layout.setContentsMargins(16, 16, 16, 16)
        layout.setSpacing(10)

        title = QLabel("Kirim vs Terima — Rekonsiliasi Dokumen 42.05")
        title.setStyleSheet("font-size: 16px; font-weight: 700; color: #1c1917;")
        layout.addWidget(title)
        sub = QLabel(
            "Left join Kirim ← Terima on No. Pemindahan # = Nomor #Kirim · "
            "Selisih Hari = Tanggal #Terima − Tanggal kirim (≥ 2 hari = terima lambat)"
        )
        sub.setObjectName("muted")
        sub.setStyleSheet("font-size: 12px;")
        layout.addWidget(sub)

        # Filter toggle: hanya kirim tanpa terima
        filter_row = QHBoxLayout()
        self._only_missing_btn = QPushButton("⚑ Hanya Kirim tanpa Terima")
        self._only_missing_btn.setCheckable(True)
        self._only_missing_btn.toggled.connect(self._refresh)
        filter_row.addWidget(self._only_missing_btn)

        self._only_lambat_btn = QPushButton("⚑ Hanya Terima Lambat (≥2 hari)")
        self._only_lambat_btn.setCheckable(True)
        self._only_lambat_btn.toggled.connect(self._refresh)
        filter_row.addWidget(self._only_lambat_btn)

        filter_row.addStretch()
        self._count_label = QLabel("0 baris")
        self._count_label.setObjectName("muted")
        filter_row.addWidget(self._count_label)
        layout.addLayout(filter_row)

        # Tabel
        self._table = DataTable()
        layout.addWidget(self._table, 1)

        self.ctx.data_ready.connect(self._refresh)

    def _refresh(self) -> None:
        df = self.ctx.df_merged
        if df is None or df.empty:
            self._table.set_dataframe(pd.DataFrame(columns=MERGE_COLS))
            self._count_label.setText("0 baris")
            return

        view = df.copy()
        if self._only_missing_btn.isChecked() and "Memiliki Dokumen [Terima]" in view.columns:
            view = view[~view["Memiliki Dokumen [Terima]"]]
        if self._only_lambat_btn.isChecked() and "Selisih Hari" in view.columns:
            view = view[view["Selisih Hari"].apply(lambda x: isinstance(x, (int, float)) and x >= 2)]

        cols_present = [c for c in MERGE_COLS if c in view.columns]
        self._table.set_dataframe(view[cols_present].copy())
        self._count_label.setText(f"{len(view)} baris")
