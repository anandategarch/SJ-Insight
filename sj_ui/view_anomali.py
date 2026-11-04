"""view_anomali.py — tab Anomali: tabel baris anomali ter-filter.

Mirror Waste view_anomali (tapi lebih sederhana — SJ Insight fokus pada
tabel anomali terurut + filter gudang/barang/status).

Kolom ditampilkan (selection dari hasil pipeline):
  No. Pemindahan | Tanggal | Gudang Tujuan | Nama Barang | Satuan | QTY Terkecil |
  Kuantitas | Status Pemindahan | Status Anomali Satuan | Status Anomali QTY |
  Status Anomali QTY Terkecil 2 | Analisa QTY Terkecil 2 | Dok Terima |
  Selisih Hari
"""
from __future__ import annotations

import pandas as pd
from PySide6.QtCore import Qt
from PySide6.QtWidgets import (
    QComboBox,
    QFrame,
    QHBoxLayout,
    QLabel,
    QPushButton,
    QVBoxLayout,
    QWidget,
)

from .state import AppContext
from .widgets import DataTable


# Kolom yang ditampilkan di tab Anomali (selection dari hasil pipeline)
ANOMALI_COLS = [
    "No. Pemindahan #",
    "Tanggal",
    "Gudang Pengirim (Kode)",
    "Gudang Tujuan (Kode)",
    "Nama Barang",
    "Satuan",
    "QTY Terkecil",
    "Kuantitas",
    "Status Pemindahan Barang",
    "Status Anomali Satuan",
    "Status Anomali QTY",
    "Status Anomali QTY Terkecil 2",
    "Analisa QTY Terkecil 2",
    "Memiliki Dokumen [Terima]",
    "Selisih Hari",
    "JENIS (Kategori)",
]


class AnomaliView(QWidget):
    """Tab Anomali — tabel baris dengan salah satu status ≠ NORMAL."""

    def __init__(self, ctx: AppContext, parent: QWidget | None = None) -> None:
        super().__init__(parent)
        self.setObjectName("anomaliView")
        self.ctx = ctx

        layout = QVBoxLayout(self)
        layout.setContentsMargins(16, 16, 16, 16)
        layout.setSpacing(10)

        # ---- Header ----
        title = QLabel("Anomali — Baris dengan Status ≠ NORMAL")
        title.setStyleSheet("font-size: 16px; font-weight: 700; color: #1c1917;")
        layout.addWidget(title)
        sub = QLabel(
            "Klik header kolom untuk sort · sel berwarna = severity "
            "(merah=anomali, oranye=perlu cek, kuning=perhatian)"
        )
        sub.setObjectName("muted")
        sub.setStyleSheet("font-size: 12px;")
        layout.addWidget(sub)

        # ---- Filter bar ----
        filter_frame = QFrame()
        filter_frame.setObjectName("card")
        fh = QHBoxLayout(filter_frame)
        fh.setContentsMargins(10, 8, 10, 8)
        fh.setSpacing(8)

        fh.addWidget(QLabel("Gudang:"))
        self._combo_gudang = QComboBox()
        self._combo_gudang.setMinimumWidth(140)
        self._combo_gudang.currentTextChanged.connect(
            lambda v: self.ctx.set_filter("gudang", v if v != "(semua)" else None)
        )
        fh.addWidget(self._combo_gudang)

        fh.addWidget(QLabel("Barang:"))
        self._combo_barang = QComboBox()
        self._combo_barang.setMinimumWidth(200)
        self._combo_barang.currentTextChanged.connect(
            lambda v: self.ctx.set_filter("barang", v if v != "(semua)" else None)
        )
        fh.addWidget(self._combo_barang)

        fh.addWidget(QLabel("Status Konversi:"))
        self._combo_status = QComboBox()
        self._combo_status.addItems([
            "(semua)", "NORMAL", "ANOMALI KONVERSI", "PERLU CEK",
            "KONVERSI TIDAK STABIL", "DATA TERBATAS", "TIDAK ADA STANDAR", "TIDAK DAPAT DIANALISA",
        ])
        self._combo_status.currentTextChanged.connect(
            lambda v: self.ctx.set_filter("status", v if v != "(semua)" else None)
        )
        fh.addWidget(self._combo_status)

        fh.addStretch()
        self._count_label = QLabel("0 baris")
        self._count_label.setObjectName("muted")
        fh.addWidget(self._count_label)

        reset_btn = QPushButton("✕ Reset Filter")
        reset_btn.clicked.connect(self._reset_filter)
        fh.addWidget(reset_btn)

        layout.addWidget(filter_frame)

        # ---- Tabel ----
        self._table = DataTable()
        layout.addWidget(self._table, 1)

        # Listen signals
        self.ctx.data_ready.connect(self._on_data_ready)
        self.ctx.filters_changed.connect(self._refresh)

    def _on_data_ready(self) -> None:
        """Isi ulang combo filter dari dataset baru."""
        df = self.ctx.df
        if df is None:
            return
        # Combo gudang
        gudang_list = ["(semua)"] + sorted(df["Gudang Tujuan (Kode)"].dropna().unique().tolist())
        self._combo_gudang.blockSignals(True)
        self._combo_gudang.clear()
        self._combo_gudang.addItems(gudang_list)
        self._combo_gudang.blockSignals(False)

        # Combo barang
        barang_list = ["(semua)"] + sorted(df["Nama Barang"].dropna().unique().tolist())
        self._combo_barang.blockSignals(True)
        self._combo_barang.clear()
        self._combo_barang.addItems(barang_list)
        self._combo_barang.blockSignals(False)

        self._refresh()

    def _refresh(self) -> None:
        df = self.ctx.anomali_df()
        if df is None or df.empty:
            self._table.set_dataframe(pd.DataFrame(columns=ANOMALI_COLS))
            self._count_label.setText("0 baris anomali")
            return
        # Ambil subset kolom
        cols_present = [c for c in ANOMALI_COLS if c in df.columns]
        df_view = df[cols_present].copy()
        # Tambah kolom __index utk tracking baris asli (utk follow-up)
        df_view["__row_idx"] = df.index
        self._table.set_dataframe(df_view)
        self._count_label.setText(f"{len(df_view)} baris anomali")

    def _reset_filter(self) -> None:
        self.ctx.reset_filters()
        self._combo_gudang.blockSignals(True)
        self._combo_gudang.setCurrentIndex(0)
        self._combo_gudang.blockSignals(False)
        self._combo_barang.blockSignals(True)
        self._combo_barang.setCurrentIndex(0)
        self._combo_barang.blockSignals(False)
        self._combo_status.blockSignals(True)
        self._combo_status.setCurrentIndex(0)
        self._combo_status.blockSignals(False)
        self.ctx.filters_changed.emit()
