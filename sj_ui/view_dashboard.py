"""view_dashboard.py — tab Dashboard: KPI cards + ringkasan chart (mirror Waste).

Tampilan:
  - 8 KPI cards (total transaksi / anomali satuan / anomali QTY / anomali konversi /
    perlu cek / kirim tanpa terima / terima lambat / duplikat concet)
  - Tabel per gudang tujuan (top 12)
  - Tabel per barang (top 12 by anomali)
  - Ringkasan periode + sumber file
"""
from __future__ import annotations

import pandas as pd
from PySide6.QtWidgets import (
    QFrame,
    QGridLayout,
    QLabel,
    QScrollArea,
    QVBoxLayout,
    QWidget,
)

from sj_anomaly.utils import fmt_date_id

from .shared import fmt_id
from .state import AppContext
from .widgets import DataTable, KpiCard


class DashboardView(QWidget):
    """Tab Dashboard — KPI + ringkasan + tabel top gudang/barang."""

    def __init__(self, ctx: AppContext, parent: QWidget | None = None) -> None:
        super().__init__(parent)
        self.setObjectName("dashboardView")
        self.ctx = ctx

        # Outer scroll area supaya konten panjang tetap terjangkau
        outer = QVBoxLayout(self)
        outer.setContentsMargins(16, 16, 16, 16)
        outer.setSpacing(12)

        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setFrameShape(QScrollArea.NoFrame)
        outer.addWidget(scroll, 1)

        content = QWidget()
        scroll.setWidget(content)
        self._layout = QVBoxLayout(content)
        self._layout.setSpacing(14)

        # Title
        title = QLabel("Dashboard — Ringkasan Hasil Analisa")
        title.setStyleSheet("font-size: 16px; font-weight: 700; color: #1c1917;")
        self._layout.addWidget(title)

        # Periode info
        self._periode_label = QLabel("")
        self._periode_label.setObjectName("muted")
        self._layout.addWidget(self._periode_label)

        # KPI grid
        kpi_frame = QFrame()
        kpi_frame.setObjectName("card")
        kpi_layout = QGridLayout(kpi_frame)
        kpi_layout.setContentsMargins(14, 12, 14, 12)
        kpi_layout.setSpacing(10)
        self._kpi_total = KpiCard("Total Transaksi", "0", "baris kirim diproses", "default")
        self._kpi_anom_sat = KpiCard("Anomali Satuan", "0", "satuan input ≠ standar", "anomali")
        self._kpi_anom_qty = KpiCard("Anomali QTY", "0", "5 rule: nol/desimal/spike/kecil", "anomali")
        self._kpi_anom_konv = KpiCard("Anomali Konversi", "0", "QTY Terkecil 2 menyimpang > 10%", "anomali")
        self._kpi_perlu_cek = KpiCard("Perlu Cek", "0", "deviasi 5-10%", "warn")
        self._kpi_tanpa = KpiCard("Kirim tanpa Terima", "0", "dokumen 42.05 tidak ditemukan", "anomali")
        self._kpi_lambat = KpiCard("Terima Lambat", "0", "selisih hari ≥ 2", "warn")
        self._kpi_dup = KpiCard("Grup Duplikat", "0", "Concet dengan count > 1", "warn")
        for i, card in enumerate([
            self._kpi_total, self._kpi_anom_sat, self._kpi_anom_qty, self._kpi_anom_konv,
            self._kpi_perlu_cek, self._kpi_tanpa, self._kpi_lambat, self._kpi_dup,
        ]):
            kpi_layout.addWidget(card, i // 4, i % 4)
        self._layout.addWidget(kpi_frame)

        # Tabel per gudang
        self._per_gudang_title = QLabel("Top 12 Gudang Tujuan (by volume)")
        self._per_gudang_title.setStyleSheet("font-weight: 600; color: #1c1917;")
        self._layout.addWidget(self._per_gudang_title)
        self._table_gudang = DataTable()
        self._layout.addWidget(self._table_gudang, 1)

        # Tabel per barang
        self._per_barang_title = QLabel("Top 12 Barang (by anomali)")
        self._per_barang_title.setStyleSheet("font-weight: 600; color: #1c1917;")
        self._layout.addWidget(self._per_barang_title)
        self._table_barang = DataTable()
        self._layout.addWidget(self._table_barang, 1)

        # Listen data_ready
        self.ctx.data_ready.connect(self._refresh)

    def _refresh(self) -> None:
        summary = self.ctx.summary
        if not summary:
            return

        # KPI
        self._kpi_total.set_value(fmt_id(summary.get("total_transaksi", 0)))
        self._kpi_anom_sat.set_value(fmt_id(summary.get("anomali_satuan", 0)))
        self._kpi_anom_qty.set_value(fmt_id(summary.get("anomali_qty", 0)))
        self._kpi_anom_konv.set_value(fmt_id(summary.get("anomali_konversi", 0)))
        self._kpi_perlu_cek.set_value(fmt_id(summary.get("perlu_cek", 0)))
        self._kpi_tanpa.set_value(fmt_id(summary.get("kirim_tanpa_terima", 0)))
        self._kpi_lambat.set_value(fmt_id(summary.get("terima_lambat", 0)))
        self._kpi_dup.set_value(fmt_id(summary.get("duplikat_concet", 0)))

        # Periode
        tmin = fmt_date_id(summary.get("tanggal_min")) if summary.get("tanggal_min") else "-"
        tmax = fmt_date_id(summary.get("tanggal_max")) if summary.get("tanggal_max") else "-"
        self._periode_label.setText(
            f"Periode: {tmin} – {tmax} · "
            f"{summary.get('total_barang', 0)} barang unik · "
            f"{summary.get('total_gudang_tujuan', 0)} gudang tujuan · "
            f"rata2 selisih hari: {fmt_id(summary.get('rata2_selisih_hari'), 2)}"
        )

        # Per gudang
        per_gudang = summary.get("per_gudang") or []
        if per_gudang:
            df_g = pd.DataFrame(per_gudang)
            df_g = df_g.rename(columns={
                "gudang": "Gudang Tujuan",
                "total": "Total Kirim",
                "anomali": "Anomali",
            })
            df_g["Anomali %"] = df_g["Anomali"] / df_g["Total Kirim"].replace(0, 1) * 100
            self._table_gudang.set_dataframe(df_g)

        # Per barang
        per_barang = summary.get("per_barang") or []
        if per_barang:
            df_b = pd.DataFrame(per_barang)
            df_b = df_b.rename(columns={
                "nama_barang": "Nama Barang",
                "total": "Total Kirim",
                "anomali": "Anomali",
            })
            df_b["Anomali %"] = df_b["Anomali"] / df_b["Total Kirim"].replace(0, 1) * 100
            self._table_barang.set_dataframe(df_b)
