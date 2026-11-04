"""widgets.py — blok UI anti-terpotong (mirror Waste widgets).

Komponen:
  - WLabel    : QLabel yang word-wrap + elide anti-terpotong
  - KpiCard   : kartu KPI (judul + nilai + sub-teks + warna accent)
  - DataTable : QTableView di atas PandasTableModel (sort + tooltip + alt row)
"""
from __future__ import annotations

from PySide6.QtCore import Qt
from PySide6.QtWidgets import (
    QFrame,
    QHBoxLayout,
    QHeaderView,
    QLabel,
    QSizePolicy,
    QTableView,
    QVBoxLayout,
    QWidget,
)

from .models import PandasTableModel
from .theme import STATUS_COLORS


class WLabel(QLabel):
    """QLabel yang word-wrap + tidak terpotong (mirror Waste WLabel)."""

    def __init__(self, text: str = "", parent: QWidget | None = None,
                 wrap: bool = True, muted: bool = False) -> None:
        super().__init__(text, parent)
        self.setWordWrap(wrap)
        self.setTextInteractionFlags(Qt.TextSelectableByMouse)
        self.setObjectName("muted" if muted else "")
        self.setSizePolicy(QSizePolicy.Preferred, QSizePolicy.Minimum)


class KpiCard(QFrame):
    """Kartu KPI: label + nilai besar + sub-teks + accent warna.

    Mirror Waste KPI card. ``accent`` bisa: 'default' | 'anomali' | 'warn'.
    """

    def __init__(
        self,
        title: str,
        value: str,
        subtitle: str = "",
        accent: str = "default",
        parent: QWidget | None = None,
    ) -> None:
        super().__init__(parent)
        self.setObjectName("kpiframe")
        layout = QVBoxLayout(self)
        layout.setContentsMargins(14, 12, 14, 12)
        layout.setSpacing(2)

        self._title = QLabel(title)
        self._title.setObjectName("hint")
        self._title.setWordWrap(True)
        layout.addWidget(self._title)

        self._value = QLabel(value)
        self._value.setStyleSheet(
            f"font-size: 22px; font-weight: 700; color: {self._accent_color(accent)};"
        )
        layout.addWidget(self._value)

        if subtitle:
            self._subtitle = QLabel(subtitle)
            self._subtitle.setObjectName("muted")
            self._subtitle.setStyleSheet("font-size: 11px;")
            self._subtitle.setWordWrap(True)
            layout.addWidget(self._subtitle)

    def set_value(self, value: str) -> None:
        self._value.setText(value)

    @staticmethod
    def _accent_color(accent: str) -> str:
        if accent == "anomali":
            return STATUS_COLORS["anomali"][0]
        if accent == "warn":
            return STATUS_COLORS["perlu-cek"][0]
        if accent == "good":
            return STATUS_COLORS["normal"][0]
        return "#1c1917"  # stone-900


class DataTable(QTableView):
    """QTableView di atas PandasTableModel.

    Fitur:
      - Sort dengan klik header (numeric-aware)
      - Alternating row colors
      - Tooltip sel = teks penuh
      - Selection behavior: select rows
      - Header word-wrap + resize mode interactive
    """

    def __init__(self, parent: QWidget | None = None) -> None:
        super().__init__(parent)
        self.setAlternatingRowColors(True)
        self.setSelectionBehavior(QTableView.SelectRows)
        self.setSelectionMode(QTableView.SingleSelection)
        self.setEditTriggers(QTableView.NoEditTriggers)
        self.setHorizontalScrollMode(QTableView.ScrollPerPixel)
        self.verticalHeader().setDefaultSectionSize(24)
        self.verticalHeader().setVisible(False)
        h = self.horizontalHeader()
        h.setSectionResizeMode(QHeaderView.Interactive)
        h.setStretchLastSection(False)
        h.setSortIndicatorShown(True)
        h.setSectionsClickable(True)

    def set_dataframe(self, df) -> None:
        """Pasang DataFrame baru ke tabel."""
        model = PandasTableModel(df)
        self.setModel(model)
        # Auto-fit column widths based on header + sample
        for c in range(model.columnCount()):
            self.resizeColumnToContents(c)
            # Cap width
            if self.columnWidth(c) > 200:
                self.setColumnWidth(c, 200)
            elif self.columnWidth(c) < 80:
                self.setColumnWidth(c, 80)
