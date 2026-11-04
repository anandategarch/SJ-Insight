"""models.py — QAbstractTableModel di atas pandas DataFrame (mirror Waste models).

Keunggulan:
  - TANPA cap baris (virtualized)
  - Sort numeric-aware (NaN di akhir)
  - Pewarnaan baris per severity
  - Tooltip sel = teks penuh (kolom panjang tidak pernah 'hilang')
"""
from __future__ import annotations

import logging

import numpy as np
import pandas as pd
from PySide6.QtCore import QAbstractTableModel, QModelIndex, Qt
from PySide6.QtGui import QColor

from sj_anomaly.utils import severity_of_status

from .theme import STATUS_COLORS

_LOG = logging.getLogger("sj_anomaly.ui")

# Kolom yang ditampilkan sebagai PERSEN (nilai 0..1)
PCT_COLS = frozenset({
    "Deviasi QTY Terkecil 2 %",
    "Dominasi Standar QTY Terkecil 2 %",
})

# Kolom yang ditampilkan sebagai rasio 'Nx' (1.11×)
RATIO_X_COLS = frozenset({"Rasio terhadap Standar QTY Terkecil 2"})

# Kolom numerik (sort numeric-aware)
NUM_COLS = {
    "QTY Terkecil", "Kuantitas", "Selisih Tgl-Tgl Create",
    "Jumlah Variasi Satuan", "Rata2 QTY Resto", "Jumlah Transaksi Resto",
    "QTY Terkecil 2", "Standar QTY Terkecil 2", "Frekuensi Standar QTY Terkecil 2",
    "Jumlah Transaksi Valid Ratio", "Jumlah Variasi QTY Terkecil 2",
    "Selisih QTY Terkecil 2", "Selisih Hari",
}

# Warna negatif (mirror Waste)
_NEG_FG = QColor("#dc2626")


class PandasTableModel(QAbstractTableModel):
    """Model tabel di atas pandas DataFrame.

    Sort: klik header → sort di DataFrame sumber (numeric-aware, NaN di akhir).
    Pewarnaan: kolom status berdasarkan severity; angka minus merah global.
    """

    def __init__(self, df: pd.DataFrame | None = None, parent=None) -> None:
        super().__init__(parent)
        self._df: pd.DataFrame = df.copy() if df is not None else pd.DataFrame()
        self._status_columns: list[str] = []
        self._compute_status_columns()

    def _compute_status_columns(self) -> None:
        """Deteksi kolom yang ber-status (diwarnai per severity)."""
        self._status_columns = [
            c for c in self._df.columns
            if isinstance(c, str) and (
                c.startswith("Status Anomali")
                or c == "Memiliki Dokumen [Terima]"
            )
        ]

    def set_dataframe(self, df: pd.DataFrame) -> None:
        """Ganti DataFrame sumber + reset model."""
        self.beginResetModel()
        self._df = df.copy()
        self._compute_status_columns()
        self.endResetModel()

    @property
    def df(self) -> pd.DataFrame:
        return self._df

    # ------------------------------------------------------------
    # QAbstractTableModel contract
    # ------------------------------------------------------------
    def rowCount(self, parent: QModelIndex = QModelIndex()) -> int:
        return 0 if parent.isValid() else len(self._df)

    def columnCount(self, parent: QModelIndex = QModelIndex()) -> int:
        return 0 if parent.isValid() else len(self._df.columns)

    def headerData(self, section: int, orientation: Qt.Orientation, role: int = Qt.DisplayRole):
        if role != Qt.DisplayRole:
            return None
        if orientation == Qt.Horizontal:
            return str(self._df.columns[section]) if section < len(self._df.columns) else ""
        return str(section + 1)

    def data(self, index: QModelIndex, role: int = Qt.DisplayRole):
        if not index.isValid():
            return None
        row, col = index.row(), index.column()
        if row >= len(self._df) or col >= len(self._df.columns):
            return None
        col_name = self._df.columns[col]
        value = self._df.iloc[row, col]

        if role == Qt.DisplayRole:
            return _format_cell(value, col_name)
        if role == Qt.ToolTipRole:
            # Tooltip = teks penuh (kolom panjang tidak terpotong)
            return _format_cell(value, col_name)
        if role == Qt.BackgroundRole or role == Qt.ForegroundRole:
            return _color_for_cell(value, col_name, role)
        if role == Qt.TextAlignmentRole:
            # Kolom numerik rata kanan
            if col_name in NUM_COLS or col_name in PCT_COLS or col_name in RATIO_X_COLS:
                return int(Qt.AlignRight | Qt.AlignVCenter)
            return int(Qt.AlignLeft | Qt.AlignVCenter)
        return None

    def sort(self, column: int, order: Qt.SortOrder = Qt.AscendingOrder) -> None:
        """Sort di DataFrame sumber (numeric-aware, NaN di akhir)."""
        if column < 0 or column >= len(self._df.columns):
            return
        col_name = self._df.columns[column]
        ascending = (order == Qt.AscendingOrder)
        # NaN selalu di akhir terlepas dari arah sort
        try:
            if col_name in NUM_COLS or pd.api.types.is_numeric_dtype(self._df[col_name]):
                # Coerce ke numeric untuk sort konsisten
                sorted_series = pd.to_numeric(self._df[col_name], errors="coerce")
                na_mask = sorted_series.isna()
                sorted_idx = sorted_series.sort_values(ascending=ascending, na_position="last").index
            else:
                sorted_idx = self._df[col_name].sort_values(ascending=ascending, na_position="last").index
            new_df = self._df.loc[sorted_idx].reset_index(drop=True)
        except Exception as e:
            _LOG.warning("Sort gagal pada kolom '%s': %s", col_name, e)
            return
        self.beginResetModel()
        self._df = new_df
        self.endResetModel()


# ============================================================
# FORMAT SEL
# ============================================================

def _format_cell(value, col_name: str) -> str:
    """Format nilai sel untuk display id-ID."""
    if value is None:
        return ""
    if isinstance(value, float):
        if value != value:  # NaN
            return ""
        if col_name in PCT_COLS:
            return f"{value * 100:.2f}%".replace(".", ",")
        if col_name in RATIO_X_COLS:
            return f"{value:.4f}×".replace(".", ",")
        # Default: 2 desimal max, buang trailing zero
        if value.is_integer():
            return f"{int(value):,}".replace(",", ".")
        s = f"{value:,.2f}"
        return s.replace(",", "X").replace(".", ",").replace("X", ".")
    if isinstance(value, int):
        return f"{value:,}".replace(",", ".")
    if isinstance(value, bool):
        return "YA" if value else "TIDAK"
    return str(value)


def _color_for_cell(value, col_name: str, role: int):
    """Background/foreground berdasarkan severity & nilai negatif."""
    # Negatif angka → merah
    if isinstance(value, (int, float)) and not (isinstance(value, float) and value != value):
        if value < 0 and role == Qt.ForegroundRole:
            return _NEG_FG
    # Kolom status → severity color
    if col_name == "Memiliki Dokumen [Terima]":
        if isinstance(value, bool):
            if not value:
                fg, bg = STATUS_COLORS["anomali"]
            else:
                fg, bg = STATUS_COLORS["normal"]
            if role == Qt.BackgroundRole:
                return QColor(bg)
            if role == Qt.ForegroundRole:
                return QColor(fg)
        return None
    if col_name.startswith("Status Anomali"):
        s = str(value or "")
        if s == "NORMAL" or not s:
            return None
        sev = severity_of_status(s)
        fg, bg = STATUS_COLORS.get(sev, STATUS_COLORS["perhatian"])
        if role == Qt.BackgroundRole:
            return QColor(bg)
        if role == Qt.ForegroundRole:
            return QColor(fg)
    return None
