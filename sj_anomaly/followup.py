"""followup.py — daftar tindak lanjut baris anomali (mirror Waste followup).

Sederhana: simpan list of fingerprints (string) ke JSON, dengan API:
  - load_followup_keys(path) → list[str]
  - save_followup_keys(keys, path) → bool
  - fingerprint_row(row) → str  (sidik jari baris — No.Pemindahan + Nama Barang + Tanggal)
  - select_followup_rows(df, fps, keys) → DataFrame (filter df by fingerprint ∈ keys)
"""
from __future__ import annotations

import json
from pathlib import Path
from typing import Any

import pandas as pd

FP_COL = "__FP"
CHECK_COL = "✓"


def fingerprint_row(row: pd.Series | dict) -> str:
    """Sidik jari baris: No. Pemindahan # + Nama Barang + Tanggal + Satuan + QTY."""
    if isinstance(row, pd.Series):
        get = lambda k: row.get(k, "")
    else:
        get = lambda k: row.get(k, "")
    return "|".join([
        str(get("No. Pemindahan #") or ""),
        str(get("Nama Barang") or ""),
        str(get("Tanggal") or ""),
        str(get("Satuan") or ""),
        str(get("QTY Terkecil") or ""),
    ])


def fingerprint_series(df: pd.DataFrame) -> pd.Series:
    """Bangun Series fingerprint untuk seluruh df (index sama dgn df)."""
    return df.apply(fingerprint_row, axis=1)


def load_followup_keys(path: Path | None) -> list[str]:
    """Load daftar fingerprint dari JSON. Return [] bila file tidak ada / gagal."""
    if path is None or not path.exists():
        return []
    try:
        with open(path, "r", encoding="utf-8") as f:
            data = json.load(f)
        if isinstance(data, list):
            return [str(k) for k in data if k]
        if isinstance(data, dict) and "keys" in data:
            return [str(k) for k in data["keys"] if k]
        return []
    except (OSError, json.JSONDecodeError):
        return []


def save_followup_keys(keys: list[str], path: Path) -> bool:
    """Simpan daftar fingerprint ke JSON. Return True bila sukses."""
    try:
        path.parent.mkdir(parents=True, exist_ok=True)
        with open(path, "w", encoding="utf-8") as f:
            json.dump({"keys": list(keys)}, f, ensure_ascii=False, indent=2)
        return True
    except OSError:
        return False


def select_followup_rows(
    df: pd.DataFrame,
    fps: pd.Series,
    keys: list[str],
) -> pd.DataFrame:
    """Filter df: hanya baris yang fingerprint-nya ada di ``keys``."""
    if df.empty or not keys:
        return df.iloc[:0].copy()
    keyset = set(keys)
    mask = fps.isin(keyset)
    out = df[mask].copy()
    out[FP_COL] = fps[mask]
    return out
