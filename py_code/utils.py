from __future__ import annotations

import csv
import re
from pathlib import Path
from typing import Any

import pandas as pd

def normalize_space(text: str | None) -> str:
    if text is None:
        return ""
    return re.sub(r"\s+", " ", str(text)).strip()

def normalize_key(text: str | None) -> str:
    text = normalize_space(text).upper()
    text = re.sub(r"[^A-Z0-9 ]+", " ", text)
    text = re.sub(r"\s+", " ", text).strip()
    return text

def normalize_invoice_no(text: str | None) -> str:
    text = normalize_space(text).upper()
    text = re.sub(r"[\s\-_/\\]+", "", text)
    text = re.sub(r"[^A-Z0-9]", "", text)
    return text

def safe_float(value: Any, default: float = 0.0) -> float:
    if value is None:
        return default
    if isinstance(value, (int, float)) and not pd.isna(value):
        return float(value)
    s = normalize_space(value)
    if not s or s in {"-", "NA", "N/A"}:
        return default
    s = s.replace(",", "")
    try:
        return float(s)
    except Exception:
        return default

def detect_file_kind(path: Path) -> str:
    name = path.name.lower()
    if "gstr1" in name:
        return "gstr1"
    if "gstr2b" in name:
        return "gstr2b"
    if "gstr3b" in name:
        return "gstr3b"
    if "gstr2a" in name:
        return "gstr2a"
    if "credit" in name:
        return "electronic_credit_ledger"
    if "cash" in name:
        return "electronic_cash_ledger"
    if "liability" in name:
        return "electronic_liability_ledger"
    return "unknown"

def read_csv_safely(path: Path) -> pd.DataFrame:
    # Handles ragged CSV exports from GST ledgers.
    try:
        return pd.read_csv(path, engine="python", on_bad_lines="skip")
    except Exception:
        return pd.read_csv(path, engine="python", header=None, on_bad_lines="skip")

def to_month_label(text: str | None) -> str:
    s = normalize_space(text)
    if not s:
        return ""
    return s.replace("'", "").replace("–", "-")
