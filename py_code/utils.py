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

MONTH_MAP = {
    "JAN": "01", "FEB": "02", "MAR": "03", "APR": "04", "MAY": "05", "JUN": "06",
    "JUL": "07", "AUG": "08", "SEP": "09", "OCT": "10", "NOV": "11", "DEC": "12",
}

def period_key_from_date(date_text: str | None) -> str:
    """Best-effort 'YYYY-MM' key from a dd/mm/yyyy style date string."""
    s = normalize_space(date_text)
    if not s or s == "-":
        return ""
    m = re.match(r"(\d{1,2})[/-](\d{1,2})[/-](\d{2,4})", s)
    if not m:
        return ""
    dd, mm, yy = m.groups()
    yy = yy if len(yy) == 4 else ("20" + yy)
    try:
        mm_i = int(mm)
        if 1 <= mm_i <= 12:
            return f"{yy}-{mm_i:02d}"
    except Exception:
        pass
    return ""

def period_key_from_label(text: str | None) -> str:
    """Best-effort 'YYYY-MM' key from labels like 'Apr-25', 'Apr'25', 'April 2025'."""
    s = to_month_label(text).upper()
    if not s:
        return ""
    m = re.match(r"([A-Z]{3,9})[\s\-]*'?(\d{2,4})", s)
    if not m:
        return ""
    mon_txt, yy = m.groups()
    mon = mon_txt[:3]
    if mon not in MONTH_MAP:
        return ""
    yy = yy if len(yy) == 4 else ("20" + yy)
    return f"{yy}-{MONTH_MAP[mon]}"

def fy_months(fy_start_year: int) -> list[tuple[str, str]]:
    """12 (label, period_key) pairs for an Indian FY: Apr(start_year) through
    Mar(start_year+1). e.g. fy_months(2025) -> [("Apr'25","2025-04"), ..., ("Mar'26","2026-03")]."""
    order = [(4, fy_start_year), (5, fy_start_year), (6, fy_start_year), (7, fy_start_year),
             (8, fy_start_year), (9, fy_start_year), (10, fy_start_year), (11, fy_start_year),
             (12, fy_start_year), (1, fy_start_year + 1), (2, fy_start_year + 1), (3, fy_start_year + 1)]
    names = {1: "Jan", 2: "Feb", 3: "Mar", 4: "Apr", 5: "May", 6: "Jun",
             7: "Jul", 8: "Aug", 9: "Sep", 10: "Oct", 11: "Nov", 12: "Dec"}
    out = []
    for mm, yy in order:
        out.append((f"{names[mm]}'{str(yy)[2:]}", f"{yy}-{mm:02d}"))
    return out


def infer_fy_start_year(period_keys: list[str]) -> int:
    """Given whatever period keys were actually found in the data, infer the
    FY they belong to (Apr-Mar). Falls back to the current FY if nothing
    usable is found."""
    import datetime
    valid = [pk for pk in period_keys if pk and pk != "UNKNOWN" and re.match(r"^\d{4}-\d{2}$", pk)]
    if not valid:
        today = datetime.date.today()
        return today.year if today.month >= 4 else today.year - 1
    y, m = valid[0].split("-")
    y, m = int(y), int(m)
    return y if m >= 4 else y - 1


def row_fingerprint(*parts: Any) -> str:
    """Stable fingerprint for duplicate detection across re-exported files."""
    import hashlib
    norm = "|".join(normalize_space(p).upper() for p in parts)
    return hashlib.sha1(norm.encode("utf-8")).hexdigest()
