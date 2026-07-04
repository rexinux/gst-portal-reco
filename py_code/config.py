from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path

SUPPORTED_EXTENSIONS = {".pdf", ".xlsx", ".xls", ".csv"}

FOLDERS = {
    "gstr1": "gstr1",
    "gstr2a": "gstr2a",
    "gstr2b": "gstr2b",
    "gstr3b": "gstr3b",
    "electronic_credit_ledger": "electronic_credit_ledger",
    "electronic_cash_ledger": "electronic_cash_ledger",
    "electronic_liability_ledger": "electronic_liability_ledger",
}

@dataclass
class Config:
    input_root: Path
    output_path: Path
    tolerance: float = 0.50
    max_months: int = 12
