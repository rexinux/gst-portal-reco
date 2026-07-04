from __future__ import annotations

import argparse
from pathlib import Path
from typing import Any

try:
    from .config import Config
    from .parsers import parse_gstr1, parse_gstr2b, parse_gstr3b, parse_ledger_csv
    from .reconciler import compare_summary
    from .report import build_workbook
    from .utils import detect_file_kind
except ImportError:  # pragma: no cover
    import sys
    from pathlib import Path as _Path
    sys.path.append(str(_Path(__file__).resolve().parent.parent))
    from py_code.config import Config
    from py_code.parsers import parse_gstr1, parse_gstr2b, parse_gstr3b, parse_ledger_csv
    from py_code.reconciler import compare_summary
    from py_code.report import build_workbook
    from py_code.utils import detect_file_kind

def scan_folder(folder: Path) -> list[Path]:
    if not folder.exists():
        return []
    files = []
    for p in folder.rglob("*"):
        if p.is_file() and p.suffix.lower() in {".pdf", ".xlsx", ".xls", ".csv"}:
            files.append(p)
    return sorted(files)

def load_inputs(input_root: Path) -> dict[str, list[Path]]:
    groups = {k: [] for k in [
        "gstr1","gstr2a","gstr2b","gstr3b",
        "electronic_credit_ledger","electronic_cash_ledger","electronic_liability_ledger",
    ]}
    for key in groups:
        groups[key] = scan_folder(input_root / key)
    return groups

def process(input_root: Path, output_path: Path) -> Path:
    groups = load_inputs(input_root)

    payload: dict[str, Any] = {"notes": []}
    ledgers: list[dict[str, Any]] = []

    # Pick latest file in each folder for this build.
    g1_file = groups["gstr1"][-1] if groups["gstr1"] else None
    g3_file = groups["gstr3b"][-1] if groups["gstr3b"] else None
    g2b_file = groups["gstr2b"][-1] if groups["gstr2b"] else None

    if g1_file:
        payload["gstr1"] = parse_gstr1(g1_file)
    else:
        payload["gstr1"] = {}

    if g3_file:
        payload["gstr3b"] = parse_gstr3b(g3_file)
    else:
        payload["gstr3b"] = {}

    if g2b_file:
        payload["gstr2b"] = parse_gstr2b(g2b_file)
    else:
        payload["gstr2b"] = {"summary": [], "detail": []}

    # Optional ledgers
    for key in ["electronic_credit_ledger", "electronic_cash_ledger", "electronic_liability_ledger"]:
        for file in groups[key]:
            led = parse_ledger_csv(file)
            ledgers.append(led)
    payload["ledgers"] = ledgers

    # notes
    payload["notes"].append(("Input root", str(input_root)))
    payload["notes"].append(("GSTR-1 files found", str(len(groups["gstr1"]))))
    payload["notes"].append(("GSTR-2B files found", str(len(groups["gstr2b"]))))
    payload["notes"].append(("GSTR-3B files found", str(len(groups["gstr3b"]))))
    payload["notes"].append(("Ledger files found", str(len(ledgers))))
    if groups["gstr2a"]:
        payload["notes"].append(("GSTR-2A files found", str(len(groups["gstr2a"]))))
    else:
        payload["notes"].append(("GSTR-2A", "Optional folder not used in this run"))

    payload["recon"] = compare_summary(payload.get("gstr1"), payload.get("gstr3b"), payload.get("gstr2b"), ledgers)

    return build_workbook(payload, output_path)

def main():
    parser = argparse.ArgumentParser(description="Offline GST reconciliation engine")
    parser.add_argument("--input-root", type=Path, default=Path("input"))
    parser.add_argument("--output", type=Path, default=Path("output/gst_reconciliation_report.xlsx"))
    args = parser.parse_args()
    out = process(args.input_root, args.output)
    print(f"Created: {out}")

if __name__ == "__main__":
    main()
