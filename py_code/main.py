from __future__ import annotations

import argparse
import sys
from pathlib import Path
from typing import Any

from . import parsers, reconciler, utils
from .config import Config, FOLDERS, SUPPORTED_EXTENSIONS
from .report import build_report


def _files_in(folder: Path) -> list[Path]:
    if not folder.exists():
        return []
    return sorted(
        p for p in folder.iterdir()
        if p.is_file() and p.suffix.lower() in SUPPORTED_EXTENSIONS and not p.name.startswith("~$")
    )


def _parse_folder(folder: Path, parse_fn) -> tuple[list[dict[str, Any]], list[str]]:
    results = []
    errors = []
    for f in _files_in(folder):
        try:
            results.append(parse_fn(f))
        except Exception as exc:  # noqa: BLE001
            errors.append(f"Failed to parse {f.name}: {exc}")
    return results, errors


def _parse_ledger_folder(folder: Path, ledger_type: str) -> tuple[dict[str, Any], list[str]]:
    files = _files_in(folder)
    if not files:
        return {"return_type": ledger_type, "transactions": [], "duplicates_dropped": 0, "file_notes": [], "source_files": []}, []
    parsed = [parsers.parse_ledger_csv(f, ledger_type) for f in files]
    merged = parsers.merge_ledger_files(parsed)
    return merged, []


def run(config: Config) -> dict[str, Any]:
    input_root = config.input_root
    data_quality_notes: list[str] = []

    gstr1_list, err = _parse_folder(input_root / FOLDERS["gstr1"], parsers.parse_gstr1)
    data_quality_notes += err
    gstr3b_list, err = _parse_folder(input_root / FOLDERS["gstr3b"], parsers.parse_gstr3b)
    data_quality_notes += err
    gstr2b_list, err = _parse_folder(input_root / FOLDERS["gstr2b"], parsers.parse_gstr2b)
    data_quality_notes += err

    if not gstr1_list:
        data_quality_notes.append(
            f"No GSTR-1 files found in '{FOLDERS['gstr1']}'. Add the GSTR-1 summary PDF/Excel for each period "
            f"and re-run to get the outward reconciliation."
        )
    if not gstr3b_list:
        data_quality_notes.append(
            f"No GSTR-3B files found in '{FOLDERS['gstr3b']}'. Add the GSTR-3B summary PDF/Excel for each period "
            f"and re-run to get the ITC and outward reconciliation."
        )
    if not gstr2b_list:
        data_quality_notes.append(
            f"No GSTR-2B files found in '{FOLDERS['gstr2b']}'. Add the GSTR-2B Excel export for each period "
            f"and re-run to get the ITC reconciliation and the uncommon-entry annexure."
        )

    gstr1_by_period, c1 = reconciler.index_by_period(gstr1_list)
    gstr3b_by_period, c2 = reconciler.index_by_period(gstr3b_list)
    gstr2b_by_period, c3 = reconciler.index_by_period(gstr2b_list)
    data_quality_notes += c1 + c2 + c3

    merged_credit, e1 = _parse_ledger_folder(input_root / FOLDERS["electronic_credit_ledger"], "electronic_credit_ledger")
    merged_cash, e2 = _parse_ledger_folder(input_root / FOLDERS["electronic_cash_ledger"], "electronic_cash_ledger")
    merged_liability, e3 = _parse_ledger_folder(input_root / FOLDERS["electronic_liability_ledger"], "electronic_liability_ledger")
    data_quality_notes += e1 + e2 + e3
    for merged, label in [(merged_credit, "Electronic Credit Ledger"), (merged_cash, "Electronic Cash Ledger"), (merged_liability, "Electronic Liability Register")]:
        if merged.get("file_notes"):
            data_quality_notes += [f"{label}: {n}" for n in merged["file_notes"]]
        if merged.get("duplicates_dropped"):
            data_quality_notes.append(
                f"{label}: {merged['duplicates_dropped']} duplicate transaction row(s) were found across the uploaded "
                f"export(s) and dropped (same reference no., date, description and amounts) - if that count looks "
                f"too high or too low, double-check which files are sitting in that folder."
            )

    all_periods = sorted(set(gstr1_by_period) | set(gstr3b_by_period) | set(gstr2b_by_period) - {"UNKNOWN"})
    if not all_periods:
        all_periods = ["UNKNOWN"]

    fy_start_year = utils.infer_fy_start_year(all_periods)
    processed_data = reconciler.build_processed_data(
        fy_start_year, gstr1_by_period, gstr3b_by_period, gstr2b_by_period,
        merged_credit, merged_cash, merged_liability,
    )

    period_payloads = []
    for pk in all_periods:
        g1 = gstr1_by_period.get(pk)
        g3b = gstr3b_by_period.get(pk)
        g2b = gstr2b_by_period.get(pk)

        outward_reco = reconciler.reconcile_outward(g1, g3b)
        itc_reco = reconciler.reconcile_itc(g3b, g2b)
        ledger_reco = reconciler.reconcile_ledgers(pk, g3b, merged_credit, merged_cash)
        exceptions = reconciler.build_exception_register(g2b)
        duplicate_invoices = reconciler.find_duplicate_invoices(g2b)
        observations = reconciler.build_observations(
            pk, g1, g3b, g2b, itc_reco, outward_reco, exceptions, duplicate_invoices, ledger_reco
        )

        missing = []
        if not g1:
            missing.append("GSTR-1")
        if not g3b:
            missing.append("GSTR-3B")
        if not g2b:
            missing.append("GSTR-2B")
        if missing:
            data_quality_notes.append(
                f"Period {pk}: missing {', '.join(missing)} - add the file(s) to the relevant input folder and re-run "
                f"for a complete reconciliation of this period."
            )

        period_payloads.append({
            "period_key": pk,
            "gstr1": g1,
            "gstr3b": g3b,
            "gstr2b": g2b,
            "outward_reco": outward_reco,
            "itc_reco": itc_reco,
            "ledger_reco": ledger_reco,
            "exceptions": exceptions,
            "duplicate_invoices": duplicate_invoices,
            "observations": observations,
            "missing_returns": missing,
        })

    payload = {
        "periods": period_payloads,
        "gstin": next((g.get("gstin") for g in gstr1_list if g.get("gstin")), None)
                 or next((g.get("gstin") for g in gstr3b_list if g.get("gstin")), None),
        "merged_credit": merged_credit,
        "merged_cash": merged_cash,
        "merged_liability": merged_liability,
        "data_quality_notes": data_quality_notes,
        "tolerance": config.tolerance,
        "fy_start_year": fy_start_year,
        "processed_data": processed_data,
        "gstr2b_by_period": gstr2b_by_period,
    }
    return payload


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="GST reconciliation report generator")
    parser.add_argument("--input-root", default="input", type=Path)
    parser.add_argument("--output", default="output/gst_reconciliation_report.xlsx", type=Path)
    parser.add_argument("--tolerance", default=0.50, type=float)
    args = parser.parse_args(argv)

    config = Config(input_root=args.input_root, output_path=args.output, tolerance=args.tolerance)
    payload = run(config)

    args.output.parent.mkdir(parents=True, exist_ok=True)
    build_report(payload, args.output)
    print(f"Created: {args.output}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
