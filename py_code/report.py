from __future__ import annotations

from pathlib import Path
from typing import Any

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.worksheet import Worksheet

from . import rules

# ---------------------------------------------------------------------------
# Styling
# ---------------------------------------------------------------------------

NAVY = "1F3864"
LIGHT_BLUE = "DCE6F1"
GREEN = "C6EFCE"
RED = "FFC7CE"
AMBER = "FFEB9C"
GREY = "F2F2F2"
WHITE = "FFFFFF"

TITLE_FONT = Font(name="Calibri", size=16, bold=True, color=WHITE)
SECTION_FONT = Font(name="Calibri", size=12, bold=True, color=NAVY)
HEADER_FONT = Font(name="Calibri", size=10, bold=True, color=WHITE)
BODY_FONT = Font(name="Calibri", size=10)
BOLD_BODY = Font(name="Calibri", size=10, bold=True)
NOTE_FONT = Font(name="Calibri", size=10, italic=True, color="808080")

HEADER_FILL = PatternFill("solid", fgColor=NAVY)
TITLE_FILL = PatternFill("solid", fgColor=NAVY)
STRIPE_FILL = PatternFill("solid", fgColor=GREY)
STATUS_FILL = {
    "Matched": PatternFill("solid", fgColor=GREEN),
    "Mismatch": PatternFill("solid", fgColor=RED),
    "Expected variance": PatternFill("solid", fgColor=AMBER),
    "Informational": PatternFill("solid", fgColor=LIGHT_BLUE),
}
THIN = Side(style="thin", color="BFBFBF")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
CURRENCY = "#,##0.00"


def _sheet(wb: Workbook, name: str) -> Worksheet:
    ws = wb.create_sheet(name[:31])
    ws.sheet_view.showGridLines = False
    return ws


def _title(ws: Worksheet, text: str, span: int = 8) -> int:
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=span)
    cell = ws.cell(row=1, column=1, value=text)
    cell.font = TITLE_FONT
    cell.fill = TITLE_FILL
    cell.alignment = Alignment(horizontal="left", vertical="center", indent=1)
    ws.row_dimensions[1].height = 28
    return 3  # next free row


def _section(ws: Worksheet, row: int, text: str, span: int = 8) -> int:
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=span)
    cell = ws.cell(row=row, column=1, value=text)
    cell.font = SECTION_FONT
    cell.alignment = Alignment(horizontal="left", vertical="center")
    return row + 1


def _header_row(ws: Worksheet, row: int, headers: list[str], widths: list[int] | None = None) -> int:
    for i, h in enumerate(headers, start=1):
        c = ws.cell(row=row, column=i, value=h)
        c.font = HEADER_FONT
        c.fill = HEADER_FILL
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        c.border = BORDER
    if widths:
        for i, w in enumerate(widths, start=1):
            ws.column_dimensions[get_column_letter(i)].width = w
    ws.row_dimensions[row].height = 30
    ws.freeze_panes = ws.cell(row=row + 1, column=1)
    return row + 1


def _data_row(ws: Worksheet, row: int, values: list[Any], currency_cols: set[int] | None = None, stripe: bool = False, status_col: int | None = None) -> None:
    currency_cols = currency_cols or set()
    for i, v in enumerate(values, start=1):
        c = ws.cell(row=row, column=i, value=v)
        c.font = BODY_FONT
        c.border = BORDER
        c.alignment = Alignment(vertical="center", wrap_text=False)
        if i in currency_cols and isinstance(v, (int, float)):
            c.number_format = CURRENCY
        if stripe and (status_col is None or i != status_col):
            c.fill = STRIPE_FILL
    if status_col:
        status_val = values[status_col - 1]
        fill = STATUS_FILL.get(status_val)
        if fill:
            ws.cell(row=row, column=status_col).fill = fill
            ws.cell(row=row, column=status_col).font = BOLD_BODY


def _note(ws: Worksheet, row: int, text: str, span: int = 8) -> int:
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=span)
    c = ws.cell(row=row, column=1, value=text)
    c.font = NOTE_FONT
    c.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
    return row + 1


def _fmt_period(pk: str) -> str:
    if not pk or pk == "UNKNOWN":
        return "Unknown period"
    try:
        y, m = pk.split("-")
        months = ["", "Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]
        return f"{months[int(m)]} {y}"
    except Exception:
        return pk


# ---------------------------------------------------------------------------
# Executive Summary
# ---------------------------------------------------------------------------

def _build_executive_summary(wb: Workbook, payload: dict[str, Any]) -> None:
    ws = _sheet(wb, "Executive Summary")
    row = _title(ws, "GST Reconciliation Report - Executive Summary")
    gstin = payload.get("gstin") or "Not detected"
    row = _note(ws, row, f"GSTIN: {gstin}   |   Tolerance applied: Rs. {payload.get('tolerance', 0.5):.2f}   |   "
                          f"Periods covered: {', '.join(_fmt_period(p['period_key']) for p in payload['periods'])}")
    row += 1

    for p in payload["periods"]:
        row = _section(ws, row, f"Period: {_fmt_period(p['period_key'])}")
        g1, g3b, g2b = p.get("gstr1"), p.get("gstr3b"), p.get("gstr2b")

        headers = ["Metric", "Value"]
        row = _header_row(ws, row, headers, widths=[46, 22])
        rows_data = []
        if g1:
            rows_data.append(("GSTR-1 ARN / Date", f"{g1.get('arn','-')}  ({g1.get('arn_date','-')})"))
        if g3b:
            rows_data.append(("GSTR-3B ARN / Date", f"{g3b.get('arn','-')}  ({g3b.get('arn_date','-')})"))
            rows_data.append(("Outward taxable value (3.1(a)+(b))", round(g3b.get("outward_taxable", 0) + g3b.get("zero_rated_taxable", 0), 2)))
            rows_data.append(("Output tax (IGST+CGST+SGST)", round(g3b.get("outward_igst", 0) + g3b.get("outward_cgst", 0) + g3b.get("outward_sgst", 0), 2)))
            rows_data.append(("Net ITC available (4C, IGST+CGST+SGST)", round(g3b.get("net_itc_igst", 0) + g3b.get("net_itc_cgst", 0) + g3b.get("net_itc_sgst", 0), 2)))
            rows_data.append(("Ineligible/restricted ITC (4(D)(2))", round(g3b.get("restricted_cgst", 0) + g3b.get("restricted_sgst", 0) + g3b.get("restricted_igst", 0), 2)))
            rows_data.append(("RCM tax payable (3.1(d))", round(g3b.get("rcm_igst", 0) + g3b.get("rcm_cgst", 0) + g3b.get("rcm_sgst", 0), 2)))
        if g2b:
            rows_data.append(("GSTR-2B invoices (B2B)", len([r for r in g2b.get("detail", []) if r.get("sheet") == "B2B"])))
            rows_data.append(("GSTR-2B uncommon entries (CDNR/DNR/ISD/Import)", len(g2b.get("uncommon", []))))

        outward_rows = p["outward_reco"].get("rows", [])
        itc_rows = p["itc_reco"].get("rows", [])
        matched_o = sum(1 for r in outward_rows if r["status"] == "Matched")
        matched_i = sum(1 for r in itc_rows if r["status"] == "Matched")
        rows_data.append(("Outward reco: matched / total checks", f"{matched_o} / {len(outward_rows)}"))
        rows_data.append(("ITC reco: matched / total checks", f"{matched_i} / {len(itc_rows)}"))
        rows_data.append(("Exception register entries", len(p["exceptions"])))
        rows_data.append(("Duplicate invoice groups flagged", len(p["duplicate_invoices"])))
        if p["missing_returns"]:
            rows_data.append(("Missing returns for this period", ", ".join(p["missing_returns"])))

        for i, (label, val) in enumerate(rows_data):
            currency_cols = {2} if isinstance(val, (int, float)) else set()
            _data_row(ws, row, [label, val], currency_cols=currency_cols, stripe=(i % 2 == 0))
            row += 1
        row += 1

    if payload.get("data_quality_notes"):
        row = _section(ws, row, "Data Quality Flags")
        row = _note(ws, row, f"{len(payload['data_quality_notes'])} item(s) - see the 'Data Quality' sheet for the full list and suggested fix.")
        row += 1

    ws.column_dimensions["A"].width = 46
    ws.column_dimensions["B"].width = 26


# ---------------------------------------------------------------------------
# Outward Reconciliation
# ---------------------------------------------------------------------------

def _build_outward_sheet(wb: Workbook, payload: dict[str, Any]) -> None:
    ws = _sheet(wb, "Outward Reconciliation")
    row = _title(ws, "Outward Supply Reconciliation - GSTR-1 vs GSTR-3B")
    row = _note(ws, row, "Compares outward-supply totals as declared in GSTR-1 against the liability declared in GSTR-3B for the same period.")
    row += 1

    headers = ["Period", "Check", "GSTR-1 (Rs.)", "GSTR-3B (Rs.)", "Difference (Rs.)", "Status", "Remark"]
    row = _header_row(ws, row, headers, widths=[12, 44, 16, 16, 14, 14, 55])

    any_row = False
    for p in payload["periods"]:
        oreco = p["outward_reco"]
        if not oreco.get("available"):
            _data_row(ws, row, [_fmt_period(p["period_key"]), oreco.get("note", "Not available"), "", "", "", "", ""])
            row += 1
            continue
        for i, r in enumerate(oreco["rows"]):
            any_row = True
            _data_row(
                ws, row,
                [_fmt_period(p["period_key"]), r["label"], r["value_1"], r["value_2"], r["diff"], r["status"], r.get("remark", "")],
                currency_cols={3, 4, 5}, stripe=(i % 2 == 0), status_col=6,
            )
            row += 1
    if not any_row:
        _note(ws, row, "No outward reconciliation could be computed - GSTR-1 and/or GSTR-3B were not found for any period.")


# ---------------------------------------------------------------------------
# ITC Reconciliation
# ---------------------------------------------------------------------------

def _build_itc_sheet(wb: Workbook, payload: dict[str, Any]) -> None:
    ws = _sheet(wb, "ITC Reconciliation")
    row = _title(ws, "ITC Reconciliation - GSTR-3B (4A) vs GSTR-2B")
    row = _note(ws, row, "Compares each ITC bucket in GSTR-3B Table 4(A) against the matching GSTR-2B summary section, tax-head by tax-head. "
                          "'Expected variance' rows are self-assessed RCM tax that GSTR-2B does not carry - not an error.")
    row += 1

    headers = ["Period", "Check", "GSTR-3B (Rs.)", "GSTR-2B (Rs.)", "Difference (Rs.)", "Status", "Remark"]
    row = _header_row(ws, row, headers, widths=[12, 48, 16, 16, 14, 16, 60])

    any_row = False
    for p in payload["periods"]:
        ireco = p["itc_reco"]
        if not ireco.get("available"):
            _data_row(ws, row, [_fmt_period(p["period_key"]), ireco.get("note", "Not available"), "", "", "", "", ""])
            row += 1
            continue
        for i, r in enumerate(ireco["rows"]):
            any_row = True
            v2 = r["value_2"] if r["value_2"] is not None else ""
            diff = r["diff"] if r["diff"] is not None else ""
            _data_row(
                ws, row,
                [_fmt_period(p["period_key"]), r["label"], r["value_1"], v2, diff, r["status"], r.get("remark", "")],
                currency_cols={3, 4, 5}, stripe=(i % 2 == 0), status_col=6,
            )
            row += 1
    if not any_row:
        _note(ws, row, "No ITC reconciliation could be computed - GSTR-3B and/or GSTR-2B were not found for any period.")


# ---------------------------------------------------------------------------
# Ledger Summary
# ---------------------------------------------------------------------------

def _ledger_period_totals(merged: dict[str, Any]) -> dict[str, dict[str, float]]:
    totals: dict[str, dict[str, float]] = {}
    for tx in merged.get("transactions", []):
        pk = tx.get("period_key") or "UNKNOWN"
        t = totals.setdefault(pk, {"credit": 0.0, "debit": 0.0, "credit_n": 0, "debit_n": 0})
        amt = tx.get("igst", 0.0) + tx.get("cgst", 0.0) + tx.get("sgst", 0.0) + tx.get("cess", 0.0)
        if tx.get("txn_type") == "Credit":
            t["credit"] += amt
            t["credit_n"] += 1
        else:
            t["debit"] += amt
            t["debit_n"] += 1
    return totals


def _build_ledger_sheet(wb: Workbook, payload: dict[str, Any]) -> None:
    ws = _sheet(wb, "Ledger Summary")
    row = _title(ws, "Electronic Ledger Summary")
    row = _note(ws, row, "Period-wise credits/debits from the Cash, Credit and Liability ledgers, after de-duplicating any overlapping exports.")
    row += 1

    for label, key in [("Electronic Cash Ledger", "merged_cash"), ("Electronic Credit Ledger", "merged_credit"), ("Electronic Liability Register", "merged_liability")]:
        merged = payload[key]
        row = _section(ws, row, label)
        n_files = len(merged.get("source_files", []))
        n_dupe = merged.get("duplicates_dropped", 0)
        row = _note(ws, row, f"Source file(s): {', '.join(merged.get('source_files', [])) or 'none found'}   |   Duplicate rows dropped: {n_dupe}")

        headers = ["Period", "Credit entries", "Total credited (Rs.)", "Debit entries", "Total debited (Rs.)"]
        row = _header_row(ws, row, headers, widths=[14, 16, 20, 16, 20])
        totals = _ledger_period_totals(merged)
        if not totals:
            _note(ws, row, "No transactions parsed for this ledger.")
            row += 2
            continue
        for i, pk in enumerate(sorted(totals.keys())):
            t = totals[pk]
            _data_row(
                ws, row,
                [_fmt_period(pk), t["credit_n"], round(t["credit"], 2), t["debit_n"], round(t["debit"], 2)],
                currency_cols={3, 5}, stripe=(i % 2 == 0),
            )
            row += 1
        row += 1

    row = _section(ws, row, "Cross-check against GSTR-3B")
    headers = ["Period", "Check", "Ledger (Rs.)", "GSTR-3B (Rs.)", "Difference (Rs.)", "Status"]
    row = _header_row(ws, row, headers, widths=[12, 50, 16, 16, 14, 14])
    any_row = False
    for p in payload["periods"]:
        for i, r in enumerate(p["ledger_reco"].get("rows", [])):
            any_row = True
            _data_row(
                ws, row,
                [_fmt_period(p["period_key"]), r["label"], r["value_1"], r["value_2"], r["diff"], r["status"]],
                currency_cols={3, 4, 5}, stripe=(i % 2 == 0), status_col=6,
            )
            row += 1
    if not any_row:
        _note(ws, row, "No ledger cross-check available (GSTR-3B or ledger data missing for all periods).")


# ---------------------------------------------------------------------------
# Exception Register (Annexure)
# ---------------------------------------------------------------------------

def _exception_row_fields(r: dict[str, Any]) -> list[Any]:
    sheet = r.get("sheet", "")
    if sheet in {"B2B"}:
        party = r.get("party_name", "")
        doc = r.get("invoice_no", "")
        doc_date = r.get("invoice_date", "")
        taxable = r.get("taxable_value", 0.0)
        tax = r.get("igst", 0.0) + r.get("cgst", 0.0) + r.get("sgst", 0.0)
        detail = f"RCM: {r.get('rcm','-')}  |  ITC available: {r.get('itc_availability','-')}  |  Reason: {r.get('reason','-')}"
    elif sheet in {"B2B-CDNR", "B2B-CDNRA", "B2B-DNR", "B2B-DNRA"}:
        party = r.get("party_name", "")
        doc = r.get("invoice_no", "")
        doc_date = r.get("invoice_date", "")
        taxable = r.get("taxable_value", 0.0)
        tax = r.get("igst", 0.0) + r.get("cgst", 0.0) + r.get("sgst", 0.0)
        detail = f"{r.get('invoice_type','Note')}  |  ITC available: {r.get('itc_availability','-')}"
    elif sheet in {"ISD", "ISDA", "ISD(Rejected)", "ISDA(Rejected)"}:
        party = r.get("party_name", "")
        doc = r.get("doc_no", "")
        doc_date = r.get("doc_date", "")
        taxable = 0.0
        tax = r.get("igst", 0.0) + r.get("cgst", 0.0) + r.get("sgst", 0.0)
        detail = f"ISD document  |  Eligibility: {r.get('eligibility','-')}"
    elif sheet in {"IMPG", "IMPGA"}:
        party = "Customs (Bill of Entry)"
        doc = r.get("boe_no", "")
        doc_date = r.get("boe_date", "")
        taxable = r.get("taxable_value", 0.0)
        tax = r.get("igst", 0.0)
        detail = f"Port: {r.get('port_code','-')}"
    elif sheet in {"IMPGSEZ", "IMPGSEZA"}:
        party = r.get("party_name", "SEZ supplier")
        doc = r.get("boe_no", "")
        doc_date = r.get("boe_date", "")
        taxable = r.get("taxable_value", 0.0)
        tax = r.get("igst", 0.0)
        detail = f"Port: {r.get('port_code','-')}"
    else:
        party = r.get("party_name", "")
        doc = r.get("invoice_no") or r.get("doc_no") or r.get("boe_no", "")
        doc_date = r.get("invoice_date") or r.get("doc_date") or r.get("boe_date", "")
        taxable = r.get("taxable_value", 0.0)
        tax = r.get("igst", 0.0) + r.get("cgst", 0.0) + r.get("sgst", 0.0)
        detail = ""
    return [sheet, party, doc, doc_date, round(taxable, 2), round(tax, 2), detail]


def _build_exception_sheet(wb: Workbook, payload: dict[str, Any]) -> None:
    ws = _sheet(wb, "Exception Register")
    row = _title(ws, "Exception Register - Uncommon GSTR-2B Entries (Annexure)")
    row = _note(ws, row, "Deliberately excludes the routine B2B feed. Only RCM invoices, POS/eligibility-restricted ITC, credit/debit notes, "
                          "ISD credits and import-of-goods entries are listed here - the transactions a CA actually needs to review.")
    row += 1

    headers = ["Period", "Category", "GSTR-2B sheet", "Party / Source", "Document no.", "Document date", "Taxable value (Rs.)", "Tax amount (Rs.)", "Detail"]
    row = _header_row(ws, row, headers, widths=[12, 26, 14, 30, 16, 14, 18, 16, 48])

    any_row = False
    for p in payload["periods"]:
        for i, r in enumerate(p["exceptions"]):
            any_row = True
            fields = _exception_row_fields(r)
            _data_row(
                ws, row,
                [_fmt_period(p["period_key"]), r["category"]] + fields,
                currency_cols={7, 8}, stripe=(i % 2 == 0),
            )
            row += 1
    if not any_row:
        _note(ws, row, "No uncommon entries found - GSTR-2B for the covered period(s) contains only routine B2B invoices.")
        row += 1

    if any(p["duplicate_invoices"] for p in payload["periods"]):
        row += 1
        row = _section(ws, row, "Possible Duplicate Invoices (same supplier GSTIN + invoice number, appears more than once)")
        headers = ["Period", "Supplier GSTIN", "Invoice No.", "Occurrences"]
        row = _header_row(ws, row, headers, widths=[12, 20, 20, 14])
        for p in payload["periods"]:
            for i, d in enumerate(p["duplicate_invoices"]):
                _data_row(ws, row, [_fmt_period(p["period_key"]), d["gstin_supplier"], d["invoice_no"], d["occurrences"]], stripe=(i % 2 == 0))
                row += 1


# ---------------------------------------------------------------------------
# Observations (Auditor's Note)
# ---------------------------------------------------------------------------

def _build_observations_sheet(wb: Workbook, payload: dict[str, Any]) -> None:
    ws = _sheet(wb, "Observations")
    row = _title(ws, "Observations & Anomalies - Auditor's Note")
    row = _note(ws, row, "Auto-generated interpretation of the reconciliation results below. Written to be read directly, not as raw data.")
    row += 1

    for p in payload["periods"]:
        row = _section(ws, row, f"Period: {_fmt_period(p['period_key'])}")
        for note in p["observations"]:
            ws.cell(row=row, column=1, value="\u2022")
            c = ws.cell(row=row, column=2, value=note)
            c.font = BODY_FONT
            c.alignment = Alignment(horizontal="left", vertical="top", wrap_text=True)
            ws.merge_cells(start_row=row, start_column=2, end_row=row, end_column=8)
            ws.row_dimensions[row].height = 30
            row += 1
        row += 1

    ws.column_dimensions["A"].width = 3
    for col in "BCDEFGH":
        ws.column_dimensions[col].width = 14
    ws.column_dimensions["B"].width = 110


# ---------------------------------------------------------------------------
# Data Quality & Action Items
# ---------------------------------------------------------------------------

def _build_data_quality_sheet(wb: Workbook, payload: dict[str, Any]) -> None:
    ws = _sheet(wb, "Data Quality")
    row = _title(ws, "Data Quality & Action Items")
    row = _note(ws, row, "Anything below means part of the reconciliation is incomplete. Fix the input folder and re-run to close the gap.")
    row += 1

    headers = ["#", "Issue"]
    row = _header_row(ws, row, headers, widths=[6, 120])
    notes = payload.get("data_quality_notes", [])
    if not notes:
        _note(ws, row, "No data quality issues detected - all expected files were found and parsed cleanly.")
        return
    for i, n in enumerate(notes, start=1):
        _data_row(ws, row, [i, n], stripe=(i % 2 == 0))
        ws.row_dimensions[row].height = 28
        ws.cell(row=row, column=2).alignment = Alignment(wrap_text=True, vertical="center")
        row += 1


# ---------------------------------------------------------------------------
# Reason code reference (only codes actually used this run)
# ---------------------------------------------------------------------------

def _build_reference_sheet(wb: Workbook, payload: dict[str, Any]) -> None:
    used_codes = set()
    for p in payload["periods"]:
        for r in p["outward_reco"].get("rows", []):
            used_codes.add(r.get("reason_code"))
        for r in p["itc_reco"].get("rows", []):
            used_codes.add(r.get("reason_code"))
    used_codes.discard(None)
    used_codes.discard("MATCHED")
    if not used_codes:
        return

    ws = _sheet(wb, "Reason Codes")
    row = _title(ws, "Reason Code Reference")
    row = _note(ws, row, "Only the codes that were actually triggered in this report are listed.")
    row += 1
    headers = ["Code", "Title", "Explanation", "Suggested action", "Severity"]
    row = _header_row(ws, row, headers, widths=[22, 26, 60, 40, 12])
    catalog = {c.code: c for c in rules.mismatch_catalog()}
    for i, code in enumerate(sorted(used_codes)):
        rc = catalog.get(code)
        if not rc:
            continue
        _data_row(ws, row, [rc.code, rc.title, rc.explanation, rc.action, rc.severity], stripe=(i % 2 == 0))
        ws.row_dimensions[row].height = 45
        for col in (3, 4):
            ws.cell(row=row, column=col).alignment = Alignment(wrap_text=True, vertical="center")
        row += 1


# ---------------------------------------------------------------------------
# Entry point
# ---------------------------------------------------------------------------

def build_report(payload: dict[str, Any], output_path: Path) -> None:
    wb = Workbook()
    wb.remove(wb.active)

    _build_executive_summary(wb, payload)
    _build_observations_sheet(wb, payload)
    _build_outward_sheet(wb, payload)
    _build_itc_sheet(wb, payload)
    _build_ledger_sheet(wb, payload)
    _build_exception_sheet(wb, payload)
    _build_data_quality_sheet(wb, payload)
    _build_reference_sheet(wb, payload)

    wb.save(output_path)
