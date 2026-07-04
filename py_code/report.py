from __future__ import annotations

from pathlib import Path
from typing import Any

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill, Border, Side
from openpyxl.utils import get_column_letter

HEADER_FILL = PatternFill("solid", fgColor="1F4E78")
HEADER_FONT = Font(color="FFFFFF", bold=True)
SUB_FILL = PatternFill("solid", fgColor="D9EAF7")
OK_FILL = PatternFill("solid", fgColor="E2F0D9")
BAD_FILL = PatternFill("solid", fgColor="FCE4D6")
WARN_FILL = PatternFill("solid", fgColor="FFF2CC")
THIN = Side(style="thin", color="B7B7B7")

def _style_header(ws, row=1):
    for cell in ws[row]:
        if cell.value is not None:
            cell.fill = HEADER_FILL
            cell.font = HEADER_FONT
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            cell.border = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)

def _auto_width(ws, max_width=42):
    for col in ws.columns:
        values = []
        for cell in col:
            if cell.value is not None:
                values.append(len(str(cell.value)))
        if not values:
            continue
        width = min(max(values) + 2, max_width)
        ws.column_dimensions[get_column_letter(col[0].column)].width = width

def build_workbook(payload: dict[str, Any], output_path: Path) -> Path:
    wb = Workbook()
    ws = wb.active
    ws.title = "Dashboard"

    dash = payload["recon"]["dashboard"]
    warnings = payload["recon"]["warnings"]

    ws["A1"] = "GST Reconciliation Dashboard"
    ws["A1"].font = Font(bold=True, size=14)
    ws.append(["Metric", "Status"])
    _style_header(ws, 2)
    rows = [
        ["Outward turnover match", "Matched" if dash["outward_match"] else "Mismatch"],
        ["IGST ITC match", "Matched" if dash["itc_match_igst"] else "Mismatch"],
        ["CGST ITC match", "Matched" if dash["itc_match_cgst"] else "Mismatch"],
        ["SGST ITC match", "Matched" if dash["itc_match_sgst"] else "Mismatch"],
    ]
    for r in rows:
        ws.append(r)
    ws.append([])
    ws.append(["Notes", ""])
    for w in warnings or ["No major warnings."]:
        ws.append([w, ""])

    for cell in ws["B"][2:6]:
        if cell.value == "Matched":
            cell.fill = OK_FILL
        elif cell.value == "Mismatch":
            cell.fill = BAD_FILL

    # GSTR-1 summary
    _add_key_value_sheet(wb, "GSTR-1 Summary", payload.get("gstr1", {}), [
        "source_file","return_type","period","gstin","arn","arn_date","b2b_invoice_count","b2b_taxable","b2b_igst",
        "b2b_cgst","b2b_sgst","sez_count","sez_taxable","total_liability","outward_total","notes"
    ])

    # GSTR-3B summary
    _add_key_value_sheet(wb, "GSTR-3B Summary", payload.get("gstr3b", {}), [
        "source_file","return_type","period","gstin","arn","arn_date","outward_taxable","zero_rated_taxable",
        "rcm_taxable","rcm_igst","rcm_cgst","rcm_sgst","import_igst","all_other_itc_igst","net_itc_igst",
        "net_itc_cgst","net_itc_sgst","restricted_cgst","restricted_sgst","outward_total","notes"
    ])

    # GSTR-2B summary
    g2b = payload.get("gstr2b", {})
    summary_rows = g2b.get("summary", [])
    ws2 = wb.create_sheet("GSTR-2B Summary")
    ws2.append(["Section", "Subsection", "Heading", "GSTR-3B Table", "IGST", "CGST", "SGST", "Cess", "Advisory"])
    for r in summary_rows:
        ws2.append([
            r.get("section",""), r.get("subsection",""), r.get("heading",""), r.get("gstr3b_table",""),
            r.get("igst",0.0), r.get("cgst",0.0), r.get("sgst",0.0), r.get("cess",0.0), r.get("advisory","")
        ])
    _style_header(ws2, 1)
    _auto_width(ws2)

    # GSTR-2B detail
    ws3 = wb.create_sheet("GSTR-2B Detail")
    detail = g2b.get("detail", [])
    if detail:
        cols = sorted({k for row in detail for k in row.keys()})
        ws3.append(cols)
        _style_header(ws3, 1)
        for row in detail:
            ws3.append([row.get(c, "") for c in cols])
    else:
        ws3.append(["No detail rows parsed"])
    _auto_width(ws3)

    # Ledgers
    ws4 = wb.create_sheet("Electronic Ledgers Summary")
    ws4.append(["Ledger", "File", "Rows", "Columns", "Note"])
    _style_header(ws4, 1)
    for led in payload.get("ledgers", []):
        ws4.append([
            led.get("return_type",""),
            led.get("source_file",""),
            "" if not led.get("shape") else led["shape"][0],
            "" if not led.get("shape") else led["shape"][1],
            led.get("notes",""),
        ])
    _auto_width(ws4)

    # Recon
    ws5 = wb.create_sheet("Reconciliation Summary")
    ws5.append(["Section", "Source 1", "Source 2", "Value 1", "Value 2", "Difference", "Status"])
    _style_header(ws5, 1)
    for r in payload["recon"]["recon_rows"]:
        ws5.append([r["section"], r["source_1"], r["source_2"], r["value_1"], r["value_2"], r["difference"], r["status"]])
    _auto_width(ws5)

    # Exceptions
    ws6 = wb.create_sheet("Exceptions")
    ws6.append(["Type", "Detail"])
    _style_header(ws6, 1)
    if not detail:
        ws6.append(["Info", "No GSTR-2B detail rows found."])
    for r in payload["recon"]["recon_rows"]:
        if r["status"] != "Matched":
            ws6.append(["Mismatch", f'{r["section"]}: {r["difference"]}'])
    if not warnings:
        ws6.append(["Info", "No warnings."])
    else:
        for w in warnings:
            ws6.append(["Warning", w])
    _auto_width(ws6)

    # Notes
    ws7 = wb.create_sheet("Processing Notes")
    ws7.append(["Item", "Value"])
    _style_header(ws7, 1)
    notes = payload.get("notes", [])
    if notes:
        for k, v in notes:
            ws7.append([k, v])
    else:
        ws7.append(["Mode", "Offline folder-based processing"])
        ws7.append(["Limit", "Up to 12 months supported"])
        ws7.append(["GSTR-2A", "Optional"])
        ws7.append(["Invoice matching", "Requires invoice-wise exports"])
    _auto_width(ws7)

    output_path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(output_path)
    return output_path

def _add_key_value_sheet(wb: Workbook, title: str, data: dict[str, Any], keys: list[str]):
    ws = wb.create_sheet(title)
    ws.append(["Field", "Value"])
    _style_header(ws, 1)
    for k in keys:
        if k in data:
            ws.append([k, data.get(k, "")])
    _auto_width(ws)
