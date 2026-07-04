from __future__ import annotations

import csv
import re
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Iterable

import pandas as pd
from openpyxl import load_workbook
from pdfplumber import open as pdfopen

from .utils import (
    detect_file_kind,
    normalize_invoice_no,
    normalize_key,
    normalize_space,
    read_csv_safely,
    safe_float,
    to_month_label,
)

SECTION_SKIP = {
    "FORM GSTR-2B",
    "FORM SUMMARY - ITC AVAILABLE",
    "PART A",
    "PART B",
    "GSTR-3B",
    "GSTIN OF SUPPLIER",
    "GSTIN OF ISD",
    "GSTIN OF THE SUPPLIER",
    "GSTIN",
    "DESCRIPTION",
    "SR.NO.",
    "S.NO.",
    "SR NO.",
    "NO. OF RECORDS",
    "DOCUMENT TYPE",
    "GOODS AND SERVICES TAX - GSTR-2B",
    "TAXABLE INWARD SUPPLIES RECEIVED FROM REGISTERED PERSONS",
    "AMENDMENTS TO PREVIOUSLY FILED INVOICES BY SUPPLIER",
    "CREDIT WHICH MAY BE AVAILED UNDER FORM GSTR-3B",
}

def _text_from_pdf(path: Path) -> str:
    parts: list[str] = []
    with pdfopen(path) as pdf:
        for page in pdf.pages:
            txt = page.extract_text() or ""
            parts.append(txt)
    return "\n".join(parts)

def parse_gstr1(path: Path) -> dict[str, Any]:
    ext = path.suffix.lower()
    if ext == ".pdf":
        text = _text_from_pdf(path)
        return _parse_gstr1_pdf_text(text, path)
    if ext in {".xlsx", ".xls"}:
        return _parse_gstr1_xlsx(path)
    raise ValueError(f"Unsupported GSTR-1 file: {path}")

def _parse_gstr1_pdf_text(text: str, path: Path) -> dict[str, Any]:
    out = {
        "source_file": path.name,
        "return_type": "GSTR-1",
        "period": _first_match(text, r"Tax period\s+([A-Za-z0-9'\- ]+)"),
        "gstin": _first_match(text, r"\bGSTIN\s+([0-9A-Z]{15})\b"),
        "arn": _first_match(text, r"\bARN\s+([A-Z0-9]+)\b"),
        "arn_date": _first_match(text, r"\bARN date\s+([0-9/]+)\b"),
        "b2b_invoice_count": _first_num(text, r"4A.*?B2B Regular\s+Total\s+([0-9,]+)"),
        "b2b_taxable": _first_num(text, r"4A.*?B2B Regular\s+Total\s+[0-9]+\s+Invoice\s+([0-9,]+\.\d+|[0-9,]+)"),
        "b2b_igst": _first_num(text, r"4A.*?B2B Regular.*?Integrated Tax\s*\(₹\)\s*([0-9,]+\.\d+|[0-9,]+)"),
        "b2b_cgst": _first_num(text, r"4A.*?B2B Regular.*?Central Tax\s*\(₹\)\s*([0-9,]+\.\d+|[0-9,]+)"),
        "b2b_sgst": _first_num(text, r"4A.*?B2B Regular.*?State/UT Tax\s*\(₹\)\s*([0-9,]+\.\d+|[0-9,]+)"),
        "sez_count": _first_num(text, r"6B.*?SEZWOP\s+(\d+)"),
        "sez_taxable": _first_num(text, r"6B.*?SEZWOP\s+\d+\s+Invoice\s+([0-9,]+\.\d+|[0-9,]+)"),
        "total_liability": _first_num(text, r"Total Liability.*?([0-9,]+\.\d+|[0-9,]+)\s+([0-9,]+\.\d+|[0-9,]+)\s+([0-9,]+\.\d+|[0-9,]+)\s+([0-9,]+\.\d+|[0-9,]+)"),
    }
    out["outward_total"] = round(out["b2b_taxable"] + out["sez_taxable"], 2)
    out["notes"] = "Summary PDF parsed; invoice-level outward lines not available."
    return out

def _parse_gstr1_xlsx(path: Path) -> dict[str, Any]:
    wb = load_workbook(path, data_only=True)
    sheet = wb[wb.sheetnames[0]]
    rows = list(sheet.iter_rows(values_only=True))
    flat = _rows_to_flat_dict(rows)
    flat["source_file"] = path.name
    flat["return_type"] = "GSTR-1"
    return flat

def parse_gstr3b(path: Path) -> dict[str, Any]:
    ext = path.suffix.lower()
    if ext == ".pdf":
        text = _text_from_pdf(path)
        return _parse_gstr3b_pdf_text(text, path)
    if ext in {".xlsx", ".xls"}:
        return _parse_gstr3b_xlsx(path)
    raise ValueError(f"Unsupported GSTR-3B file: {path}")

def _parse_gstr3b_pdf_text(text: str, path: Path) -> dict[str, Any]:
    out = {
        "source_file": path.name,
        "return_type": "GSTR-3B",
        "period": _first_match(text, r"\bPeriod\s+([A-Za-z0-9'\- ]+)"),
        "gstin": _first_match(text, r"\bGSTIN of the supplier\s+([0-9A-Z]{15})\b"),
        "arn": _first_match(text, r"\bARN\s+([A-Z0-9]+)\b"),
        "arn_date": _first_match(text, r"\bDate of ARN\s+([0-9/]+)\b"),
        "outward_taxable": _first_num(text, r"\(a\)\s+Outward taxable supplies.*?([0-9,]+\.\d+|[0-9,]+)\s+([0-9,]+\.\d+|[0-9,]+)\s+([0-9,]+\.\d+|[0-9,]+)\s+([0-9,]+\.\d+|[0-9,]+)"),
        "zero_rated_taxable": _first_num(text, r"\(b\)\s+Outward taxable supplies \(zero rated\)\s+([0-9,]+\.\d+|[0-9,]+)"),
        "rcm_taxable": _first_num(text, r"\(d\)\s+Inward supplies \(liable to reverse charge\)\s+([0-9,]+\.\d+|[0-9,]+)"),
        "rcm_igst": _first_num(text, r"\(d\)\s+Inward supplies \(liable to reverse charge\)\s+[0-9,]+\.\d+\s+([0-9,]+\.\d+|[0-9,]+)"),
        "rcm_cgst": _first_num(text, r"\(d\)\s+Inward supplies \(liable to reverse charge\)\s+[0-9,]+\.\d+\s+[0-9,]+\.\d+\s+([0-9,]+\.\d+|[0-9,]+)"),
        "rcm_sgst": _first_num(text, r"\(d\)\s+Inward supplies \(liable to reverse charge\)\s+[0-9,]+\.\d+\s+[0-9,]+\.\d+\s+[0-9,]+\.\d+\s+([0-9,]+\.\d+|[0-9,]+)"),
        "import_igst": _first_num(text, r"\(1\)\s+Import of goods\s+([0-9,]+\.\d+|[0-9,]+)\s+0\.00\s+0\.00\s+0\.00"),
        "all_other_itc_igst": _first_num(text, r"\(5\)\s+All other ITC\s+([0-9,]+\.\d+|[0-9,]+)\s+([0-9,]+\.\d+|[0-9,]+)\s+([0-9,]+\.\d+|[0-9,]+)"),
        "net_itc_igst": _first_num(text, r"\bC\. Net ITC available \(A-B\)\s+([0-9,]+\.\d+|[0-9,]+)"),
        "net_itc_cgst": _first_num(text, r"\bC\. Net ITC available \(A-B\)\s+[0-9,]+\.\d+\s+([0-9,]+\.\d+|[0-9,]+)"),
        "net_itc_sgst": _first_num(text, r"\bC\. Net ITC available \(A-B\)\s+[0-9,]+\.\d+\s+[0-9,]+\.\d+\s+([0-9,]+\.\d+|[0-9,]+)"),
        "restricted_cgst": _first_num(text, r"\(2\)\s+Ineligible ITC under section 16\(4\) & ITC restricted due to PoS rules\s+[0-9,]+\.\d+\s+([0-9,]+\.\d+|[0-9,]+)"),
        "restricted_sgst": _first_num(text, r"\(2\)\s+Ineligible ITC under section 16\(4\) & ITC restricted due to PoS rules\s+[0-9,]+\.\d+\s+[0-9,]+\.\d+\s+([0-9,]+\.\d+|[0-9,]+)"),
        "tax_paid_cash_igst": _first_num(text, r"\(A\) Other than reverse charge.*?Integrated\s+tax.*?Tax paid in cash\s+([0-9,]+\.\d+|[0-9,]+)"),
    }
    out["outward_total"] = round(out["outward_taxable"] + out["zero_rated_taxable"], 2)
    out["notes"] = "Summary PDF parsed; invoice-level liability lines not available."
    return out

def _parse_gstr3b_xlsx(path: Path) -> dict[str, Any]:
    wb = load_workbook(path, data_only=True)
    sheet = wb[wb.sheetnames[0]]
    rows = list(sheet.iter_rows(values_only=True))
    flat = _rows_to_flat_dict(rows)
    flat["source_file"] = path.name
    flat["return_type"] = "GSTR-3B"
    return flat

def parse_gstr2b(path: Path) -> dict[str, Any]:
    wb = load_workbook(path, data_only=True)
    summary = _parse_gstr2b_summary(wb)
    detail = _parse_gstr2b_detail(wb)
    return {
        "source_file": path.name,
        "return_type": "GSTR-2B",
        "summary": summary,
        "detail": detail,
    }

def _parse_gstr2b_summary(wb) -> list[dict[str, Any]]:
    ws = wb["ITC Available"] if "ITC Available" in wb.sheetnames else wb[wb.sheetnames[0]]
    rows = list(ws.iter_rows(values_only=True))
    records: list[dict[str, Any]] = []
    current_block = ""
    seen = set()

    roman = {"I", "II", "III", "IV", "V", "VI", "VII", "VIII", "IX", "X"}

    for row in rows:
        raw = [normalize_space(v) for v in row]
        if not any(raw):
            continue

        col0 = normalize_key(row[0]) if len(row) > 0 else ""
        col1 = normalize_space(row[1]) if len(row) > 1 else ""
        col2 = normalize_space(row[2]) if len(row) > 2 else ""

        if col0 == "FORM SUMMARY - ITC AVAILABLE" or col0 == "S.NO.":
            continue
        if col0 == "PART A":
            current_block = col1 or "Part A"
            continue
        if col0 in SECTION_SKIP:
            continue

        # Only keep compact top-level summary rows: Roman numeral rows and explicit section rows.
        is_top_level = col0 in roman or col0 in {"PART B"}
        if not is_top_level:
            continue

        numeric = [safe_float(row[i] if len(row) > i else 0) for i in range(3, 7)]
        if not any(v != 0 for v in numeric):
            # keep zero rows only if they carry a meaningful heading
            if not col1 and not col2:
                continue

        key = (col0, normalize_key(col1), normalize_key(col2), *numeric)
        if key in seen:
            continue
        seen.add(key)

        records.append({
            "section": current_block,
            "subsection": col0,
            "heading": col1,
            "gstr3b_table": col2,
            "igst": numeric[0],
            "cgst": numeric[1],
            "sgst": numeric[2],
            "cess": numeric[3],
            "advisory": normalize_space(row[7]) if len(row) > 7 else "",
            "source_sheet": "ITC Available",
        })
    return records

def _parse_gstr2b_detail(wb) -> list[dict[str, Any]]:
    records: list[dict[str, Any]] = []
    for sheet_name in wb.sheetnames:
        if sheet_name in {"Read me", "ITC Available", "ITC not available", "ITC Reversal", "ITC Rejected"}:
            continue
        ws = wb[sheet_name]
        rows = list(ws.iter_rows(values_only=True))
        if not rows:
            continue
        headers = _find_header_rows(rows)
        if not headers:
            continue
        header_row_idx, header = headers
        for r in rows[header_row_idx + 1 :]:
            if not any(v is not None and str(v).strip() != "" for v in r):
                continue
            rec = _parse_detail_row(sheet_name, header, r)
            if rec:
                records.append(rec)
    return records

def _find_header_rows(rows):
    for idx, row in enumerate(rows[:10]):
        texts = [normalize_key(v) for v in row if normalize_space(v)]
        joined = " | ".join(texts)
        if ("GSTIN OF SUPPLIER" in joined or "GSTIN OF ISD" in joined or "ICEGATE REFERENCE DATE" in joined or "DOCUMENT NUMBER" in joined):
            # use this row and possibly next row as continuation
            header = [normalize_space(v) for v in row]
            return idx, header
    return None

def _parse_detail_row(sheet_name: str, header: list[Any], row: tuple[Any, ...]) -> dict[str, Any] | None:
    # Provide section-specific parsers
    if sheet_name in {"B2B", "B2BA", "B2B-CDNR", "B2B-CDNRA", "B2B (ITC Reversal)", "B2BA (ITC Reversal)", "B2B-DNR", "B2B-DNRA"}:
        return _parse_b2b_like_row(sheet_name, header, row)
    if sheet_name in {"ISD", "ISDA", "ISD(Rejected)", "ISDA(Rejected)"}:
        return _parse_isd_like_row(sheet_name, header, row)
    if sheet_name in {"IMPG", "IMPGA"}:
        return _parse_impg_like_row(sheet_name, header, row)
    if sheet_name in {"IMPGSEZ", "IMPGSEZA"}:
        return _parse_impgsez_like_row(sheet_name, header, row)
    return None

def _parse_b2b_like_row(sheet_name: str, header: list[Any], row: tuple[Any, ...]) -> dict[str, Any] | None:
    cells = [normalize_space(v) for v in row]
    if not cells or not any(cells):
        return None
    gstin = cells[0] if len(cells) > 0 else ""
    inv_no = cells[2] if len(cells) > 2 else ""
    inv_date = cells[4] if len(cells) > 4 else ""
    taxable = safe_float(row[8] if len(row) > 8 else 0)
    igst = safe_float(row[9] if len(row) > 9 else 0)
    cgst = safe_float(row[10] if len(row) > 10 else 0)
    sgst = safe_float(row[11] if len(row) > 11 else 0)
    cess = safe_float(row[12] if len(row) > 12 else 0)

    # Skip repeated header rows / labels masquerading as data.
    header_markers = {"GSTIN OF SUPPLIER", "INVOICE NUMBER", "INVOICE DETAILS", "TRADE/LEGAL NAME", "INVOICE TYPE"}
    if normalize_key(gstin) in header_markers or normalize_key(inv_no) in header_markers:
        return None
    if not gstin and not inv_no and taxable == 0 and igst == 0 and cgst == 0 and sgst == 0:
        return None

    return {
        "sheet": sheet_name,
        "gstin_supplier": gstin,
        "party_name": cells[1] if len(cells) > 1 else "",
        "invoice_no": inv_no,
        "invoice_no_norm": normalize_invoice_no(inv_no),
        "invoice_type": cells[3] if len(cells) > 3 else "",
        "invoice_date": inv_date,
        "invoice_value": safe_float(row[5] if len(row) > 5 else 0),
        "place_of_supply": cells[6] if len(cells) > 6 else "",
        "rcm": cells[7] if len(cells) > 7 else "",
        "taxable_value": taxable,
        "igst": igst,
        "cgst": cgst,
        "sgst": sgst,
        "cess": cess,
        "period": cells[13] if len(cells) > 13 else "",
        "filing_date": cells[14] if len(cells) > 14 else "",
        "itc_availability": cells[15] if len(cells) > 15 else "",
        "reason": cells[16] if len(cells) > 16 else "",
        "rate": cells[17] if len(cells) > 17 else "",
        "source": cells[18] if len(cells) > 18 else "",
        "irn": cells[19] if len(cells) > 19 else "",
        "irn_date": cells[20] if len(cells) > 20 else "",
    }

def _parse_isd_like_row(sheet_name: str, header: list[Any], row: tuple[Any, ...]) -> dict[str, Any] | None:
    cells = [normalize_space(v) for v in row]
    if not any(cells):
        return None
    return {
        "sheet": sheet_name,
        "isd_gstin": cells[0] if len(cells) > 0 else "",
        "party_name": cells[1] if len(cells) > 1 else "",
        "doc_type": cells[2] if len(cells) > 2 else "",
        "doc_no": cells[3] if len(cells) > 3 else "",
        "doc_date": cells[4] if len(cells) > 4 else "",
        "original_invoice_no": cells[5] if len(cells) > 5 else "",
        "original_invoice_date": cells[6] if len(cells) > 6 else "",
        "igst": safe_float(row[7] if len(row) > 7 else 0),
        "cgst": safe_float(row[8] if len(row) > 8 else 0),
        "sgst": safe_float(row[9] if len(row) > 9 else 0),
        "cess": safe_float(row[10] if len(row) > 10 else 0),
        "period": cells[11] if len(cells) > 11 else "",
        "filing_date": cells[12] if len(cells) > 12 else "",
        "eligibility": cells[13] if len(cells) > 13 else "",
    }

def _parse_impg_like_row(sheet_name: str, header: list[Any], row: tuple[Any, ...]) -> dict[str, Any] | None:
    cells = [normalize_space(v) for v in row]
    if not any(cells):
        return None
    return {
        "sheet": sheet_name,
        "icegate_ref_date": cells[0] if len(cells) > 0 else "",
        "port_code": cells[1] if len(cells) > 1 else "",
        "boe_no": cells[2] if len(cells) > 2 else "",
        "boe_date": cells[3] if len(cells) > 3 else "",
        "taxable_value": safe_float(row[4] if len(row) > 4 else 0),
        "igst": safe_float(row[5] if len(row) > 5 else 0),
        "cess": safe_float(row[6] if len(row) > 6 else 0),
    }

def _parse_impgsez_like_row(sheet_name: str, header: list[Any], row: tuple[Any, ...]) -> dict[str, Any] | None:
    cells = [normalize_space(v) for v in row]
    if not any(cells):
        return None
    return {
        "sheet": sheet_name,
        "supplier_gstin": cells[0] if len(cells) > 0 else "",
        "party_name": cells[1] if len(cells) > 1 else "",
        "icegate_ref_date": cells[2] if len(cells) > 2 else "",
        "port_code": cells[3] if len(cells) > 3 else "",
        "boe_no": cells[4] if len(cells) > 4 else "",
        "boe_date": cells[5] if len(cells) > 5 else "",
        "taxable_value": safe_float(row[6] if len(row) > 6 else 0),
        "igst": safe_float(row[7] if len(row) > 7 else 0),
        "cess": safe_float(row[8] if len(row) > 8 else 0),
    }

def parse_ledger_csv(path: Path) -> dict[str, Any]:
    df = read_csv_safely(path)
    # use a string-based summary that survives ragged rows
    first_rows = df.head(12).fillna("").astype(str).to_dict("records")
    text_rows = [" | ".join(str(v) for v in row.values()) for row in first_rows]
    full_text = "\n".join(text_rows)
    ledger_type = detect_file_kind(path)
    totals = {}
    for col in df.columns:
        ser = pd.to_numeric(df[col], errors="coerce")
        if ser.notna().any():
            totals[str(col)] = float(ser.fillna(0).sum())
    return {
        "source_file": path.name,
        "return_type": ledger_type,
        "shape": [int(df.shape[0]), int(df.shape[1])],
        "columns": [str(c) for c in df.columns],
        "top_rows": first_rows,
        "numeric_totals": totals,
        "notes": _ledger_note_from_name(path.name),
        "raw_preview": full_text[:4000],
    }

def _ledger_note_from_name(name: str) -> str:
    lower = name.lower()
    if "cash" in lower:
        return "Electronic Cash Ledger summary"
    if "credit" in lower:
        return "Electronic Credit Ledger summary"
    if "liability" in lower:
        return "Electronic Liability Register summary"
    return "Ledger summary"

def _rows_to_flat_dict(rows: list[tuple[Any, ...]]) -> dict[str, Any]:
    out = {}
    for row in rows:
        vals = [normalize_space(v) for v in row if normalize_space(v)]
        if len(vals) >= 2 and len(vals) <= 6:
            key = normalize_key(vals[0])
            if key and key not in out:
                out[key] = vals[1] if len(vals) > 1 else ""
    return out

def _first_match(text: str, pattern: str) -> str:
    m = re.search(pattern, text, re.IGNORECASE | re.DOTALL)
    return normalize_space(m.group(1)) if m else ""

def _first_num(text: str, pattern: str) -> float:
    m = re.search(pattern, text, re.IGNORECASE | re.DOTALL)
    if not m:
        return 0.0
    if m.lastindex and m.lastindex >= 1:
        return safe_float(m.group(1))
    return 0.0
