from __future__ import annotations

from dataclasses import dataclass
from typing import Any

def compare_summary(gstr1: dict[str, Any] | None, gstr3b: dict[str, Any] | None, gstr2b: dict[str, Any] | None, ledgers: list[dict[str, Any]]) -> dict[str, Any]:
    gstr1 = gstr1 or {}
    gstr3b = gstr3b or {}
    gstr2b = gstr2b or {}

    g1_outward = float(gstr1.get("outward_total", gstr1.get("total_liability", 0.0)) or 0.0)
    g3_outward = float(gstr3b.get("outward_total", 0.0) or 0.0)

    g2b_igst = sum(float(r.get("igst", 0.0) or 0.0) for r in gstr2b.get("detail", []))
    g2b_cgst = sum(float(r.get("cgst", 0.0) or 0.0) for r in gstr2b.get("detail", []))
    g2b_sgst = sum(float(r.get("sgst", 0.0) or 0.0) for r in gstr2b.get("detail", []))
    g2b_cess = sum(float(r.get("cess", 0.0) or 0.0) for r in gstr2b.get("detail", []))

    g3_igst = float(gstr3b.get("net_itc_igst", 0.0) or 0.0)
    g3_cgst = float(gstr3b.get("net_itc_cgst", 0.0) or 0.0)
    g3_sgst = float(gstr3b.get("net_itc_sgst", 0.0) or 0.0)

    cash_ledger = next((x for x in ledgers if x.get("return_type") == "electronic_cash_ledger"), {})
    credit_ledger = next((x for x in ledgers if x.get("return_type") == "electronic_credit_ledger"), {})
    liability_ledger = next((x for x in ledgers if x.get("return_type") == "electronic_liability_ledger"), {})

    recon_rows = [
        {
            "section": "Outward turnover",
            "source_1": "GSTR-1",
            "source_2": "GSTR-3B",
            "value_1": round(g1_outward, 2),
            "value_2": round(g3_outward, 2),
            "difference": round(g1_outward - g3_outward, 2),
            "status": "Matched" if abs(g1_outward - g3_outward) <= 0.5 else "Mismatch",
        },
        {
            "section": "ITC - IGST",
            "source_1": "GSTR-2B",
            "source_2": "GSTR-3B",
            "value_1": round(g2b_igst, 2),
            "value_2": round(g3_igst, 2),
            "difference": round(g2b_igst - g3_igst, 2),
            "status": "Matched" if abs(g2b_igst - g3_igst) <= 0.5 else "Mismatch",
        },
        {
            "section": "ITC - CGST",
            "source_1": "GSTR-2B",
            "source_2": "GSTR-3B",
            "value_1": round(g2b_cgst, 2),
            "value_2": round(g3_cgst, 2),
            "difference": round(g2b_cgst - g3_cgst, 2),
            "status": "Matched" if abs(g2b_cgst - g3_cgst) <= 0.5 else "Mismatch",
        },
        {
            "section": "ITC - SGST",
            "source_1": "GSTR-2B",
            "source_2": "GSTR-3B",
            "value_1": round(g2b_sgst, 2),
            "value_2": round(g3_sgst, 2),
            "difference": round(g2b_sgst - g3_sgst, 2),
            "status": "Matched" if abs(g2b_sgst - g3_sgst) <= 0.5 else "Mismatch",
        },
        {
            "section": "ITC - Cess",
            "source_1": "GSTR-2B",
            "source_2": "GSTR-3B",
            "value_1": round(g2b_cess, 2),
            "value_2": 0.0,
            "difference": round(g2b_cess, 2),
            "status": "Check",
        },
    ]

    warnings = []
    if not gstr2b.get("detail"):
        warnings.append("GSTR-2B detail sheet was empty or not parsed.")
    if not gstr1:
        warnings.append("GSTR-1 summary not parsed.")
    if not gstr3b:
        warnings.append("GSTR-3B summary not parsed.")
    if cash_ledger:
        warnings.append("Electronic Cash Ledger included in this run.")
    if credit_ledger:
        warnings.append("Electronic Credit Ledger included in this run.")
    if liability_ledger:
        warnings.append("Electronic Liability Register included in this run.")

    return {
        "recon_rows": recon_rows,
        "warnings": warnings,
        "dashboard": {
            "outward_match": abs(g1_outward - g3_outward) <= 0.5,
            "itc_match_igst": abs(g2b_igst - g3_igst) <= 0.5,
            "itc_match_cgst": abs(g2b_cgst - g3_cgst) <= 0.5,
            "itc_match_sgst": abs(g2b_sgst - g3_sgst) <= 0.5,
        },
    }
