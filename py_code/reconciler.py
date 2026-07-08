from __future__ import annotations

from typing import Any

from . import rules
from .utils import normalize_key, safe_float

TOLERANCE = 0.50


def _cmp(label: str, v1: float, v2: float, tolerance: float = TOLERANCE) -> dict[str, Any]:
    v1, v2 = safe_float(v1), safe_float(v2)
    diff = round(v1 - v2, 2)
    status = "Matched" if abs(diff) <= tolerance else "Mismatch"
    return {"label": label, "value_1": v1, "value_2": v2, "diff": diff, "status": status}


# ---------------------------------------------------------------------------
# Period indexing (multi-month support with duplicate/conflict sanity check)
# ---------------------------------------------------------------------------

def index_by_period(parsed_list: list[dict[str, Any]]) -> tuple[dict[str, dict[str, Any]], list[str]]:
    """Group parsed return dicts by period_key, keeping the most complete/most
    recent one if two files claim the same period, and reporting the
    conflict so it isn't silently dropped."""
    by_period: dict[str, dict[str, Any]] = {}
    conflicts: list[str] = []
    for item in parsed_list:
        pk = item.get("period_key") or "UNKNOWN"
        if pk in by_period:
            prev = by_period[pk]
            conflicts.append(
                f"Two files both claim period {pk}: '{prev.get('source_file')}' and "
                f"'{item.get('source_file')}'. Kept '{item.get('source_file')}' (last one processed); "
                f"remove the stale file if this wasn't intentional."
            )
        by_period[pk] = item
    return by_period, conflicts


# ---------------------------------------------------------------------------
# Outward reconciliation (GSTR-1 vs GSTR-3B)
# ---------------------------------------------------------------------------

def reconcile_outward(gstr1: dict[str, Any] | None, gstr3b: dict[str, Any] | None) -> dict[str, Any]:
    if not gstr1 or not gstr3b:
        return {"available": False, "rows": [], "note": "GSTR-1 and/or GSTR-3B missing for this period."}

    g1_total = round(
        safe_float(gstr1.get("b2b_taxable")) + safe_float(gstr1.get("sez_taxable"))
        + safe_float(gstr1.get("export_taxable")) + safe_float(gstr1.get("b2cl_taxable"))
        + safe_float(gstr1.get("b2cs_taxable")) + safe_float(gstr1.get("deemed_export_taxable")), 2
    )
    g3_total = round(safe_float(gstr3b.get("outward_taxable")) + safe_float(gstr3b.get("zero_rated_taxable")), 2)

    rows = [
        _cmp("Total outward taxable value (GSTR-1 tables 4-6 vs 3B 3.1(a)+(b))", g1_total, g3_total),
        _cmp("IGST on regular + RCM-attracting outward supply", gstr1.get("b2b_igst"), gstr3b.get("outward_igst")),
        _cmp("CGST on regular outward supply", gstr1.get("b2b_cgst"), gstr3b.get("outward_cgst")),
        _cmp("SGST on regular outward supply", gstr1.get("b2b_sgst"), gstr3b.get("outward_sgst")),
    ]
    for r in rows:
        if r["status"] == "Mismatch":
            explain = rules.explain_pair(
                "OUTWARD", "GSTR-1", "GSTR-3B", r["value_1"], r["value_2"],
                payload={"gstr1": gstr1, "gstr3b": gstr3b},
            )
            r.update({"reason_code": explain["reason_code"], "remark": explain["explanation"]})
        else:
            r.update({"reason_code": "MATCHED", "remark": ""})

    return {"available": True, "rows": rows}


# ---------------------------------------------------------------------------
# ITC reconciliation (GSTR-3B 4A buckets vs GSTR-2B summary sections)
# ---------------------------------------------------------------------------

def _g2b_bucket(summary: list[dict[str, Any]], table_marker: str) -> dict[str, float]:
    for rec in summary:
        if table_marker in (rec.get("gstr3b_table") or ""):
            return {"igst": rec["igst"], "cgst": rec["cgst"], "sgst": rec["sgst"], "cess": rec["cess"]}
    return {"igst": 0.0, "cgst": 0.0, "sgst": 0.0, "cess": 0.0}


def reconcile_itc(gstr3b: dict[str, Any] | None, gstr2b: dict[str, Any] | None) -> dict[str, Any]:
    if not gstr3b or not gstr2b:
        return {"available": False, "rows": [], "note": "GSTR-3B and/or GSTR-2B missing for this period."}

    summary = gstr2b.get("summary", [])
    buckets = [
        ("Import of goods (4(A)(1))", "import_goods_igst", "import_goods_cgst", "import_goods_sgst", "4(A)(1)"),
        ("Inward supplies from ISD (4(A)(4))", "isd_igst", "isd_cgst", "isd_sgst", "4(A)(4)"),
        ("Inward supplies liable to RCM (4(A)(3))", "itc_rcm_igst", "itc_rcm_cgst", "itc_rcm_sgst", "4(A)(3)"),
        ("All other ITC (4(A)(5))", "all_other_itc_igst", "all_other_itc_cgst", "all_other_itc_sgst", "4(A)(5)"),
    ]

    rows = []
    rcm_cgst_sgst_note_needed = False
    for label, ig_key, cg_key, sg_key, marker in buckets:
        g2b_bucket = _g2b_bucket(summary, marker)
        is_rcm = marker == "4(A)(3)"
        for head, g3b_key, g2b_val in [
            ("IGST", ig_key, g2b_bucket["igst"]),
            ("CGST", cg_key, g2b_bucket["cgst"]),
            ("SGST", sg_key, g2b_bucket["sgst"]),
        ]:
            g3b_val = safe_float(gstr3b.get(g3b_key))
            row = _cmp(f"{label} - {head}", g3b_val, g2b_val)
            if is_rcm and head in {"CGST", "SGST"} and row["status"] == "Mismatch":
                row["status"] = "Expected variance"
                row["reason_code"] = "RCM"
                row["remark"] = (
                    "GSTR-2B does not populate the CGST/SGST leg of self-assessed RCM ITC (only the vendor-linked "
                    "IGST portion, if any, is shown). The taxpayer self-assesses and pays this via cash ledger under "
                    "3.1(d)/Table 6.1(B) - this is expected, not a data error."
                )
                rcm_cgst_sgst_note_needed = True
            elif row["status"] == "Mismatch":
                explain = rules.explain_pair(
                    "ITC", "GSTR-3B", "GSTR-2B", g3b_val, g2b_val,
                    payload={"gstr3b": gstr3b},
                )
                row.update({"reason_code": explain["reason_code"], "remark": explain["explanation"]})
            else:
                row.update({"reason_code": "MATCHED", "remark": ""})
            rows.append(row)

    # Net ITC check: GSTR-3B 4C vs sum of the four GSTR-2B buckets above (I-IV),
    # explicitly adding back the RCM CGST/SGST that 2B does not carry.
    g2b_sum_igst = sum(_g2b_bucket(summary, m)["igst"] for _, _, _, _, m in buckets)
    g2b_sum_cgst = sum(_g2b_bucket(summary, m)["cgst"] for _, _, _, _, m in buckets)
    g2b_sum_sgst = sum(_g2b_bucket(summary, m)["sgst"] for _, _, _, _, m in buckets)
    rcm_cgst_g3b = safe_float(gstr3b.get("itc_rcm_cgst"))
    rcm_sgst_g3b = safe_float(gstr3b.get("itc_rcm_sgst"))

    net_row_igst = _cmp("Net ITC available (4C) - IGST", gstr3b.get("net_itc_igst"), g2b_sum_igst)
    net_row_cgst = _cmp("Net ITC available (4C) - CGST (2B buckets + self-assessed RCM)", gstr3b.get("net_itc_cgst"), g2b_sum_cgst + rcm_cgst_g3b)
    net_row_sgst = _cmp("Net ITC available (4C) - SGST (2B buckets + self-assessed RCM)", gstr3b.get("net_itc_sgst"), g2b_sum_sgst + rcm_sgst_g3b)
    for r in (net_row_igst, net_row_cgst, net_row_sgst):
        r["reason_code"] = "MATCHED" if r["status"] == "Matched" else "AMOUNT_MISMATCH"
        r["remark"] = "" if r["status"] == "Matched" else "Residual variance after accounting for self-assessed RCM - review manually."
        rows.append(r)

    # Ineligible / restricted ITC - informational, not a "mismatch" bucket, but
    # worth surfacing since it reduces claimable credit.
    ineligible_row = {
        "label": "Ineligible / PoS-restricted ITC (4(D)(2)) - CGST+SGST",
        "value_1": safe_float(gstr3b.get("restricted_cgst")) + safe_float(gstr3b.get("restricted_sgst")),
        "value_2": None,
        "diff": None,
        "status": "Informational",
        "reason_code": "BLOCKED_ITC",
        "remark": "Reported by taxpayer in 4(D)(2); already excluded from Net ITC in 4C. See exception register for the source invoice(s).",
    }
    rows.append(ineligible_row)

    return {"available": True, "rows": rows, "rcm_note_shown": rcm_cgst_sgst_note_needed}


# ---------------------------------------------------------------------------
# Ledger cross-check
# ---------------------------------------------------------------------------

def reconcile_ledgers(period_key: str, gstr3b: dict[str, Any] | None, merged_credit: dict[str, Any], merged_cash: dict[str, Any]) -> dict[str, Any]:
    credit_tx = [t for t in merged_credit.get("transactions", []) if t.get("period_key") == period_key]
    cash_tx = [t for t in merged_cash.get("transactions", []) if t.get("period_key") == period_key]

    itc_accrued = sum(
        t.get("cgst", 0.0) + t.get("sgst", 0.0) + t.get("igst", 0.0)
        for t in credit_tx if t.get("txn_type") == "Credit" and "ITC ACCRUED" in normalize_key(t.get("description", ""))
    )
    itc_utilised_normal = sum(
        t.get("cgst", 0.0) + t.get("sgst", 0.0) + t.get("igst", 0.0)
        for t in credit_tx if t.get("txn_type") == "Debit" and normalize_key(t.get("description", "")) == "OTHER THAN REVERSE CHARGE"
    )
    rcm_cash_paid = sum(
        t.get("cgst", 0.0) + t.get("sgst", 0.0) + t.get("igst", 0.0)
        for t in cash_tx if t.get("txn_type") == "Debit" and "REVERSE CHARGE" in normalize_key(t.get("description", ""))
    )

    rows = []
    if gstr3b:
        g3b_net_itc = safe_float(gstr3b.get("net_itc_igst")) + safe_float(gstr3b.get("net_itc_cgst")) + safe_float(gstr3b.get("net_itc_sgst"))
        if itc_accrued:
            rows.append(_cmp("ITC accrued in Credit Ledger vs GSTR-3B Net ITC (4C)", itc_accrued, g3b_net_itc))
        g3b_rcm_cash = safe_float(gstr3b.get("rcm_igst")) + safe_float(gstr3b.get("rcm_cgst")) + safe_float(gstr3b.get("rcm_sgst"))
        if rcm_cash_paid:
            rows.append(_cmp("RCM tax paid via Cash Ledger vs GSTR-3B 3.1(d) RCM liability", rcm_cash_paid, g3b_rcm_cash))

    itc_carried_forward = round(itc_accrued - itc_utilised_normal, 2) if itc_accrued else None

    return {
        "credit_ledger_txn_count": len(credit_tx),
        "cash_ledger_txn_count": len(cash_tx),
        "itc_accrued": itc_accrued,
        "itc_utilised_normal": itc_utilised_normal,
        "itc_carried_forward": itc_carried_forward,
        "rcm_cash_paid": rcm_cash_paid,
        "rows": rows,
    }


# ---------------------------------------------------------------------------
# Exception register (uncommon GSTR-2B entries the CA should actually look at)
# ---------------------------------------------------------------------------

def build_exception_register(gstr2b: dict[str, Any] | None) -> list[dict[str, Any]]:
    if not gstr2b:
        return []
    seen_fp = set()
    out = []
    for bucket_name, rows in [
        ("Uncommon document type", gstr2b.get("uncommon", [])),
        ("RCM-liable invoice", gstr2b.get("rcm_rows", [])),
        ("ITC not fully available", gstr2b.get("ineligible_rows", [])),
    ]:
        for r in rows:
            fp = r.get("fingerprint")
            key = (bucket_name, fp)
            if key in seen_fp:
                continue
            seen_fp.add(key)
            out.append({"category": bucket_name, **r})
    return out


def find_duplicate_invoices(gstr2b: dict[str, Any] | None) -> list[dict[str, Any]]:
    """Flag supplier+invoice-number combinations that appear more than once
    within the same GSTR-2B - a genuine data anomaly worth a CA's attention,
    as opposed to routine repeat business with the same vendor."""
    if not gstr2b:
        return []
    from collections import defaultdict
    groups: dict[tuple, list[dict[str, Any]]] = defaultdict(list)
    for r in gstr2b.get("detail", []):
        if r.get("sheet") != "B2B":
            continue
        key = (r.get("gstin_supplier"), r.get("invoice_no_norm"))
        if key[0] and key[1]:
            groups[key].append(r)
    return [
        {"gstin_supplier": k[0], "invoice_no": recs[0].get("invoice_no"), "occurrences": len(recs), "rows": recs}
        for k, recs in groups.items() if len(recs) > 1
    ]


# ---------------------------------------------------------------------------
# Narrative observations (auto-generated, CA-readable)
# ---------------------------------------------------------------------------

def build_observations(
    period_key: str,
    gstr1: dict[str, Any] | None,
    gstr3b: dict[str, Any] | None,
    gstr2b: dict[str, Any] | None,
    itc_reco: dict[str, Any],
    outward_reco: dict[str, Any],
    exception_rows: list[dict[str, Any]],
    duplicate_invoices: list[dict[str, Any]],
    ledger_reco: dict[str, Any],
) -> list[str]:
    notes: list[str] = []

    if itc_reco.get("rcm_note_shown"):
        notes.append(
            "The CGST/SGST leg of reverse-charge ITC will always appear as a variance against GSTR-2B by design - "
            "GSTR-2B only carries the IGST portion reported by the counterparty (if any); the taxpayer self-assesses "
            "and pays the CGST/SGST leg directly. This has been reclassified as an expected variance, not an error."
        )

    ineligible = [r for r in exception_rows if r["category"] == "ITC not fully available"]
    if ineligible:
        total_blocked = sum(safe_float(r.get("cgst")) + safe_float(r.get("sgst")) + safe_float(r.get("igst")) for r in ineligible)
        reasons = {r.get("reason") for r in ineligible if r.get("reason")}
        reason_txt = f" Reason(s) on record: {'; '.join(sorted(reasons))}." if reasons else ""
        notes.append(
            f"₹{total_blocked:,.2f} of ITC across {len(ineligible)} invoice(s) is restricted/ineligible per GSTR-2B "
            f"and has correctly been excluded from Net ITC in GSTR-3B 4(D)(2).{reason_txt}"
        )

    impg = [r for r in exception_rows if r.get("sheet") in {"IMPG", "IMPGA"}]
    if impg:
        total_igst = sum(safe_float(r.get("igst")) for r in impg)
        notes.append(
            f"{len(impg)} bill(s) of entry for import of goods contributed ₹{total_igst:,.2f} IGST credit via customs "
            f"(Table 4(A)(1)) - this is outside the normal vendor-invoice ITC flow and won't appear in the B2B annexure."
        )

    rcm_rows = [r for r in exception_rows if r["category"] == "RCM-liable invoice"]
    if rcm_rows:
        cross_period = [r for r in rcm_rows if r.get("period") and period_key and r.get("period").replace("'", "-20") != period_key]
        if cross_period:
            notes.append(
                f"{len(cross_period)} RCM-liable invoice(s) carry a GSTR-1/IFF filing period different from the "
                f"GSTR-2B period they surfaced in - check they haven't been claimed twice across two months."
            )

    if duplicate_invoices:
        notes.append(
            f"{len(duplicate_invoices)} supplier+invoice-number combination(s) appear more than once in the B2B "
            f"feed for this period - please verify these aren't duplicate vendor uploads before relying on the ITC total."
        )

    if outward_reco.get("available"):
        mismatches = [r for r in outward_reco["rows"] if r["status"] == "Mismatch"]
        if mismatches:
            notes.append(
                f"{len(mismatches)} outward-supply figure(s) differ between GSTR-1 and GSTR-3B beyond rounding "
                f"tolerance - see the Outward Reconciliation sheet for the exact heads affected."
            )
        else:
            notes.append("Outward supply values in GSTR-1 and GSTR-3B agree in full for this period.")

    if itc_reco.get("available"):
        real_mismatches = [r for r in itc_reco["rows"] if r["status"] == "Mismatch"]
        if real_mismatches:
            notes.append(
                f"{len(real_mismatches)} ITC figure(s) remain unexplained after applying the RCM and blocked-ITC "
                f"adjustments above - see the ITC Reconciliation sheet."
            )

    lr = ledger_reco.get("rows", [])
    for r in lr:
        if r["status"] == "Mismatch":
            notes.append(f"Ledger cross-check flagged: {r['label']} differs by ₹{r['diff']:,.2f} - worth a manual look.")

    if ledger_reco.get("itc_carried_forward") not in (None, 0):
        notes.append(
            f"₹{ledger_reco['itc_carried_forward']:,.2f} of ITC accrued this period was not utilised against liability "
            f"and carries forward as a Credit Ledger balance - normal when available ITC exceeds the month's output tax."
        )

    if not notes:
        notes.append("No anomalies detected for this period beyond the routine RCM/ITC-restriction items noted above.")

    return notes
