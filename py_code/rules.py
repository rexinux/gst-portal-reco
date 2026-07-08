from __future__ import annotations

from dataclasses import dataclass, asdict
from typing import Any

DEFAULT_TOLERANCE = 0.50

@dataclass(frozen=True)
class RuleExplanation:
    code: str
    title: str
    applies_to: str
    explanation: str
    action: str
    confidence: str = "High"
    severity: str = "Info"

def mismatch_catalog() -> list[RuleExplanation]:
    return [
        RuleExplanation(
            code="MATCHED",
            title="Matched",
            applies_to="Any",
            explanation="The two values are within tolerance and no exception is required.",
            action="No action.",
            severity="Info",
        ),
        RuleExplanation(
            code="SUMMARY_SCOPE_MISMATCH",
            title="Summary-scope mismatch",
            applies_to="Summary comparisons",
            explanation="The two sources are not measuring the same population. This often happens when a summary return is compared with a line-item bucket or when GSTR-2B detail is mixed with a GSTR-3B net figure.",
            action="Split the comparison into equivalent buckets before treating it as a GST variance.",
            severity="High",
        ),
        RuleExplanation(
            code="CATEGORY_SCOPE_MISMATCH",
            title="Category scope mismatch",
            applies_to="ITC comparisons",
            explanation="The ITC categories on both sides are different. One source may include only B2B rows while the other includes import, RCM, ISD, or restricted buckets.",
            action="Compare like-for-like ITC buckets only.",
            severity="High",
        ),
        RuleExplanation(
            code="PERIOD_SHIFT",
            title="Period shift",
            applies_to="All return comparisons",
            explanation="The same invoice or liability likely belongs to a different tax period.",
            action="Check adjacent months, amendment tables, and cut-off dates.",
            severity="Medium",
        ),
        RuleExplanation(
            code="AMENDMENT_FOUND",
            title="Amendment detected",
            applies_to="GSTR-1 / GSTR-2A / GSTR-2B",
            explanation="The original line was later amended, and the amended value should be checked against the reporting period where the amendment actually appears.",
            action="Link original and amended documents before concluding a mismatch.",
            severity="High",
        ),
        RuleExplanation(
            code="CREDIT_NOTE",
            title="Credit note adjustment",
            applies_to="Outward / ITC",
            explanation="A credit note reduces the original value and can make a direct comparison look mismatched if the credit note is ignored.",
            action="Apply credit note impact to the base invoice before evaluating the difference.",
            severity="High",
        ),
        RuleExplanation(
            code="DEBIT_NOTE",
            title="Debit note adjustment",
            applies_to="Outward / ITC",
            explanation="A debit note increases the original value and can make a direct comparison look mismatched if the debit note is ignored.",
            action="Apply debit note impact to the base invoice before evaluating the difference.",
            severity="High",
        ),
        RuleExplanation(
            code="RCM",
            title="Reverse charge component",
            applies_to="ITC / liability",
            explanation="Reverse charge tax is reported through a separate GST flow and should be compared in its own bucket.",
            action="Keep RCM separate from normal purchase ITC.",
            severity="High",
        ),
        RuleExplanation(
            code="BLOCKED_ITC",
            title="Blocked or restricted ITC",
            applies_to="ITC",
            explanation="Some ITC is disallowed or restricted under GST rules and must not be mixed with normal claimable ITC.",
            action="Separate blocked and restricted ITC from eligible ITC.",
            severity="High",
        ),
        RuleExplanation(
            code="POS_RESTRICTION",
            title="Place of supply restriction",
            applies_to="ITC / tax head",
            explanation="The place of supply or state logic can make IGST / CGST / SGST splits look wrong even when the base value is right.",
            action="Verify GSTIN state code and the POS rule before concluding an error.",
            severity="High",
        ),
        RuleExplanation(
            code="GSTIN_MISMATCH",
            title="GSTIN mismatch",
            applies_to="All invoice-level comparisons",
            explanation="The supplier or recipient GSTIN is different across sources or has been captured incorrectly.",
            action="Normalize GSTIN and verify the source document.",
            severity="High",
        ),
        RuleExplanation(
            code="INVOICE_NO_MISMATCH",
            title="Invoice number mismatch",
            applies_to="Invoice-level comparisons",
            explanation="The same transaction may exist but the invoice number was typed differently, OCR-ed incorrectly, or amended later.",
            action="Normalize the invoice number and look for series changes or OCR errors.",
            severity="Medium",
        ),
        RuleExplanation(
            code="DATE_MISMATCH",
            title="Document date mismatch",
            applies_to="Invoice-level comparisons",
            explanation="The same invoice may be captured in a different day or month, or the source may use accounting date instead of invoice date.",
            action="Check the source document date and posting date separately.",
            severity="Medium",
        ),
        RuleExplanation(
            code="AMOUNT_MISMATCH",
            title="Amount mismatch",
            applies_to="All comparisons",
            explanation="The transaction matches, but taxable value or tax amount differs.",
            action="Split the mismatch into taxable value, IGST, CGST, SGST, cess, and total value components.",
            severity="High",
        ),
        RuleExplanation(
            code="ROUNDING_ONLY",
            title="Rounding difference",
            applies_to="All comparisons",
            explanation="The mismatch is within tolerance and is likely only a paise-level rounding effect.",
            action="Mark as accepted rounding variance.",
            severity="Info",
        ),
        RuleExplanation(
            code="DUPLICATE",
            title="Duplicate record",
            applies_to="All sources",
            explanation="The same record appears more than once in the source or across sources.",
            action="Deduplicate on GSTIN + invoice number + date + amount + document type.",
            severity="High",
        ),
        RuleExplanation(
            code="DATA_QUALITY",
            title="Data quality issue",
            applies_to="Parsing",
            explanation="The file layout, OCR text, or exported structure is incomplete or inconsistent.",
            action="Flag the source file for manual review and keep the raw evidence.",
            severity="High",
        ),
        RuleExplanation(
            code="MANUAL_REVIEW",
            title="Manual review",
            applies_to="Fallback",
            explanation="The engine cannot confidently classify the mismatch with the available offline evidence.",
            action="Review the document manually and then refine the rule set.",
            severity="Medium",
        ),
    ]

def as_rows() -> list[dict[str, Any]]:
    return [asdict(x) for x in mismatch_catalog()]

def explain_pair(
    section: str,
    source_1: str,
    source_2: str,
    value_1: float,
    value_2: float,
    payload: dict[str, Any] | None = None,
    tolerance: float = DEFAULT_TOLERANCE,
) -> dict[str, Any]:
    payload = payload or {}
    diff = round(float(value_1) - float(value_2), 2)
    abs_diff = abs(diff)

    if abs_diff <= tolerance:
        return {
            "reason_code": "MATCHED",
            "reason_text": "Values are within tolerance.",
            "explanation": "No GST variance after applying tolerance.",
            "confidence": "High",
            "action": "No action required.",
            "severity": "Info",
            "evidence": f"{source_1}={value_1:.2f}; {source_2}={value_2:.2f}",
        }

    sec = section.upper()

    if "OUTWARD" in sec:
        if "SUMMARY" in (payload.get("gstr1", {}).get("notes", "") + payload.get("gstr3b", {}).get("notes", "")).upper():
            return {
                "reason_code": "SUMMARY_SCOPE_MISMATCH",
                "reason_text": "Summary-level comparison only.",
                "explanation": (
                    "This comparison is driven by summary PDFs, so the engine can verify return-level totals "
                    "but cannot isolate invoice-wise reasons until invoice exports are provided."
                ),
                "confidence": "High",
                "action": "Use invoice-wise GSTR-1 export for detailed matching.",
                "severity": "Medium",
                "evidence": f"{source_1}={value_1:.2f}; {source_2}={value_2:.2f}",
            }

    if sec.startswith("ITC") or "ITC" in sec:
        g3 = payload.get("gstr3b", {}) or {}
        comps = []
        for k in ["import_igst", "rcm_igst", "rcm_cgst", "rcm_sgst", "all_other_itc_igst", "all_other_itc_cgst", "all_other_itc_sgst", "restricted_cgst", "restricted_sgst"]:
            v = g3.get(k)
            if isinstance(v, (int, float)) and float(v) != 0.0:
                comps.append(f"{k}={float(v):.2f}")
        if comps:
            return {
                "reason_code": "CATEGORY_SCOPE_MISMATCH",
                "reason_text": "ITC buckets are not like-for-like.",
                "explanation": (
                    "The compared figures are not the same GST bucket. GSTR-3B net ITC can contain import, RCM, "
                    "other ITC and restricted components, while the GSTR-2B extraction may only include one subset. "
                    "Compare equivalent categories only."
                ),
                "confidence": "High",
                "action": "Split ITC into matching buckets before variance analysis.",
                "severity": "High",
                "evidence": "; ".join(comps),
            }

    # Generic fallback
    if abs_diff > tolerance:
        return {
            "reason_code": "AMOUNT_MISMATCH",
            "reason_text": "Amount differs.",
            "explanation": "A measurable variance remains after the current classification rules.",
            "confidence": "Medium",
            "action": "Review the source rows, then add a more specific rule if this pattern repeats.",
            "severity": "High",
            "evidence": f"{source_1}={value_1:.2f}; {source_2}={value_2:.2f}; diff={diff:.2f}",
        }

    return {
        "reason_code": "ROUNDING_ONLY",
        "reason_text": "Rounding only.",
        "explanation": "The variance is within tolerance and is treated as rounding.",
        "confidence": "High",
        "action": "Accept the variance.",
        "severity": "Info",
        "evidence": f"{source_1}={value_1:.2f}; {source_2}={value_2:.2f}",
    }
