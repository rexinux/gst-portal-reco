# Changelog / Known-issue track record

This file exists so nothing gets silently "fixed" without a record of what was
wrong and how it was verified. Every entry below was confirmed against the
actual April'25 sample data before being marked fixed.

## 2026-07-09 - Full revamp

1. **Ledger CSV parser produced garbage output.**
   - What: `electronic_cash_ledger` / `electronic_credit_ledger` /
     `electronic_liability_ledger` CSVs use a 2-row nested header (tax head
     group row + Tax/Interest/Penalty/Fee/Others/Total sub-header row). The
     old parser fed this straight into `pandas.read_csv`, which produced a
     4-row x 1-column frame - unusable.
   - Fix: rewrote as a dedicated fixed-layout parser per ledger type that
     locates the header row, skips both header rows, and reads each
     tax-head's Tax sub-column explicitly.
   - Verified: manually cross-checked parsed IGST/CGST/SGST figures against
     the raw CSV cell-by-cell for April'25.

2. **GSTR-3B PDF parser silently misread a real figure as zero.**
   - What: the old code used regex on `page.extract_text()`. pdfplumber
     sometimes renders `0.00` as `0.0 0` (stray space) due to column
     kerning. The old regex didn't tolerate this, so restricted/ineligible
     ITC (4(D)(2)) was read as ₹0/₹0 when the PDF actually said ₹60/₹60.
   - Fix: switched to `page.extract_tables()` (structured table extraction)
     for both GSTR-1 and GSTR-3B, keyed on row labels instead of raw-text
     regex. Removes the whitespace-fragility entirely.
   - Verified: every GSTR-3B field now matches the source PDF exactly,
     confirmed line by line against a manual `pdfplumber` dump.

3. **`main.py` only read the single latest file per folder.**
   - What: multi-month data (the actual point of the "up to 12 months"
     design) was silently discarded - only `sorted(files)[-1]` was parsed.
   - Fix: every file in a folder is now parsed and grouped by detected
     return period. If two files claim the same period, the conflict is
     recorded in the Data Quality sheet instead of one silently overwriting
     the other.

4. **`rules.py` was written but never called from anywhere.**
   - Fix: wired into `reconciler.py` for outward and residual ITC mismatches.

5. **GSTR-2B detail parser leaked one fake "header echo" row per sheet
   for ISD/IMPG/IMPGSEZ (and would have for CDNR/DNR too, if any existed).**
   - What: the sheet has *two* header rows (group header + units row), but
     the loop only skipped one. For B2B this was masked by an explicit
     header-text filter; the other sheet types had no such filter, so an
     empty/garbage row was emitted for every one of them even when the
     sheet had zero real transactions.
   - Fix: loop now starts two rows after the detected header row, for all
     sheet types.
   - Verified: uncommon-sheet counts dropped from "1 fake row per empty
     sheet" to the correct 0, while the 2 genuine IMPG rows remained.

6. **GSTR-1 4A (B2B Regular) total was read as all-zero.**
   - What: an `exclude="reverse charge"` guard was added to stop the 4A
     lookup from matching the 4B (reverse-charge) section - but the 4A
     header text itself contains the phrase "other than reverse charge
     supplies", so the guard rejected the correct row.
   - Fix: removed the exclude guard; section headers are distinct enough
     (`"4A -"` vs `"4B -"`) without it.

7. **Ledger cross-check compared the wrong pair of numbers.**
   - What: the first version compared Credit Ledger *utilisation* (Debit
     entries, i.e. ITC actually used to pay liability) against GSTR-3B Net
     ITC (which is the *accrual* figure). These are conceptually different
     any month a closing ITC balance is carried forward, so it always
     "mismatched" even when nothing was wrong.
   - Fix: now compares Credit entries ("ITC accrued through - Inputs")
     against Net ITC (apples to apples), and reports the
     accrual-minus-utilisation gap separately as "ITC carried forward"
     (informational, not a mismatch).

8. **Outward reconciliation had a meaningless comparison.**
   - What: GSTR-1 Table 4B (outward supplies *made by* the taxpayer that
     attract reverse charge) was being compared against GSTR-3B 3.1(d)
     (inward supplies *received* by the taxpayer under RCM). These are
     unrelated flows - GSTR-1 has no field for inward RCM at all.
   - Fix: removed the comparison row entirely.

9. **Credit ledger balance columns were off by one, portal-wide.**
   - What: the credit-ledger layout has a trailing "Total" column after
     both the Credit/Debit group and the Balance group (5 columns per
     group: IGST, CGST, SGST, CESS, Total) - unlike the cash/liability
     ledgers, whose "Total" is already inside each head's 6-wide
     Tax/Interest/Penalty/Fee/Others/Total sub-block. The parser didn't
     skip this extra column, so every *balance* field was reading one
     column early (`igst_balance` actually held the group's Total, etc).
     Tax/debit/credit amounts themselves were unaffected.
   - Fix: added a `group_has_total_col` layout flag that skips the extra
     column after both groups.
   - Verified: recomputed April's closing SGST balance by hand from the
     raw CSV (opening 1,789,790 -> March txns -> April credit) and it now
     matches the parsed `sgst_balance` exactly (2,032,607).
   - Note: this bug never reached the delivered workbook - no sheet
     displayed ledger balances, only period debit/credit totals (which use
     the transaction amount columns, not balance columns). Fixed anyway
     since it would have produced wrong numbers the moment balances were
     used for anything (e.g. the interpretive review requested afterwards).

## 2026-07-11 - Formula-driven Reconciliation Grid + Table 6.1 mining

10. **New: "Processed Data" engine sheet.** One visible row per FY month
    (Apr-Mar, fixed order, 12 rows always present even for months with no
    file uploaded yet), holding every figure the report sheets need. This
    is the single editable source of truth - change a number here and every
    sheet that references it recalculates, without re-running the tool.

11. **New: "Reconciliation Grid" sheet**, modelled directly on the user's
    own reference template (`GST_Comparison_FY_2025-26.xlsx`): GSTR-3B
    Sales/ITC/RCM/Cash-paid blocks, GSTR-1 Sales, and a Tally
    (books) comparison block with manual-entry cells (pale yellow) and
    live diff formulas. Every non-Tally cell is a formula pointing at
    Processed Data - nothing here is a static Python-written number.
    Added one thing beyond the reference template: a live GSTR-1 vs
    GSTR-3B sales diff block with conditional formatting (green/red), so
    the same outward-turnover figure doesn't have to be read twice from
    two different blocks and compared by eye.

12. **New: "Payment Detail" sheet**, mining GSTR-3B Table 6.1 (Payment of
    Tax) in full - for each tax head, how much of the liability was
    cleared via same-head ITC, cross-utilised ITC from another head, or
    cash, plus interest and late fee *as recorded on the return itself*.
    This directly settles a question the previous version could only
    guess at from the cash ledger (see the July 9 interpretive review,
    observation #4) - Table 6.1 shows the return-level interest/late-fee
    figures directly (₹0 for April'25, confirmed, not inferred).

13. **Full B2B invoice list and all three ledger transaction logs are now
    hidden sheets** (kept for audit trail, not deleted), rather than
    omitted or shown inline - cuts the visual clutter the user flagged
    without losing traceability. The Exception Register (uncommon entries
    only) stays visible since it was never the noisy part.

14. **Removed** the old static "Outward Reconciliation" and "Ledger
    Summary" sheets - both are now superseded by the live Reconciliation
    Grid and Payment Detail sheets. Kept "ITC Reconciliation" (GSTR-3B vs
    GSTR-2B bucket-level, with the RCM-expected-variance logic) since it
    covers something the new grid doesn't: a live tie-out against GSTR-2B
    specifically, not just GSTR-3B's own figures.

