# GST Reconciliation Offline Repo

Offline GST reconciliation workflow for monthly/annual return files.

## Folder layout

```text
input/
  gstr1/
  gstr2a/
  gstr2b/
  gstr3b/
  electronic_credit_ledger/
  electronic_cash_ledger/
  electronic_liability_ledger/
py_code/
  main.py
  config.py
  parsers.py
  reconciler.py
  report.py
  utils.py
```

## What it does

- Reads files only from the local folders - no API calls or portal automation.
- Processes **all** files found in each folder per run (multi-month), grouping by
  tax period detected from the file itself. If two files claim the same period,
  the later one wins and a note is added to Data Quality so nothing is silently lost.
- Parses GSTR-1 and GSTR-3B summary PDFs at table level (via pdfplumber table
  extraction, not brittle text regex), GSTR-2B Excel exports, and the electronic
  cash/credit/liability ledger CSVs (2-row nested header format from the portal).
- De-duplicates ledger transactions across re-exported/overlapping files using a
  content fingerprint (reference no. + date + description + amounts), and reports
  how many duplicates were dropped.
- Builds a bucket-by-bucket ITC reconciliation (GSTR-3B Table 4(A) vs the matching
  GSTR-2B summary section), correctly treating the self-assessed RCM CGST/SGST leg
  as an expected variance rather than an error.
- Builds an outward reconciliation (GSTR-1 vs GSTR-3B).
- Cross-checks the Credit and Cash ledgers against GSTR-3B (ITC accrual, RCM cash payment).
- Produces an Exception Register limited to entries a CA actually needs to look at:
  RCM invoices, POS/eligibility-restricted ITC, credit/debit notes, ISD credits,
  and import-of-goods bills of entry - **not** the full routine B2B feed.
- Flags possible duplicate invoices (same supplier GSTIN + invoice number appearing
  more than once within one GSTR-2B).
- Auto-generates a plain-language "Observations" note per period interpreting the
  anomalies found (RCM variance, ineligible ITC and its stated reason, imports,
  ITC carried forward, etc).
- Reports missing files / incomplete periods explicitly in the workbook so you know
  to add a file and re-run, rather than silently producing a partial report.

## Expected inputs

Place month-wise files inside the matching folders. Multiple months' files can sit
in the same folder - they're grouped by period automatically.

- GSTR-1: PDF summary or Excel
- GSTR-2A: not yet used in reconciliation (folder reserved for future use)
- GSTR-2B: Excel (as exported from the portal)
- GSTR-3B: PDF summary or Excel
- Electronic ledgers: CSV (as exported from the portal - cash ledger only allows
  a 12-month export window, so re-exports with overlapping dates are expected
  and handled)

## Run

```bash
pip install -r requirements.txt
python -m py_code.main --input-root input --output output/gst_reconciliation_report.xlsx
```

## Output

The workbook includes:
- Executive Summary (headline numbers + filing ARNs per period)
- Observations (auto-generated auditor's note)
- Outward Reconciliation (GSTR-1 vs GSTR-3B)
- ITC Reconciliation (GSTR-3B 4(A) buckets vs GSTR-2B, tax-head by tax-head)
- Ledger Summary (period-wise cash/credit/liability activity + GSTR-3B cross-check)
- Exception Register (uncommon GSTR-2B entries only, plus duplicate-invoice flags)
- Data Quality (missing files, parse warnings, duplicate ledger rows dropped)
- Reason Codes (reference table, only for codes actually triggered this run)

## Important limit

Invoice-level GSTR-1 vs GSTR-2A/2B matching still needs an invoice-wise GSTR-1
export (the portal summary PDF only gives table totals). GSTR-3B summary PDFs are
parsed at table level and give exact figures for every line used in this report.
