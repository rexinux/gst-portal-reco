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

- Reads files only from the local folders.
- Supports up to 12 months of data.
- Works without API calls or portal automation.
- Parses summary PDFs, Excel files, and ledger CSVs.
- Builds a clean GSTR-1 vs GSTR-3B summary.
- Builds a clean GSTR-2B summary without repeated clutter.
- Adds electronic cash/credit/liability ledger summaries.
- Generates an Excel audit workbook with exceptions and notes.

## Expected inputs

Place month-wise files inside the matching folders.

Supported in this build:
- GSTR-1: PDF summary or Excel
- GSTR-2A: optional PDF/Excel
- GSTR-2B: Excel
- GSTR-3B: PDF summary or Excel
- Electronic ledgers: CSV

## Run

```bash
pip install -r requirements.txt
python -m py_code.main --input-root input --output output/gst_reconciliation_report.xlsx
```

## Output

The workbook includes:
- Dashboard
- GSTR-1 summary
- GSTR-3B summary
- GSTR-2B summary
- GSTR-2B detail
- Electronic ledger summary
- Recon summary
- Exceptions
- Processing notes

## Important limit

This version is strongest for summary reconciliation. Invoice-level matching needs invoice-wise GSTR-1 / purchase exports. Keep the data set within 12 months per run.
