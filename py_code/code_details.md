# py_code module map

- `main.py`  
  Entry point. Scans the input folders, loads the latest file in each return folder, runs reconciliation, and writes the workbook.

- `parsers.py`  
  Offline parsers for GSTR-1, GSTR-2B, GSTR-3B, and the electronic ledgers.

- `reconciler.py`  
  Builds the summary comparison rows and dashboard flags.

- `report.py`  
  Writes the Excel workbook with clean sheets and basic formatting.

- `utils.py`  
  Small helpers for normalization, CSV reading, invoice cleaning, and number conversion.

- `config.py`  
  Central settings for folder names, supported file types, and tolerance.

## Input folders expected

- `input/gstr1/`
- `input/gstr2a/` optional
- `input/gstr2b/`
- `input/gstr3b/`
- `input/electronic_credit_ledger/`
- `input/electronic_cash_ledger/`
- `input/electronic_liability_ledger/`

## Notes

- No API usage.
- No portal automation.
- Summary PDFs are supported first; invoice-level exports can be added later.
