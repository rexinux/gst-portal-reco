# py_code module map

- `main.py`
  Entry point. Scans every input folder, parses **all** files in each
  (not just the latest), groups them by detected tax period, merges and
  de-duplicates the ledger exports, runs reconciliation per period, and
  writes the workbook.

- `parsers.py`
  Parsers for GSTR-1 and GSTR-3B (PDF via pdfplumber table extraction, or
  Excel), GSTR-2B (Excel: summary + per-sheet detail, with an uncommon/RCM/
  ineligible split), and the three electronic ledgers (CSV, fixed
  portal layout, multi-file merge with fingerprint-based de-dup).

- `reconciler.py`
  Period indexing/conflict detection, bucket-level ITC reco (GSTR-3B 4(A) vs
  GSTR-2B summary sections, tax-head by tax-head, with RCM CGST/SGST treated
  as an expected variance), ledger cross-check, exception register,
  duplicate-invoice detection, auto-generated narrative observations, and
  `build_processed_data()` - assembles the one-row-per-FY-month table that
  the formula-driven report sheets read from.

- `report.py`
  Writes the Excel workbook. Two kinds of sheets:
  - *Formula-driven*: Processed Data (engine sheet, Python writes values),
    Reconciliation Grid, Payment Detail - both of the latter are almost
    entirely `="Processed Data"!...` formulas, not Python-computed numbers,
    so editing Processed Data in Excel updates them live.
  - *Static*: Executive Summary, Observations, ITC Reconciliation, Exception
    Register, Data Quality, Reason Codes - values written directly by
    Python, refreshed by re-running the tool.
  Also writes 4 hidden sheets (full B2B list + all 3 ledgers) for audit
  trail without cluttering the visible workbook.

- `rules.py`
  Reason-code catalog and explanation lookup, used by `reconciler.py` for
  any mismatch that isn't already deterministically bucketed.

- `utils.py`
  Normalization, period-key parsing (from dates or "Apr-25"-style labels),
  FY month-skeleton generation (`fy_months`, `infer_fy_start_year` - the
  fixed Apr-Mar row order the formula-driven sheets rely on), duplicate-
  detection fingerprinting, and number parsing.

- `config.py`
  Central settings for folder names, supported file types, and tolerance.

See `CHANGELOG.md` for a record of bugs found and fixed during the July 2026
revamp - useful if a number ever looks wrong again and you want to check
whether it's a known, already-fixed class of issue.

## Input folders expected

- `input/gstr1/`
- `input/gstr2a/` - reserved, not yet used in reconciliation
- `input/gstr2b/`
- `input/gstr3b/`
- `input/electronic_credit_ledger/`
- `input/electronic_cash_ledger/`
- `input/electronic_liability_ledger/`

Any number of files per folder is fine - they're grouped by period
automatically, and duplicate/overlapping ledger exports are de-duplicated.


## Notes

- No API usage.
- No portal automation.
- Summary PDFs are supported first; invoice-level exports can be added later.
