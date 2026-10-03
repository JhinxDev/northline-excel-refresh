# Use and reproduce the workbook

## First run

1. Download and extract the complete release ZIP. Keep the workbook and `fixtures` folder together. Cached synthetic results are provided for browsing.
2. Open the workbook in Windows Excel with Power Query. Go to Data > Queries & Connections, right-click `DataRoot`, choose Edit, then Advanced Editor.
3. Replace the single string `"SET_YOUR_FIXTURES_FOLDER"` with the absolute path to your extracted `fixtures` folder, enclosed in double quotes. Forward slashes can be used in a Windows path. Close & Load. Use Excel's normal local-file privacy prompts if needed; do not disable privacy controls globally.
4. Choose Data > Refresh All and wait for all queries to finish. A failure means the results may be stale, even if numbers remain visible.
5. Review Reconciliation: both differences must be zero. Then read Exceptions. Zero differences do not mean every source row was accepted.

## Updating and correcting data

Add monthly CSVs immediately inside `fixtures/monthly`; only lowercase `.csv` files in that directory are read. Keep backups elsewhere to avoid duplicate ingestion. Use the exact nine-column header and UTF-8 CSV layout shown in the fixtures. IDs/SKUs are case-sensitive. Prices/costs are integer CAD cents, quantities positive integers. The source has one product per order.

Exceptions preserve the input values, filename and CSV row number. Fix the original source, then refresh. Never repair a loaded Excel result table because refresh replaces it. Every copy of a repeated order ID is excluded. Unknown SKUs must be reconciled with an authoritative product master. Blank/whitespace master keys and source-header changes fail refresh.

Numeric errors are counted as unvalued rows. Raw-valued revenue equals accepted plus excluded-valued revenue; it does not recover invalid amounts. Keep input files stable during refresh: separate queries do not form a transactional snapshot.

The chart is configured for seven fixture months. Extend its source and print range when adding further periods. Human practice and buyer-specific acceptance remain separate from the automated tests.

## Run the included tests

Run tests in a disposable extracted copy, since they change that copy's DataRoot, inputs and output. Close its workbook first. The scripts create and close only their own hidden Excel instance, with macros disabled.

To test the shipped sanitized workbook and first-run setup:

```powershell
powershell -NoProfile -File source/build-and-verify.ps1 -VerifyExisting
```

This expects the release placeholder, configures the copy's local fixture directory, tests final totals and exception reasons, repeats refresh, tests bad master/schema inputs and restores them. It writes `evidence/native-excel-reverification.json`. To repeat this first-run test, use another fresh extracted copy. The saved test copy contains your local path; do not redistribute it as the clean release.

To rebuild from editable source and execute the six-month-to-seven-month exercise:

1. In another disposable copy, move `fixtures/monthly/2026-10.csv` outside `monthly`.
2. Install the pinned build dependency into your chosen Python environment: `python -m pip install -r requirements-build.txt`.
3. Run `python source/prepare.py`. This uses the bundled original synthetic sales CSV and replaces the workbook shell.
4. Run `powershell -NoProfile -File source/build-and-verify.ps1`. It checks the original six months, adds October, verifies the final result, repeats refresh, tests source failures and recovery, and writes `evidence/native-excel-verification.json`.
5. Inspect the receipt and nonzero exit status on failure. Full rebuild acceptance has 135 checks. No success is implied by an old receipt left from a previous run.

To prepare a distributable workbook from your test copy, run `python source/sanitize_workbook.py tested.xlsx release.xlsx` with distinct input/output paths. The sanitizer replaces DataRoot with the placeholder, clears author/printer and machine-bound permission metadata, and preserves the privacy setting. Always test the sanitized result on a separate copy before distributing. It is not a generic sanitizer for arbitrary client data or unrelated workbook formats.

The included `source/readback.ps1` reopens a release workbook read-only and verifies its placeholder, embedded source and saved values. It does not refresh or prove another computer's privacy settings. Only the declared Windows Excel environment was exercised.
