# Verification for 1.1.0

Executed 3 October 2026 using native Windows Excel 16.0. Synthetic data, AI-assisted implementation. This is automated operator evidence, not a human usability trial or a cross-platform certification.

## Executed tests

1. A fresh source rebuild in a disposable directory ran the six-month baseline, added October, checked all monthly totals and exception reasons, repeated refresh, rejected empty/whitespace product-master keys and an incorrect header, then restored and reconciled the source. All 135 assertions passed. See `evidence/native-excel-verification.json`.
2. The sanitizer produced the release workbook with `SET_YOUR_FIXTURES_FOLDER`, generic author metadata, no saved printer blobs and no user-bound permission signatures. Power Query privacy enforcement remains enabled.
3. A copy of that sanitized workbook was moved to a different directory. The first-run test confirmed the placeholder, configured the new local path, executed real Power Query refresh, checked totals/reasons and repeated the negative-input/recovery checks. See `evidence/native-excel-reverification.json` for every assertion.
4. The unchanged sanitized release workbook was reopened read-only in another fresh Excel instance. All 14 checks passed: the placeholder, five embedded-source comparisons and eight saved reconciliation values. See `evidence/saved-workbook-readback.json`.
5. A recursive publication audit inspected the intended file tree, XLSX ZIP parts and decoded Power Query package/metadata. It checked for workstation paths, personal identifiers and credential patterns. It verified the placeholder, removal of machine permission signatures and retained privacy enforcement. Audit results appear in `evidence/publication-audit.json`. The Git tree and release ZIP are also checked before publication.

The Excel-generated monthly and exception previews are included as PNGs. Currency uses the tested machine's localized separators. Every chart starts at zero; labels and the seven-month print layout were visually reviewed. Numeric results are independent of display locale.

## Expected result

| Measure | Result |
|---|---:|
| Input rows | 1,193 |
| Accepted orders | 1,182 |
| Exceptions | 11 |
| Unvalued numeric-error rows | 3 |
| Accepted revenue | CAD 90,028.00 |
| Accepted product cost | CAD 34,867.70 |
| Gross profit | CAD 55,160.30 |
| Row reconciliation difference | 0 |
| Valued-revenue reconciliation difference | 0 cents |

No source files were modified during each successful refresh. Failure tests deliberately change only the disposable copy. Error recovery was verified after restoration. A failed refresh can leave older table values visible, so operators must confirm query completion.

## Limits

Only the supplied seven-month fixture and stated Windows Excel environment are verified. Further chart periods, large inputs, arbitrary monetary magnitudes, simultaneous source edits, Mac, Excel Online and buyer-specific business rules remain untested. Queries are not a database transaction. No client results, guaranteed efficiency gains, live website synchronization or production-scale capability are claimed.

## File-format reference

The sanitizer uses the documented [Microsoft Query Definition File Format](https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-qdeff/22557f6d-7c29-4554-8fe4-7b7a54ac7a2b). Sanitization was followed by native Excel opening and refresh testing; static inspection alone is not the compatibility proof.
