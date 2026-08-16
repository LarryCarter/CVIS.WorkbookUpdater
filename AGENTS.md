# CVIS Workbook Updater Engineering Contract

## Objective

Recover and modernize the WPF workbook updater on .NET 10 while preserving the supplied workbook's presentation rows, Excel tables, formulas, validations, and auditability.

## Non-negotiable rules

- Never test writes against the user's original workbook. Use a copied or generated fixture.
- Discover worksheets, header rows, and tables from their declared structure; do not assume headers are on row 1.
- Preview and validate changes before applying them.
- Preserve typed dates and numbers instead of converting every value to text.
- Do not overwrite a nonblank workbook value with a blank import value.
- Save through a recoverable backup/atomic-replace workflow before direct source-table updates are enabled.
- Every defect fix requires a regression test.

## Verification

Run `dotnet test CVIS.WorkbookUpdater.sln -c Release` before reporting completion. For workbook changes, reopen the produced copy and verify the expected cells and unchanged presentation rows.
