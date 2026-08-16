# Field Mapping Guide

The WPF app reads mappings from the workbook `Import Mapping` sheet.

The app does not need every import file to use the exact workbook column names. Instead, each row in `Import Mapping` tells the app how to translate an incoming CSV/YAML field into a workbook field.

## Required Mapping Columns

| Column | Purpose |
|---|---|
| Mapping Profile | Name of the import profile, such as `ApplicationOwnerIntake` |
| Source Type | CSV, YAML, or Both |
| Target Table | Workbook table/sheet receiving the update |
| Source Field | Column/property name from the imported file |
| Target Field | Workbook field to update |
| Required? | Yes/No |
| Match Key? | Yes/No; exactly one row should identify the record key |
| Transform Rule | Optional: trim, uppercase, lowercase |
| Default Value | Optional fallback value |
| Notes | Any clarification |

## Example

| Mapping Profile | Source Type | Target Table | Source Field | Target Field | Required? | Match Key? |
|---|---|---|---|---|---|---|
| ApplicationOwnerIntake | CSV | Application Inventory | Application / Component | Application / Component | Yes | Yes |
| ApplicationOwnerIntake | CSV | Application Inventory | Owner | Owner | No | No |
| ApplicationOwnerIntake | CSV | Application Inventory | Phase | Phase | No | No |
| ApplicationOwnerIntake | CSV | Application Inventory | Status | RYG | No | No |
| ApplicationOwnerIntake | CSV | Application Inventory | Known blockers | Notes / Risks | No | No |

## Practical Rule

For tomorrow, keep every team import file simple. Use one row per application/component and these columns:

- Application / Component
- Owner
- Phase
- Status
- Known blockers
- Disposition
- Required for Phase 1?
- Complexity
- Notes

The app will convert each populated field into a `Data Intake` update row and log the import in `Change Log`.
