# CVIS Workbook Updater

This is a Windows WPF helper for the CVIS / Unity modernization recovery workbook.

It is intentionally designed as a controlled intake tool:

- Manual update form for one-off owner, date, RYG, blocker, or dependency updates.
- Bulk CSV/YAML import.
- Field mapping based on the workbook `Import Mapping` sheet.
- Append-only writes to `Data Intake`.
- Audit rows in `Change Log`.
- Workbook lock detection before saving.

## Why It Works This Way

Excel web coauthoring is safe when users edit through Excel or Excel for Web. A desktop WPF app that writes the `.xlsx` file directly is not participating in the Excel web coauthoring session. If the workbook is open or locked, this app stops instead of overwriting someone else's edits.

For tomorrow's workflow:

1. Put the workbook in the team OneDrive or SharePoint location.
2. Ask the team to either update `Data Intake` directly in Excel web or send CSV/YAML files.
3. Use this app to import those CSV/YAML files into the workbook when the local synced file is closed.
4. Review and apply updates from `Data Intake` into source tabs.
5. Dashboard and executive tabs remain the presentation layer.

## Build

Requirements:

- Windows
- Visual Studio 2022
- .NET 8 SDK

Open `CVIS.WorkbookUpdater.sln`, restore NuGet packages, then run the app.

Packages used:

- ClosedXML for `.xlsx` updates
- CsvHelper for CSV import
- YamlDotNet for YAML import

## Mapping

Mappings are read from the workbook `Import Mapping` sheet.

Each mapping profile needs one row marked `Match Key? = Yes`. That source field identifies the record to update, such as `Application / Component`.

Example profile:

| Mapping Profile | Source Field | Target Field | Match Key? |
|---|---|---|---|
| ApplicationOwnerIntake | Application / Component | Application / Component | Yes |
| ApplicationOwnerIntake | Owner | Owner | No |
| ApplicationOwnerIntake | Phase | Phase | No |
| ApplicationOwnerIntake | Status | RYG | No |

## Safe Use Rules

- Do not run the app while the workbook is open in desktop Excel.
- If the workbook is being actively edited in Excel web, wait for OneDrive sync before importing.
- Treat `Data Intake` as the staging queue.
- Treat `Change Log` as the audit trail.
- Do not let the app write directly to dashboard cells.

## Next Phase

If true live coauthoring from WPF is required, the next version should use Microsoft Graph / Excel workbook APIs or a SharePoint List as the backing data source. That requires tenant app registration, authentication, and permission approval.
