# CVIS Workbook Updater

This is a Windows WPF helper for safely previewing, staging, and applying updates to the CVIS / Unity modernization recovery workbook.

It is intentionally designed as a controlled intake tool:

- Manual update form for one-off owner, date, RYG, blocker, or dependency updates.
- Focused blocker/risk and ticket/work status forms.
- Bulk CSV/YAML import.
- Field mapping based on the workbook `Import Mapping` sheet.
- Preview with current value, proposed value, target resolution, and conflict result.
- Choice of staging updates in `Data Intake` or safely applying ready updates to source sheets.
- Audit rows in `Change Log` with real old and new values.
- Verified backups, atomic file replacement, and reopen verification.
- Workbook lock detection before saving.
- Compatibility repair for Excel inline validation lists that exceed 255 characters.

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
- .NET 10 SDK (the repository pins SDK 10.0.204)

Open `CVIS.WorkbookUpdater.sln`, restore NuGet packages, then run the app. Verify changes with:

```powershell
dotnet test CVIS.WorkbookUpdater.sln -c Release
```

Packages used:

- ClosedXML and the Open XML SDK for `.xlsx` updates and compatibility normalization
- CsvHelper for CSV import
- YamlDotNet for YAML import

## Mapping

Mappings are read from the workbook `Import Mapping` sheet.

Each mapping profile needs at least one row marked `Match Key? = Yes`. When multiple key rows exist, they are treated as ordered fallbacks. For example, `App ID` can be preferred while `Source Row` is used when an ID is unavailable.

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
- Preview before applying. Conflicting overlapping changes and blank-over-nonblank changes are blocked.
- Every successful save creates a timestamped recovery copy in `.cvis-backups` beside the workbook.

## OneDrive and SharePoint

Use the local path from the OneDrive sync client. This preserves OneDrive upload/versioning behavior while allowing the application to use file locks, backups, and atomic replacement. Wait for OneDrive to report that the file is synchronized before and after a save.

Direct `https://` OneDrive or SharePoint links are intentionally rejected in this version. Writing through a link requires Microsoft Graph authentication, tenant app registration, and approved delegated permissions. The application does not collect or embed those credentials.

## Update Workflows

- **Manual Update**: enter any target sheet, record key, field, and value; stage it or apply it immediately after preview.
- **Bulk Import**: select CSV/YAML and a workbook mapping profile; preview all resolved and blocked changes, then stage or apply.
- **Blocker / Risk**: update `Current RYG`, `Reason`, `Action Plan`, `Leadership Ask`, and `Next Action Date` together. Red and Yellow require an explanation and recovery action.
- **Ticket / Work**: update status and next action on `Action Plan`, `Delivery Tasks`, `Dev Tasks`, `NFRs`, `Platform Onboarding`, or an editable target.

## Next Phase

If true live coauthoring without the OneDrive sync client is required, add a Microsoft Graph adapter or use a SharePoint List as the backing data source. That requires tenant app registration, authentication, and permission approval.
