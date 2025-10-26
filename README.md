# Admin Office Licenses Web App

This Apps Script project exposes a web dashboard for tracking license records stored in a Google Sheet.

## Spreadsheet configuration

The server code expects to read and write license data from a spreadsheet that contains a sheet named `Licenses` with the header defined in `code.gs`.

1. Create (or identify) the spreadsheet that should back the dashboard.
2. Open the Apps Script project and set the `SPREADSHEET_ID` constant in `code.gs` to that spreadsheet's ID.
   - The ID is the long string between `/d/` and `/edit` in the spreadsheet URL.
   - Leave the constant blank only when the script is container-bound to the desired spreadsheet.
3. Save the script changes before testing or deploying the web app.

If the ID is missing and the script is not bound to a spreadsheet, calls to `getDashboardData` will throw an error instructing you to set `SPREADSHEET_ID`.

## Deployment notes

When deploying as a web app, ensure the published script has access to the configured spreadsheet and, if necessary, update the Drive folder permissions for uploaded files.

## License renewals & history tracking

The backend exposes dedicated helpers for updating existing rows and for fetching their revision history:

* `updateLicense(payload)` replaces the editable cells (expiry labels/dates, status, and file metadata) for a given `id`. When a replacement file is supplied it is uploaded to Drive before the sheet is updated. The function recomputes status columns with `computeStatus_` and stores a snapshot of the pre-update values in `LicenseHistory`.
* `renewLicense(payload)` is a thin alias around `updateLicense` that always records the action as a renewal. Use this when the UI offers a “renew” workflow.
* `getLicenseHistory(id)` reads the `LicenseHistory` sheet and returns the previous values ordered from newest to oldest. The payload includes timestamps, the previous expiry labels/dates, status labels, and file URLs so the frontend can render a detailed timeline.

The history sheet is initialised automatically with the `LICENSE_HISTORY_HEADER` column order. Every update (or renewal) appends a new row capturing the prior expiry dates, status labels, Drive file link, and a timestamp of when the change occurred. Dashboard rows also include a `hasHistory` flag so the UI can disable or hide history toggles for licenses that have never been updated.

## Regression check

To confirm the dashboard tallies status buckets correctly, seed the sheet with a couple of contrasting expiry dates and allow the script to recompute the status columns:

| Field | Record A (expired) | Record B (upcoming) |
| --- | --- | --- |
| `name` | `Expired sample` | `Upcoming sample` |
| `expiryLabel` | `Trade license` | `Insurance renewal` |
| `expiryDate` | `2023-01-01` | _30 days from today_ |

After refreshing the dashboard the Expired card increments by one (record A) and the Upcoming card increments by one (record B). Clearing the status cells keeps the calculation server driven and highlights any regressions in the bucketing logic.

### Invalid payload fallback

To guard against regressions where the Apps Script response is missing or malformed:

1. Open the deployed web app in a browser tab and launch the developer tools console.
2. Run `google.script.run.withSuccessHandler(cb => cb({})).getDashboardData('')` in the console to simulate a refresh that returns an empty object.
3. Confirm the dashboard displays the fallback/debug banner instead of throwing an exception and that the rest of the UI remains interactive.
