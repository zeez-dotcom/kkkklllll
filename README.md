# Admin Office Licenses Web App

This Apps Script project exposes a web dashboard for tracking license records stored in a Google Sheet.

## Unified Sales & Cash Dashboard (unified_app)

`unified_app/code.gs` and `unified_app/index.html` combine the sales tracker and cash-in-hand ledger into one Apps Script web app:

- Provides two tabs (Sales / Cash) so operators can capture daily revenue, expenses, receipts, and cash movements from a single deployment.
- Reuses one spreadsheet: the script auto-creates `SalesRecords` and `CashLedger` tabs (with their expected headers) the first time each service is invoked.
- Handles file uploads for sales receipts (JPG/PNG ≤ 5 MB) and stores them under a Drive folder named *Sales Receipts*, optionally sharing files publicly when `SHARE_RECEIPTS_PUBLIC` is true.
- Surfaces KNET-specific actions: pending/overdue badges, mark-as-received, and profit-transfer lifecycle that logs the balancing cash-out for personal withdrawals.
- Returns summary cards for totals, pending KNET batches, profit transfers, and net income/cash-in-hand so the UI stays consistent across both tabs.

Deploy this bundle when you want a single web app that manages both workflows. Set `SALES_SPREADSHEET_ID`/`CASH_SPREADSHEET_ID` in `unified_app/code.gs` if the script is not bound to the desired spreadsheet.

## Sales & Expenses Dashboard (sales_app)

A parallel Apps Script/HTML bundle lives in `sales_app/code.gs` and `sales_app/index.html`. It mirrors the existing license dashboard but is focused on daily sales capture:

- Records total sales, KNET batches, cash sales, expenses, and free-form notes for each report date.
- Tracks expected KNET deposit dates (defaulting to 10 days after the report date when left blank) and allows marking deposits as received later on.
- Computes running KPIs (totals, net income, pending/overdue KNET) for the table and statistic cards.
- Exposes `getSalesDashboardData`, `recordSalesEntry`, and `markKnetReceived` for use by the frontend.
- Automatically creates the backing `SalesRecords` sheet with the expected header the first time any endpoint runs (no manual setup needed).
- Accepts an optional receipt image (JPG/PNG up to 5 MB). Uploaded receipts are stored in Drive under the `Sales Receipts` folder (created on demand) and linked back to each record. Control sharing via `SHARE_RECEIPTS_PUBLIC`.

Use this bundle only if you need the sales dashboard in isolation; otherwise prefer the unified app. Deploy it as a separate project and configure `SALES_SPREADSHEET_ID` when necessary.

## Cash In Hand Dashboard (cash_app)

`cash_app/code.gs` and `cash_app/index.html` provide a lightweight cash ledger focused on daily liquidity:

- Tracks money in/out with direction and category tags while keeping a running cash-in-hand figure.
- Monitors KNET inflows with expected deposit dates, late/overdue detection, and a one-click action to mark deposits received.
- Offers a "transfer to profits" control that marks income as withdrawn for personal use and automatically records the balancing outflow to close the lifecycle.
- Summaries highlight total inflows/outflows, pending KNET batches, and the amount that has already been moved to profits.

Keep this bundle if you want a standalone ledger; for most scenarios the unified build is the recommended path. Deploy it the same way as the other bundles, optionally setting `CASH_SPREADSHEET_ID` when the script isn't bound to the destination spreadsheet.

## Spreadsheet configuration

The server code expects to read and write license data from a spreadsheet that contains a sheet named `Licenses` with the header defined in `code.gs`.

1. Create (or identify) the spreadsheet that should back the dashboard.
2. Open the Apps Script project and set the `SPREADSHEET_ID` constant in `code.gs` to that spreadsheet's ID.
   - The ID is the long string between `/d/` and `/edit` in the spreadsheet URL.
   - Leave the constant blank only when the script is container-bound to the desired spreadsheet.
3. Save the script changes before testing or deploying the web app.

The `Licenses` sheet should expose these columns, in order:

```
id, name, nameAr, description, descriptionAr,
expiryLabel, expiryLabelAr, expiryDate, expiryStatus,
validityLabel, validityLabelAr, validityDate, validityStatus,
status, fileUrl, fileName, createdAt
```

If the ID is missing and the script is not bound to a spreadsheet, calls to `getDashboardData` will throw an error instructing you to set `SPREADSHEET_ID`.

## Deployment notes

When deploying as a web app, ensure the published script has access to the configured spreadsheet and, if necessary, update the Drive folder permissions for uploaded files.

## License renewals & history tracking

The backend exposes dedicated helpers for updating existing rows and for fetching their revision history:

* `updateLicense(payload)` replaces the editable cells (expiry labels/dates, status, and file metadata) for a given `id`. When a replacement file is supplied it is uploaded to Drive before the sheet is updated. The function recomputes status columns with `computeStatus_` and stores a snapshot of the pre-update values in `LicenseHistory`.
* `renewLicense(payload)` is a thin alias around `updateLicense` that always records the action as a renewal. Use this when the UI offers a "renew" workflow.
* `getLicenseHistory(id)` reads the `LicenseHistory` sheet and returns the previous values ordered from newest to oldest. The payload includes timestamps, the previous expiry and validity labels/dates, status labels, and file URLs so the frontend can render a detailed timeline.

The history sheet is initialised automatically with the `LICENSE_HISTORY_HEADER` column order. Every update (or renewal) appends a new row capturing the prior expiry/validity dates, status labels, Drive file link, and a timestamp of when the change occurred. Dashboard rows also include a `hasHistory` flag so the UI can disable or hide history toggles for licenses that have never been updated.

## Regression check

To confirm the dashboard tallies status buckets correctly, seed the sheet with contrasting expiry/validity dates and allow the script to recompute the status columns:

| Field | Record A (expired) | Record B (upcoming) |
| --- | --- | --- |
| `name` | `Expired sample` | `Upcoming sample` |
| `expiryLabel` | `Trade license` | `Insurance renewal` |
| `expiryDate` | `2023-01-01` | _30 days from today_ |
| `validityLabel` | `Operating permit` | `Operating permit` |
| `validityDate` | `2023-01-05` | _40 days from today_ |

After refreshing the dashboard the Expired card increments by one (record A) and the Upcoming card increments by one (record B). Clearing the status cells keeps the calculation server driven and highlights any regressions in the bucketing logic.

### Invalid payload fallback

To guard against regressions where the Apps Script response is missing or malformed:

1. Open the deployed web app in a browser tab and launch the developer tools console.
2. Run `google.script.run.withSuccessHandler(cb => cb({})).getDashboardData('')` in the console to simulate a refresh that returns an empty object.
3. Confirm the dashboard displays the fallback/debug banner instead of throwing an exception and that the rest of the UI remains interactive.
