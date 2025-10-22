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
