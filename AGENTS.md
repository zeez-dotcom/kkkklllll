# Repository Agent Guide

## Project Purpose
- This project implements an **Admin Office Licenses** manager intended to run as a Google Apps Script web app.
- The goal is to display, search, and upload license documents stored in Google Drive while tracking expiry information in a Google Sheet.
- The solution is bilingual (English/Arabic) and the UI is rendered from `index.html`, while business logic lives in `code.gs`.

## High-Level Architecture
1. **Google Apps Script backend (`code.gs`)**
   - Serves the web UI via `doGet()`.
   - Reads and writes license data in the `Licenses` sheet of the active spreadsheet.
   - Handles dashboard data retrieval (`getDashboardData`) and upload processing (`uploadDocument`).
   - Stores uploaded files in Drive (optionally inside a dedicated folder) and shares them for public viewing when `SHARE_FILES_PUBLIC` is true.
2. **Client-side web app (`index.html`)**
   - Provides dashboard statistics, search filtering, multilingual text, and a file upload form.
   - Calls Apps Script services through `google.script.run.getDashboardData` and `google.script.run.uploadDocument`.
   - Renders results, handles status calculations, and previews uploaded files.

## Key Data Flows
- **Dashboard load**: `refresh()` (frontend) → `google.script.run.getDashboardData(query)` → Apps Script reads the sheet, computes status counts, and returns rows → frontend renders table and stats.
- **File upload**: User selects file → frontend converts to Base64 → `google.script.run.uploadDocument(payload)` → Apps Script writes sheet row, saves Drive file, returns metadata → frontend shows success, resets form, refreshes dashboard.

## Important Constants & Config
- `SHEET_NAME`: Name of the Google Sheet tab containing license records.
- `HEADER`: Column order expected in the sheet; maintained by `getSheet_()`.
- `FOLDER_ID`: Optional Drive folder ID; when blank a folder named *Admin Office Licenses* is created.
- `SHARE_FILES_PUBLIC`: Controls whether uploaded files become "Anyone with link → Viewer".
- `MAX_UPLOAD_SIZE_BYTES`: Upload size guard enforced on the server.

## Extending the Project
- **Adding fields**: Update both `HEADER`/normalization logic in `code.gs` and the form/table bindings in `index.html`.
- **Localization**: Translations live inside `index.html` in the `translations` object; every string requires both English and Arabic keys.
- **Validation**: Client-side validation uses built-in form validity plus helper trimming. Server-side normalization occurs in `normalizeUploadInput_()`.
- **Deployment**: Publish as a Google Apps Script web app; ensure the linked spreadsheet and Drive folder permissions match your organization’s policy.

## Testing Tips
- Run `getDashboardData` and `uploadDocument` manually from the Apps Script editor with sample payloads to verify sheet/Drive integration.
- Use browser developer tools to inspect network calls from `index.html` when embedded in the Apps Script environment; the `google` object is only available there.
- For local HTML previews, mock the `google.script.run` interface to avoid runtime errors.

## File Inventory
- `index.html`: Frontend UI, translations, and client logic.
- `code.gs`: Apps Script backend handling data storage and Drive interactions.
- `README.md`: Minimal placeholder; update with deployment instructions if needed.

## Contribution Guidelines
- Maintain parity between English and Arabic UI text.
- Keep Apps Script functions free of global state beyond documented configuration constants.
- When adding Drive interactions, sanitize URLs via `sanitizeUrl_()` to avoid invalid links.
- Prefer pure functions where possible for easier unit testing/mocking.
