# Werkly OMS — Agent Guide

This guide documents how the Werkly Order Management view works in this folder, plus the complete Warehouse Tracking logic cycle that augments orders with outbound/inbound photo evidence. The scope of this AGENTS.md is the entire directory rooted at `tracking/werklyoms`.

## Purpose
- Render a dated orders dashboard from the `Order` sheet.
- Provide fast search and export/print utilities.
- Track warehouse handover by capturing per‑item photos when orders leave (outbound) and when items return (inbound).

## Files
- `werkly_app.gs` — Web app entry point; builds the template context and wires HTML partials together.
- `werkly_orders.gs` — Order data loader: normalises sheet rows into rich order objects.
- `werkly_warehouse.gs` — Warehouse logging API and Drive/sheet orchestration (prefixed `WERKLY_…` config variables to avoid collisions).
- `werkly_utils.gs` — Shared helpers for formatting, date/time handling, and Drive utilities.
- `werkly_index.html` — Base layout that includes other HTML fragments at render time.
- `werkly_styles.html` — CSS token definitions and layout styling.
- `werkly_body.html` — Markup for the dashboard, tables, modal, and controls.
- `werkly_scripts.html` — Client-side logic, translations, and Apps Script calls.

## Data Sources
- Orders come from the active spreadsheet tab named `Order`.
  - Column usage (0‑based indices from `werkly_orders.gs`):
    - `0` Order Number
    - `4` Location
    - `6` Customer Name
    - `8` Comments
    - `9` Phone
    - `11` Payment Method
    - `14` Grand Total (number)
    - `20` Product Name (also used to detect cancellations: contains "CANCEL - كنسل")
    - `21` Product Image (URL)
    - `22` Product Category
    - `24` Tarkeeb Date
    - `25` Tarkeeb Time
    - `26` Sheel Date
    - `27` Sheel Time

## Rendering Flow
1. `doGet()` in `werkly_app.gs` reads the `Order` sheet via `loadOrderDataset_()` (in `werkly_orders.gs`) and normalizes rows into order objects with a `products` array (one order may span multiple rows/products).
2. Maintenance orders are detected (product name includes "Maintenance"/"صيانة") and enriched with a date range.
3. The template `werkly_index.html` (which includes `werkly_body.html` and `werkly_scripts.html`) is evaluated with these injected JSON blobs:
   - `allOrders` — all non‑maintenance and maintenance orders (deduped by order number).
   - `maintenanceOrders` — subset flagged as maintenance.
4. The frontend defaults the date to today, then filters `allOrders` by selected date (tarkeeb or sheel date) and optional text query.

## Warehouse Tracking — End‑to‑End Cycle
Goal: For each product on an order, capture a photo when it leaves the warehouse (outbound) and when it returns (inbound). Persist these with operator/timestamp, and surface quick status back into the orders tables.

### Backend (Apps Script)
Implemented in `werkly_warehouse.gs`:

- Configuration constants:
  - `WERKLY_WAREHOUSE_SHEET_NAME = 'WarehouseLog'`
  - `WERKLY_WAREHOUSE_ROOT_NAME = 'OMS Warehouse Photos'`
  - `WERKLY_WAREHOUSE_ROOT_ID = ''` (optional override; blank creates/finds by name in My Drive)
  - `WERKLY_WAREHOUSE_SHARE_PUBLIC = false` (set true to share files as Anyone‑with‑link → Viewer)

- Persistence model:
  - Drive folder hierarchy: `OMS Warehouse Photos/<OrderNumber>/outbound` and `.../inbound`.
  - Log sheet: A tab `WarehouseLog` with columns `[timestamp, operatorEmail, orderNumber, productIndex, productName, direction, photoUrl, notes]`.

- API surface:
  - `getWarehouseLogs(orderNumbers: string[]): Map<string, LogEntry[]>`
    - Reads `WarehouseLog`, returns only entries matching requested order numbers.
  - `uploadWarehousePhoto(payload)`
    - `payload = { orderNumber, productIndex, productName, direction: 'out'|'in', fileName, mimeType, base64, notes }`
    - Decodes `base64`, writes to Drive (under order folder + direction), shares if configured, appends a row to `WarehouseLog`, and returns the new log entry.

- Helpers:
  - `ensureWarehouseSheet_()` bootstraps the log sheet with header.
  - `getOrCreateWarehouseRootFolder_()`, `getOrCreateWarehouseOrderFolder_(orderNumber, direction)` manage Drive folders.

### Frontend (HTML/JS)
Implemented across `werkly_body.html` and `werkly_scripts.html`:

- Tables:
  - Tarkeeb/Sheel tables include a `Warehouse` column.
  - Each row shows a quick status badge: `X/Y out • A/Y in` where `Y` is product count.
  - An `Update` button opens a modal for that order.
- Additional tabs:
  - **Tarkeeb Items** — product-level view filtered by tarkeeb date range with outbound/inbound status summaries.
  - **Sheel Items** — mirrors the tarkeeb tab but keyed on sheel/return dates.
  - **Items Currently Out** — highlights items with outbound logs (within a selectable range) that are not yet recorded as returned.
  - **Return Compliance Report** — date/time-based report listing items past their scheduled sheel date/time without an inbound log, including overdue duration.
- UI preferences:
  - Language toggle (English ↔ Arabic) that localises all labels, messages, and aria attributes on the fly.
  - Body direction flips to RTL for Arabic.
  - Dark/light mode toggle backed by CSS variables and persisted in `localStorage`.

- Modal (`Warehouse Update — Order <number>`):
  - Shows each product with thumbnail and two panels:
    - Outbound: latest photo + operator/timestamp, choose new file (camera), optional notes, Upload.
    - Inbound: same pattern.
  - Inputs use `accept="image/*" capture="environment"` to hint camera usage on mobile devices.

- Data flow:
  1. Orders render immediately from `allOrders` (so original data always appears).
  2. Visible order numbers are collected and sent to `google.script.run.getWarehouseLogs(...)` asynchronously.
  3. On success, logs are kept in a client map `warehouseLogs{ orderNumber -> LogEntry[] }` and the tables are re-rendered to update quick counts.
  4. Upload:
     - Reads the image as DataURL → base64, posts to `uploadWarehousePhoto(payload)`.
     - On success, merges the returned entry into `warehouseLogs`, re-renders modal and tables.
  5. A one-time `loadAllWarehouseLogs()` call populates the warehouse log cache for all known order numbers (guarded to run only when Apps Script is available), enabling the additional tabs to work from the same data source.
  6. Translation/theme toggles re-run the render helpers so table/tab content is refreshed in the selected language/theme without reloading the page.

- Compatibility choices:
  - Uses modern JS features (`flatMap`, `Set`)—ensure the Apps Script web runtime is used for execution.
  - Calls to `google.script.run` are gated through `hasAppsScript()` so local (non-Apps Script) previews still render orders even though warehouse status is disabled.

### Status Semantics
- `productIndex` is the zero‑based index of the product within the order’s `products` array as rendered in the modal.
- For quick counts, an index is considered complete for a direction when at least one log entry exists for that `(orderNumber, productIndex, direction)` pair. The latest entry per pair is displayed in the modal.
- If no logs exist for an order, quick status shows `0/Y out • 0/Y in`. Orders still render normally.
- The per-item tabs reuse the same latest-log logic to show outbound/inbound status chips; the reporting tab considers items overdue when the scheduled sheel date/time has passed but no inbound log exists before the selected “as-of” timestamp.

## Configuring Behavior
- Make photos public: set `WERKLY_WAREHOUSE_SHARE_PUBLIC = true` in `werkly_warehouse.gs` if images need to be viewable without auth.
- Use a specific Drive folder: set `WERKLY_WAREHOUSE_ROOT_ID` to a folder ID you control; otherwise a folder by name is created at My Drive root.

## Deployment
- Publish the Apps Script as a web app (execute as you; access within your organization per policy).
- Ensure the script has Drive and Spreadsheet scopes to create folders/files and append rows.
- First upload creates the `WarehouseLog` sheet and the root Drive folder if missing.

## Testing Checklist
- Orders rendering:
  - Verify today’s date shows expected orders.
  - Use the search box to filter by name/phone/order/location/product.
- Warehouse log fetch:
  - Confirm quick status appears after a moment for visible orders.
- Upload flow:
  - For an order with 2 products, upload Outbound for both, then Inbound for one.
  - Expect status like `2/2 out • 1/2 in` and see thumbnails/timestamps in modal.
- Export/Print still function and do not include modal controls.

## Troubleshooting
- Empty tables: check date selection; open browser console for errors (syntax errors can block rendering).
- No status: ensure `WarehouseLog` tab exists or attempt an upload to create it; verify Drive folder permissions if `WERKLY_WAREHOUSE_SHARE_PUBLIC` is true.
- Slow uploads: consider adding client‑side image compression if warehouse network is constrained.

## Contribution Guidelines (folder scope)
- Don’t block order rendering with tracking failures: UI must render orders first, then layer tracking.
- Keep frontend JS compatible (avoid modern features that break on older devices used in the warehouse).
- When modifying logging schema, update both `werkly_warehouse.gs` and the script logic in `werkly_scripts.html` that consumes log entries.
- Keep Drive paths stable: changing folder naming affects existing links—migrate carefully if needed.
