## Warehouse Dashboard Module

This Google Apps Script module renders the warehouse dashboard (`index.html`) and the supporting tracking console.

### File layout

| File | Purpose |
| ---- | ------- |
| `code.gs` | Apps Script backend. Reads the `All Orders` sheet, normalises rows, exposes dashboard data, and records tracking events + media. |
| `index.html` | Base HTML template plus embedded `<style>` + `<script>` blocks that render the dashboard, exports, and tracking console. |

### Front-end structure

- The UI is initialised from `WarehouseApp.init()` on `DOMContentLoaded`.
- Event listeners are registered programmatically (no inline attributes) so behaviour lives next to implementation.
- Server-provided data (`allOrders`, `maintenanceOrders`) is parsed once into a `dataStore` object; rendering functions receive filtered slices to keep logic small and composable.
- Tracking state from Sheets (orders/items/events/media) is fetched only when the page runs inside Apps Script (`google.script.run`) and cached in `trackingState`.

### Extending tips

1. **Add columns/fields**
   - Update the sheet constants in `code.gs` (`TRACKING_HEADERS`, etc.).
   - Extend the rendering helpers inside the `<script>` block in `index.html` to read/print the new fields.
   - Keep translations/user text in that script block within small helper functions so additions stay centralised.
   - Product-related columns are auto-detected from the sheet header row via `PRODUCT_HEADER_KEYWORDS` in `code.gs`. If you rename those headers, add the new label patterns there so product names/categories don’t fall back to “Unnamed Product (Unknown)”.
2. **New dashboard widgets**
   - Derive filtered arrays in `filterData()` and pass them to a dedicated `renderXYZ()` helper.
   - Reuse DOM caching via the `dom` map to avoid repetitive `document.getElementById` calls.
3. **Styling changes**
   - The `<style>` block inside `index.html` is grouped by component (layout, tables, tracking console, statuses, print styles) to keep edits easy to scan.

With this structure the folder stays approachable—HTML holds only markup, CSS is grouped per component, and the JS module is an easy entry point for future enhancements.
