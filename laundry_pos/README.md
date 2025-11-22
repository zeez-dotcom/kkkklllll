# Laundry POS

Point-of-sale style Google Apps Script to log laundry orders with per-item services and automatic price lookups.

## Sheets created
- `LaundryOrders`: order-level info (customer, walk-in/exit dates, totals).
- `LaundryOrderItems`: line items with product, service, quantity, unit price, and line total.
- `Catalog`: drives the price book and options for items/services (product keys, labels EN/AR, service codes, prices, image URLs, active flag).
- `Settings`: key/value store (currently used for a global logo URL that appears in the UI and receipts).
- `Customers`: customer directory (id, name, nickname, phone, notes).
- `Subscriptions`: packages tied to a customer and a specific product/service with free allowance tracking.

Headers are auto-created on first run.
Catalog header order: `product_key, service_code, product_label_en, product_label_ar, service_label_en, service_label_ar, price, image_url, active`.

## Default price book
- Dishdasha: wash (0.350), wash & iron (0.500), iron (0.250)
- Abaya: wash (0.400), wash & steam (0.600), spot clean (0.300)
- Bedding set: wash (0.750), wash & iron (1.000), deep clean (1.250)
- Custom item/service: free-form pricing

Update the `Catalog` sheet (or `DEFAULT_CATALOG` seed in `code.gs`) to change labels, prices, and images.

## Flows
- Frontend (`index.html`) collects customer info and service lines, previews totals, and lists saved orders.
- Backend (`code.gs`) stores orders/items, enforces basic validation, and returns stats for the dashboard.
- Catalog editor lives on its own tab in the UI and reads/writes the `Catalog` sheet so prices, images, and services stay in sync with the sheet-driven price book. Logo URL is edited here and saved to `Settings`.
- Customers tab lets you add/search customers and create subscriptions (per product/service free allowance). POS flow can link orders to a customer/subscription so free allowances decrement automatically.
- Dashboards: Order management (status counts) and reporting (revenue, avg ticket, top customer, last 7 days) render from fetched data.

Deploy as a web app in Apps Script; the UI calls `getOrders` and `recordOrder` via `google.script.run`.
