# Werkly Order Tracking Console

This folder now layers a lightweight logistics console on top of the legacy Werkly report without altering the existing dashboards.

## Data model

All tracking data is persisted in four auto-provisioned sheets (created the first time any tracking API runs):

| Sheet | Purpose | Header |
| --- | --- | --- |
| `TrackingOrders` | One row per order with UUID, customer snapshot, current status, and the last recorded event. | `order_uuid, order_number, created_at, customer_name, phone, location, status, last_event_type, last_event_at, last_event_uuid, notes` |
| `TrackingOrderItems` | One row per item/SKU per order along with its last known event metadata. | `item_uuid, order_uuid, order_number, item_label, item_category, quantity, unit, last_event_uuid, last_event_type, last_event_at` |
| `TrackingOrderEvents` | Append-only ledger capturing every outbound/return/delivery event. Each entry points back to the order, optional item, and media bundle. | `event_uuid, order_uuid, order_number, timeline_group, item_uuid, item_label, event_type, status_after, driver_name, notes, location_note, media_urls, media_drive_ids, recorded_at, recorded_by` |
| `TrackingOrderMedia` | Index of Drive files uploaded for a given event. | `media_uuid, event_uuid, order_uuid, order_number, item_uuid, file_name, mime_type, drive_id, file_url, uploaded_at` |

`Utilities.getUuid()` is used everywhere so nothing depends on row numbers. Media is stored under a Drive folder named **Order Tracking Media** (auto-created unless `TRACKING_FOLDER_ID` is set) and the files are shared as *Anyone with the link → Viewer* when allowed by the domain policy.

## Apps Script services

Two new entry points are available to the UI (and future automations):

- `getTrackingData()` returns the latest orders, items, events, and media rows (plus the list of supported event types) so the frontend can render dropdowns and timelines.
- `recordOrderEvent(payload)` validates the request, resolves the backing order from the legacy `Order` sheet, ensures the tracking sheets exist, writes/updates order + item metadata, uploads optional photos, and appends one ledger row per affected item. Supported events map directly to order statuses: `OUT_FOR_DELIVERY`, `DELIVERED`, `RETURNED_TO_WAREHOUSE`, and `CANCELLED`.

Helper utilities take care of:

- Auto-creating the tracking sheets with the expected headers.
- Generating UUIDs for orders, items, events, and media.
- Linking Drive uploads back to both the event row and the media index sheet.
- Updating the `TrackingOrderItems`/`TrackingOrders` sheets with the most recent event metadata so dashboards can query current status without scanning the ledger.

## Frontend experience

The legacy summary/location/order tables remain untouched. Beneath them, the **Fulfillment Tracking Console** introduces:

- An order selector that merges the existing `allOrders` payload with the new tracking metadata so you always log against the right order number.
- An event form (driver, timestamp, location note, internal notes, per-item checkboxes, and up to 5 photos ≤5 MB each) that posts to `recordOrderEvent`. Leaving all items unchecked applies the event to the full order.
- Quick-select buttons to mark all/none of the items and a reset action that clears the form without disturbing the rest of the dashboard.
- A live timeline that groups rows by `timeline_group`, so multi-item events show once with a consolidated item list and media links.

All client-side calls gracefully no-op when the HTML is opened outside the Apps Script runtime (useful for local previews).

## Quick test plan

1. Deploy the Apps Script project, open the web app, and confirm the legacy Werkly report still loads.
2. Scroll to **Fulfillment Tracking Console**, pick any order, and log an **Out For Delivery** event with at least one item, driver name, and (optionally) a photo. The new tracking sheets should be created automatically.
3. Refresh the page; the order dropdown should now show the updated status and the timeline should list the event with your note/media link.
4. Repeat with a **Returned To Warehouse** event selecting only one item to see per-item ledger rows and status updates.

These steps validate the new sheets, Drive uploads, frontend wiring, and status propagation without touching the historical reporting views.
