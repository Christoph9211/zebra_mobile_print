# Label-to-Website Catalog Draft Workflow

## Summary

Use successful label prints as a reviewed product-intake queue, not as a direct publishing system.

The printer already captures name, size, type, and price, but structured data currently survives only in device-specific `localStorage`; FastAPI discards those extra fields. The website already provides the required review/merge tool, so no admin backend or database is needed.

## Implementation Changes

- Extend the label form and `/print` payload with:
  - Website category
  - Optional parent product/group for “Vapes & Carts” and “Other”
  - “Add/update website draft” checkbox, off by default and reset after printing
- When draft capture is enabled:
  - Require name, category, size, and a simple numeric/currency price.
  - Require a parent group for grouped categories.
  - Reject invalid draft data before printing; ordinary printing remains unchanged when capture is disabled.
- After a successful print, append the structured candidate to an ignored `data/catalog-draft.jsonl` file on the office PC.
  - Use a file lock.
  - A storage failure must report “label printed, draft not saved” without encouraging a reprint.
- Add `GET /catalog-draft`, returning a downloadable website-compatible array:
  - Standard categories: printed name becomes product name; canonical size becomes `size_options`.
  - Vapes/Other: parent group becomes product name; printed item name becomes the option.
  - Group by case-insensitive name/category, combine variants, and let the newest print win for duplicate variants.
- Add a “Download Website Draft” action to the printer page.
- Import the current catalog into `tools/json_gen_v_2_FINAL.html`, merge the downloaded draft, review optional metadata, and export the final `products.json`.
- Run the existing Clover reconciliation before publication so Clover remains authoritative for prices and availability.

## Interfaces

- `/print` gains optional `category`, `catalog_group`, `size`, `price_input`, and `website_draft` fields; existing callers remain compatible.
- `/catalog-draft` returns:
  ```json
  [{"name":"Example","category":"Flower","size_options":["1/8 oz"],"prices":{"1/8 oz":25}}]
  ```
- Canonicalize known label sizes such as `3 gram` → `3 grams`; preserve unknown custom sizes after trimming.

## Test Plan

- Python unit tests using a temporary JSONL file:
  - Unflagged or failed prints are not recorded.
  - Invalid flagged drafts do not print.
  - Variants aggregate correctly and newest duplicate prices win.
  - Vapes/Other use the required parent grouping.
  - Malformed stored lines do not break the complete export.
- Browser verification from the office PC and a Tailscale client.
- Confirm the downloaded draft merges without validation errors in the existing product manager.
- After replacing the catalog, run `npm run lint` and `npm run build`.

## Assumptions

- Existing browser history is not backfilled; centralized capture begins after deployment.
- The JSONL queue is intentionally append-only and re-exportable; merging is idempotent.
- THCa percentage, images, descriptions, banners, and availability remain review-time fields.
- No database, cloud credentials, direct Git writes, or automatic deployment are added.
