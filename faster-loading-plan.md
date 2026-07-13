# Plan: Faster Invoice Loading

Current latency most likely comes from two sources: the API endpoint always returns the full invoice set including all items, and the frontend immediately renders a large HTML block from it. In save/delete flows, there is also a full refresh of all lookup data. Since UI changes are allowed, the most effective approach is to combine server-side limits with incremental loading instead of loading everything at once.

## Steps

1. Identify the real bottleneck with a short baseline measurement.
   Measure timings separately for `GET /api/invoices`, `GET /api/stores`, `GET /api/categories`, and DOM rendering in `renderInvoices()`. Use the same parameters with about 100 invoices and document the current duration. This is the reference for later validation.
2. Introduce server-side pagination for the invoice list.
   Extend `GET /api/invoices` with `page`/`page_size` or a cursor equivalent, and return only visible list rows plus metadata such as total count. This immediately reduces payload size and rendering cost.
3. Stop loading item-level data for every invoice in the initial listing.
   For the default view, a compact list without all `invoice_items` is enough. Load detailed line items only when an invoice is expanded or through a separate detail request. This reduces both SQL workload and HTML volume.
4. Simplify the frontend rendering path.
   Update `loadInvoices()` and `renderInvoices()` to render paginated data cleanly without rebuilding the full list. Use a small, stable pagination control and keep DOM size per page intentionally low.
5. Reduce unnecessary refreshes after save/edit/delete.
   Replace `refreshAllData()` in save/delete flows with targeted reload paths: refresh invoice list, and reload lookups only if store or category actually changed. This applies to create, single-entry edits, and bulk actions so the UI responds faster after each mutation.
6. Add suitable database indexes for the new access patterns.
   Prioritize indexes on `invoices.deleted_at`, filter/sort columns, and `invoice_items.invoice_id`. If searching by item names remains important, verify that the current search strategy is still performant enough.
7. If list view performance is still not sufficient, add a second stage with lazy loading for detail areas.
   Load invoice detail rows on first expand click or by visibility-based loading so larger pages remain smooth.

## Relevant Files

- `summa/routes/invoices.py`: `GET /api/invoices` is the central data provider; this is where pagination, compact list payloads, and detail loading belong.
- `summa/db.py`: schema and index adjustments, plus the foundation for faster queries.
- `static/js/api.js`: controls list refresh, lookups, and later incremental loading paths.
- `static/js/render.js`: currently builds the full invoice view as one large HTML string; this is where paginated/compact rendering should be implemented.
- `static/js/invoices.js`: save/delete flows that currently trigger full refresh behavior after each mutation.
- `static/js/state.js`: stores additional pagination and detail-loading state.
- `tests/test_invoices_api.py`: coverage for pagination, compact response shape, and detail loading.
- `tests/test_web.py`: coverage for UI-adjacent data delivery if route contracts change.

## Verification

1. Measure loading time before and after the change for a list of around 100 invoices; compare API time, JSON size, and browser render duration.
2. Verify that one page of invoice list renders quickly on initial load and that pagination switches correctly between pages.
3. Re-test save/delete and ensure only necessary data is reloaded and the UI responds noticeably sooner.
4. Run existing tests for invoice API and web routes, and add new tests for pagination and reduced response payloads.

## Decisions

- The UI is allowed to change, so pagination or lazy loading is acceptable.
- Priority is a real reduction in data volume and DOM size; micro-optimizations alone are not enough.
- The current full loading of all invoice line items in the initial view is treated as the main root cause.

## Further Considerations

- Should the default view use pagination or infinite scroll? Recommendation: pagination, because it keeps DOM size clearly bounded and is easier to test.
- Should detail line items load on first click, or earlier on hover/viewport proximity? Recommendation: load on first expand click to keep complexity low.
