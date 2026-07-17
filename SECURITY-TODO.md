# Security & Safety TODO

Findings from a full security/safety review of Summa (Flask + SQLite backend, vanilla-JS PWA frontend, Docker deployment). Ordered by severity. Checkboxes track remediation.

The app is designed as a personal, single-user tool with no auth, exposed via Docker on port 8000 with global CORS — several findings follow directly from that trust model and their fix depends on the intended exposure (LAN-only vs. internet-facing).

## High

- [ ] **H1 — No auth/authorization, exposed on all interfaces with wildcard CORS.**
      `CORS(app)` in [summa/**init**.py:33](summa/__init__.py#L33) sends `Access-Control-Allow-Origin: *` on every route, and [docker-compose.yml:6-7](docker-compose.yml#L6-L7) publishes `8000:8000` (binds `0.0.0.0`). Anyone with network reachability can read, modify, import and soft-delete all financial data. Because CORS is wildcard, _any website_ opened on a machine that can reach the server (same LAN, `localhost`) can silently read and mutate the data from JavaScript.
      **Fix:** bind to `127.0.0.1:8000:8000` / VPN, restrict CORS to an env-driven allowlist (keep `*` only as an explicit opt-in for native clients), and/or add a shared-token check on `/api/*` writes or reverse-proxy auth.

- [x] **H2 — Stored XSS via attribute injection in the edit modal.**
      [static/js/modals.js:19](static/js/modals.js#L19): `value="${escapeHtml(item.item_name)}"`. `escapeHtml()` ([static/js/dom.js:92](static/js/dom.js#L92)) uses `textContent → innerHTML`, which escapes `& < >` but **not** double quotes. An item name like `" autofocus onfocus="fetch(...)` breaks out of the `value="…"` attribute and injects an event handler that fires when the invoice is edited. Item names are attacker-controllable via the unauthenticated API (H1) and via JSON import files. Other `innerHTML` sites interpolate escaped text into _element content_, where the missing quote-escape is harmless — this attribute context is the one real hole.
      **Fix:** make `escapeHtml()` also encode `"` and `'` (explicit `String(text).replace(...)` over `& < > " '`). One function, fixes every attribute context. Bump `CACHE_NAME` in [static/sw.js](static/sw.js#L6).

## Medium

- [x] **M1 — No security headers / CSP.**
      No `Content-Security-Policy`, `X-Content-Type-Options`, `X-Frame-Options`/`frame-ancestors`, `Referrer-Policy`. A CSP would neutralize H2's exfiltration and constrain any future XSS.
      **Fix:** add an `after_request` hook in `create_app()` setting CSP (`default-src 'self'`), `X-Content-Type-Options: nosniff`, `X-Frame-Options: DENY`, `Referrer-Policy: no-referrer`.

- [x] **M2 — Third-party assets without SRI.**
      Chart.js from jsdelivr ([templates/index.html:109-112](templates/index.html#L109-L112)) and Google Fonts, no `integrity` attribute. A CDN compromise executes arbitrary JS in an app holding all financial data; also leaks the user's IP/UA to Google/jsdelivr on every visit and breaks the PWA offline (these assets are not in the SW cache).
      **Fix:** self-host Chart.js under `static/js/vendor/` (auto-picked-up by the JS manifest/SW) and self-host the fonts; enables a strict `default-src 'self'` CSP.

- [x] **M3 — No request-size limit or import bounds.**
      `MAX_CONTENT_LENGTH` unset; `/api/invoices/import` parses arbitrarily large JSON fully into memory and inserfiles are detected immediatelyts unbounded rows (per-row `SELECT` + `INSERT`). One large request can exhaust memory/disk. No rate limiting anywhere.
      **Fix:** set `app.config["MAX_CONTENT_LENGTH"]` (e.g. 5 MB); optionally cap the import batch size in `parse_invoice_batch`.

- [x] **M4 — Raw exception text returned to clients.**
      Write handlers return `error_response(str(e), 500)` on `sqlite3.Error` (e.g. [summa/routes/invoices.py:239](summa/routes/invoices.py#L239)), leaking schema/internal detail.
      **Fix:** return a generic `"Internal server error"`, keep `logger.error` with details server-side.

- [x] **M5 — Unvalidated JSON shape on bulk endpoints.**
      `bulk_update_invoices` / `bulk_delete_invoices` call `data.get(...)` on `request.json` without checking it is a dict and never validate that `ids` is a list of ints ([summa/routes/invoices.py:356-431](summa/routes/invoices.py#L356-L431)) → unhandled `AttributeError`/type confusion → HTML 500 that bypasses the JSON error convention. (SQL is parameterized, so no injection.)
      **Fix:** require a dict body and a non-empty list-of-ints `ids`, reusing the `ValidationError`/`error_response` pattern from `summa/helpers.py`.

- [x] **M6 — Dev server hardcodes `debug=True`.**
      [summa/**init**.py:48](summa/__init__.py#L48). The Werkzeug debugger is RCE-by-design if the dev server is ever reachable by others.
      **Fix:** `debug=os.environ.get("FLASK_DEBUG") == "1"`.

## Low

- [x] **L1 — Soft-delete blind spots.**
      `PUT /api/invoices/<id>`, `DELETE` and `bulk-update` do not filter `deleted_at IS NULL`, so "deleted" rows can still be modified; the import duplicate check ([summa/routes/invoices.py:263](summa/routes/invoices.py#L263)) also matches deleted rows, silently blocking re-import. No purge path (data kept forever).
      **Fix:** add `AND deleted_at IS NULL` to those WHERE clauses and the duplicate check.

- [x] **L2 — No TLS in the deployment.**
      gunicorn serves plain HTTP; fine behind a reverse proxy, but nothing in the repo documents/enforces that.
      **Fix:** document/require a TLS-terminating reverse proxy.

- [x] **L3 — SW caches API responses on device.**
      [static/sw.js](static/sw.js#L118) network-first caches all `/api/` GETs in Cache Storage — financial data persists unencrypted on any device that loaded the app. Acceptable for a personal device; worth documenting or excluding `/api/` from caching.

## Positive (no action needed)

- SQL consistently parameterized; `sort_by` whitelisted; LIKE wildcards escaped (`escape_like`); page params clamped (`parse_bounded_int`); `IN (...)` built only from counted placeholders.
- Frontend escaping is otherwise diligent and deliberate (documented "safe to inline" invariants; category colors emit only CSS-var references).
- Docker: multi-stage build, frozen lockfile, prod-only deps, non-root `appuser` via gosu, no secrets in repo.
- Dependencies current (Flask ≥3.1.2, flask-cors ≥6.0.2 incl. the 2024 flask-cors CVE fixes, gunicorn ≥25).

## Verification

- Backend: `uv run pytest`, `uv run ruff check .`, `uv run mypy`; add tests for bulk-payload validation and soft-delete filters.
- H2 fix: create an item named `" autofocus onfocus="window.__xss=1` via the API, open the edit modal, confirm the input shows the literal text and no attribute injection occurs.
- Headers/CORS: `curl -i http://localhost:8000/api/invoices` and check `Access-Control-Allow-Origin` / security headers.
- Frontend: `npm run lint`; bump `CACHE_NAME` for any changed static asset.
