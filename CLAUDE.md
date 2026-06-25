# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Overview

Summa is an invoice management and expense-tracking web app: a Flask REST backend
backed by SQLite, plus a vanilla-JS Progressive Web App frontend. There is no
build step for the frontend and no test suite.

## Code Style

See @docs/code-style.md for code style and convention rules (applies on every machine).

## Commands

Tooling is driven through `uv` (Python 3.12+):

```bash
uv sync                       # install deps + create .venv
uv run flask run --port 8000  # run dev server (DB at ./invoices.db)
uv run ruff format .          # format
uv run ruff check .           # lint (E, F, I; E501 intentionally ignored)
uv run mypy                   # strict type check (files set in pyproject.toml)
```

Frontend (JS/CSS/HTML) is linted and formatted through `npm` (Node 22):

```bash
npm install        # install the lint/format toolchain
npm run lint       # eslint + stylelint + prettier --check (what CI runs)
npm run format     # prettier --write across the repo
```

ESLint config (`eslint.config.js`) lints `static/js/app.js` (browser script) and
`static/sw.js` (service worker), and runs `@html-eslint` over `templates/*.html`
for semantic/a11y checks (its formatting rules are off — Prettier owns
formatting, via `prettier-plugin-jinja-template` for the Jinja template).
Stylelint (`.stylelintrc.json`) enforces the `docs/code-style.md` CSS rules
(recess-order property order, no `!important`, no id selectors, value hygiene).
Functions called only from inline HTML handlers are listed in a top-of-file
`/* exported … */` directive in `app.js` so `no-unused-vars` does not flag them.

CI (`.github/workflows/ci.yml`) has two jobs that must pass: `lint` (`ruff check
.`, `ruff format --check .`, `mypy`) and `frontend` (`npm run lint`). Run both
before pushing.

Docker: `docker compose up -d` serves on `http://localhost:8000` with the DB
persisted in the `summa_data` volume.

## Architecture

**Backend — `app.py` (single module).** All routes, DB access, and schema live
here. Key conventions:

- **SQLite with two tables:** `invoices` and `invoice_items` (FK with
  `ON DELETE CASCADE`). Connections come from `get_db()`, which sets
  `row_factory` and enables WAL mode. `DATABASE_PATH` env var overrides the
  default `invoices.db`.
- **Schema + migrations live in `init_db()`**, which runs on module import (so it
  works under both gunicorn and the dev server). Migrations are done inline by
  inspecting `PRAGMA table_info` and conditionally `ALTER TABLE`-ing new columns
  (e.g. `deleted_at`, `category`). Add future column migrations the same way.
- **Soft deletes:** rows are never physically deleted. Delete endpoints set
  `deleted_at = CURRENT_TIMESTAMP`, and every read query filters
  `WHERE deleted_at IS NULL`. Preserve this filter in any new query.
- **REST API under `/api/`** (invoices CRUD, `/import`, `/bulk-update`,
  `/bulk-delete`, `/stores`, `/categories`, `/stats`). Handlers return
  `Response | tuple[Response, int]` (the `ApiResponse` alias) and wrap writes in
  try/commit/except-rollback/finally-close. `strip_text()` normalizes input
  (empty string -> `None`). CORS is enabled globally for native mobile clients.

**Frontend — `static/js/app.js` + `templates/index.html`.** Plain JS (no
framework, no bundler) talking to the API; styling in `static/css/style.css`.

**PWA.** `static/sw.js` caches static assets under `CACHE_NAME`
(`summa-cache-v2`). **When you change any cached static asset, bump
`CACHE_NAME`** or clients keep serving the stale cached version. The service
worker is registered from `app.js`.

## Deployment

Multi-stage `Dockerfile`: the builder stage generates PWA icons via
`generate_icons.py` (Pillow, from the locked `icons` dependency group); the
runtime stage installs prod-only deps and runs `gunicorn` (2 workers, 4 threads)
as a non-root `appuser`. `entrypoint.sh` fixes `/data` volume ownership via
`gosu` before dropping privileges. In the container the DB lives at
`/data/invoices.db`.
