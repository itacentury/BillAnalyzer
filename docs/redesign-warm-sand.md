# Summa "Warm Sand" Redesign — Implementation Plan

## Context

Summa's UI is currently a single hard-coded **dark** theme (`#0a0a0a` bg, blue `#3b82f6`
accent), a top **header** with view-toggle tabs + Import/New buttons, and a large
7-column filter block. The `README.md` handoff (with `Summa Redesign.dc.html` design
reference + screenshots `3a–3e`, `4b`) specifies a complete redesign to a **light,
warm-sand** theme: a left **sidebar** for navigation, a compact filter row with a
collapsible filter panel, category **badges/combobox** with per-category colors, and
two bug fixes (filter overflow, always-visible bulk toolbar). Fidelity is high — colors,
type, spacing and radii from the handoff are final and taken verbatim.

This is a frontend-only change (templates, CSS, JS, PWA icons); the Flask API and DB are
unchanged. No build step, no test suite — CI runs `ruff`/`mypy` (lint job) and
`npm run lint` (frontend job); **both must pass at every commit**.

### Decisions locked with the user

1. **Icons:** change the gradient source **and regenerate the PNGs now** (run
   `generate_icons.py`, commit the new binaries).
2. **Mobile actions:** on ≤640px the sidebar becomes the existing bottom tab bar;
   **New Invoice** becomes a floating action button (FAB); **Import** sits as a small
   secondary icon button in the mobile action cluster (overflow of the FAB).
3. **Category combobox scope:** replaces the category field in **Add/Edit + Bulk-Edit
   modals _and_ the filter panel** (Design 4b everywhere), not just the modals.
4. **Delivery:** ship as **independently-committable milestones** (below). Each milestone
   leaves the app working and CI green so work can be committed and resumed between them.
   The plan itself is committed into the repo (Milestone 0).

## Working agreement for milestones

- Each milestone = one focused, self-contained commit that **builds, runs, and passes
  both CI jobs** on its own.
- Milestones are ordered so the app is never left broken: tokens first (light theme over
  the old layout still works), then structure, then per-view restyles, then the new
  combobox, then PWA/SW.
- **Bump `CACHE_NAME` in `static/sw.js`** in any milestone that changes a cached CSS/JS
  asset (simplest: bump once per milestone that touches `static/`).
- Before each commit run: `uv run ruff format .` · `uv run ruff check .` · `uv run mypy`
  · `npm run lint`.

---

## Key architecture facts (verified)

- Vanilla ES modules under `static/js/`, wired via `data-el` (element handles) and
  `data-action` (delegated clicks) — no framework/bundler, no `window` globals. Entry:
  `app.js` → `init()` calls each module's `setup*Listeners()`.
- Value-bearing controls are read by `.value` on `data-el` hooks: `buildFilterParams()`
  (`api.js`) reads `store-filter`, `type-filter`, `sort-by`, `sort-order`;
  `saveBulkEdit()` (`bulk.js`) reads `bulk-edit-category`; the save-invoice path reads
  `invoice-type`. **Strategy: restyle/replace the visible control but keep a hidden
  value-bearing element with the same `data-el`** so these readers stay untouched.
- View toggle lives in `stats.js` (`showInvoicesView`/`showStatsView`,
  `setupStatsListeners`) reading `btn.dataset.view`, delegating on `.view-toggle`.
- Bulk-toolbar visibility is `updateBulkActionToolbar()` in `render.js` — it already
  only adds `.visible` when `count > 0` (so the JS half of the README "bugfix" is already
  correct; we only harden the CSS with `visibility/opacity`).
- Collapsible filter panel already exists: `[data-el="filters-collapsible"]` animated via
  `grid-template-rows:0fr→1fr` (`filters.css`), toggled by `toggleAdvancedFilters()` in
  `stats.js` (currently mobile-only) — reuse it for the desktop panel.
- Data: `GET /api/categories` (flat `string[]`), `/api/stores`, `/api/stats`
  (`by_category`/`by_store`). No API/schema change.
- `sw.js` precaches JS/CSS discovered via `/static/js-manifest.json` +
  `/static/css-manifest.json` globs — new/renamed assets auto-register; only
  `CACHE_NAME` must be bumped.

---

## Milestone 0 — Plan in repo

**Commit:** `docs: add warm-sand redesign plan`

- Copy this plan into the repo at **`docs/redesign-warm-sand.md`** so it can be committed
  and used to resume work. (No code changes; nothing to lint beyond markdown.)

## Milestone 1 — Design tokens + fonts (light theme over existing layout)

**Files:** `static/css/variables.css`, `static/css/base.css`, `templates/index.html`
**Commit:** `feat(ui): warm-sand design tokens and Sora font`

- Replace the dark `:root` block in `variables.css` with the README warm-sand token set
  (surfaces, lines, text, accents, `--cat-*`, `--chart-1…8`, radii `sm/md/lg/xl`,
  shadows, `--transition`). Overlapping names keep their name with new values
  (`--bg-primary/-card`, `--border-color/-subtle`, `--text-primary/-secondary/-muted`,
  `--accent`, `--radius-*`, `--shadow-sm/md`, `--transition`); **add** the new ones
  (`--bg-sidebar/-inset/-hover/-selected`, `--border-strong/-dashed`, `--text-faint`,
  `--accent-text-on`, `--amount`, `--danger-*`, `--cat-*`, `--chart-*`, `--radius-xl`,
  `--shadow-dropdown/-modal/-toolbar`).
- **Compat shim:** keep `--accent-hover` (`#b87d5c`), `--accent-subtle`, `--success`/
  `--success-subtle`, `--danger`/`--danger-subtle` (re-derived warm) so `components.css`
  and existing buttons don't break before their restyle.
- `base.css` + `index.html`: body/UI font DM Sans → **Sora**; update the Google Fonts
  link to `Sora:wght@400;500;600;700&family=JetBrains+Mono:wght@400;500;600`.
- `index.html`: `<meta name="theme-color">` → `#c98d6b`.
- **Done when:** app renders in the light warm-sand palette with the _existing_ layout;
  no visual breakage; CI green.

## Milestone 2 — Sidebar shell + navigation

**Files:** `templates/index.html`, `templates/partials/sidebar.html` (new, ← `header.html`),
`static/css/sidebar.css` (new, ← `header.css`), `static/js/stats.js`
**Commit:** `feat(ui): left sidebar navigation and app shell`

- `index.html`: wrap content in `<div class="app">` (flex) = sidebar + `<main
class="container">`; swap the header include for `sidebar.html`; swap the
  `header.css` link for `sidebar.css`.
- Build `sidebar.html` per handoff §1 (232px, logo, nav with `data-view`, bottom
  Import/New actions with existing `data-action="open-import"|"open-add"`). Reuse the
  file-text/bar-chart SVGs.
- `stats.js`: rename the view-toggle selectors to the sidebar nav container/`.nav-item`
  (keep the `dataset.view` logic and active-class toggle).
- `sidebar.css`: desktop sidebar + **≤640px bottom tab bar** (port the mobile
  `.view-toggle` rules currently in `stats.css`) + bottom-right **action cluster**
  (New FAB + Import icon) for mobile.
- **Done when:** Invoices↔Statistics switch from the sidebar; mobile shows tab bar + FAB;
  CI green.

## Milestone 3 — Filters: compact row + panel (+ overflow bugfix)

**Files:** `templates/partials/filters.html`, `static/css/filters.css`,
`static/js/filters.js`, `static/js/stats.js`
**Commit:** `feat(ui): compact filter row with collapsible panel; fix overflow`

- Compact row (Design 3a): period pills (`data-filter`), month navigator, search
  (`data-el="search"`), **Filter button** (`data-el="filters-toggle"`) with count badge.
- **Delete** the 7-col grid + its `@media(width<=1400px)` collapse (**overflow bugfix**);
  panel becomes a 4-col card (`--shadow-dropdown`) reusing the `grid-template-rows`
  collapsible, now on desktop too.
- Panel cells: Store (`select`), **Category stays a styled `<select data-el="type-filter">`
  for now** (→ combobox in Milestone 6), From/To, **Sort By / Order as pill groups**
  backed by hidden `data-el="sort-by"/"sort-order"` inputs (keeps `buildFilterParams`
  unchanged).
- `filters.js`: `updateFilterBadge()` (count of active non-default filters); small
  delegated handler for Sort/Order pills → set hidden value + active pill + `loadInvoices()`.
- Hide search + Filter button in **stats mode** (toggle a `main`/`body` class in
  `showStatsView/showInvoicesView`; hide via CSS).
- Optional: active-filter **chips** row (Design 2a) if low-cost, else defer.
- **Done when:** filtering works, panel opens on desktop without overflow, badge counts;
  CI green.

## Milestone 4 — Invoice list, rows & bulk toolbar

**Files:** `templates/partials/invoices.html`, `templates/partials/bulk-toolbar.html`,
`static/css/invoices.css`, `static/js/render.js`, `static/js/bulk.js`
**Commit:** `feat(ui): restyle invoice list, category badges, bulk toolbar`

- Restyle list card + rows + expanded detail (§2). Add a pure `categoryBadgeStyle(cat)`
  helper: fixed map for known categories → `--cat-*`, else hash → `--chart-1…8`; apply to
  the `.invoice-type` badge in `render.js`. (Export the helper for reuse by the combobox
  in Milestone 6.)
- Bulk toolbar: dark pill (§2). ✕ button takes over `data-action="deselect-all"`; remove
  the separate "Deselect All" text button. **Visibility bugfix (CSS):** add
  `visibility:hidden;opacity:0` to the resting state; `.visible` →
  `visibility:visible;opacity:1;bottom:24px`. (No JS change — gate already correct.)
- **Done when:** colored badges, expand/edit/delete work, toolbar hidden at 0 selected on
  tall viewports, ✕ deselects; CI green.

## Milestone 5 — Statistics restyle + chart palette

**Files:** `templates/partials/stats.html`, `static/css/stats.css`, `static/js/stats.js`,
`static/js/state.js`
**Commit:** `feat(ui): warm-sand statistics cards and chart palette`

- Tinted stat cards + chart-card restyle (§3).
- `state.js`: `chartColors` → `--chart-1…8` hex
  (`#c9a87c,#a8bfa0,#d9a48a,#b5a184,#c4b3d6,#d6bfa0,#a3c2c2,#e0cdb0`).
- `stats.js`: donut `cutout 65%→55%`; recolor Chart.js **tooltip** to light theme and bar
  grid/ticks to warm tones. Restyle the existing legend markup via CSS.
- **Done when:** cards/charts render in warm palette with light tooltips; CI green.

## Milestone 6 — Modals restyle + Category combobox (new component)

**Files:** `static/css/modals.css`, `templates/partials/modals/{add-invoice,import,bulk-edit,confirm-delete}.html`,
`static/js/combobox.js` (new), `static/css/combobox.css` (new), `templates/partials/filters.html`,
`static/js/api.js`, `static/js/dom.js`, `templates/index.html`
**Commit:** `feat(ui): restyle modals and add category combobox`

- `modals.css` + partials: overlay/modal/header/body/footer/input restyle (§4–6);
  import dropzone; confirm-delete icon tile + `--danger-solid` button.
- **Combobox** (Design 4b), reusable `createCombobox(root,{onChange})` with
  `setOptions/setValue/getValue`, live substring filter, keyboard nav, "Keine Kategorie"
  (empty) + "＋ Neue Kategorie '<x>' anlegen" entries, color dots via the shared
  `categoryBadgeStyle` helper. Each instance wraps a **hidden input keeping the original
  `data-el`** (`invoice-type`, `bulk-edit-category`, `type-filter`) so all `.value`
  readers are untouched.
- Replace the datalist population in `api.js`/`loadCategories` with feeding options into
  the combobox instances; filter combobox `onChange` → `loadInvoices()`. Remove the two
  `<datalist>`s and now-unused `populateDatalist` (`dom.js`). Add the `combobox.css`
  link to `index.html`.
- **Done when:** combobox filters/creates/selects and its value flows into add/edit,
  bulk-edit, and the filter query unchanged; CI green (no orphan helpers).

## Milestone 7 — PWA icons & favicon recolor

**Files:** `generate_icons.py`, `static/icons/*` (regenerated), `static/manifest.json`,
`static/favicon.svg`
**Commit:** `feat(pwa): recolor icons and favicon to terracotta`

- `generate_icons.py`: `GRADIENT_START (59,130,246)` → `(201,141,107)` (`#c98d6b`); set a
  warm `GRADIENT_END` (or END=START for a flat fill per handoff). Keep it ruff/mypy-clean.
- **Run** `uv run --group icons python generate_icons.py` (verify group name in
  `pyproject.toml`); commit regenerated PNGs.
- `manifest.json`: `theme_color`→`#c98d6b`, `background_color`→`#f3ede2`.
- `favicon.svg`: recolor gradient stops to terracotta.
- **Done when:** favicon/app icons terracotta; manifest colors updated; CI green.

## Milestone 8 — Service worker bump + final polish

**Files:** `static/sw.js` (+ any final lint/format sweeps)
**Commit:** `chore(pwa): bump cache; final redesign polish`

- Ensure `CACHE_NAME` is bumped past all prior changes (final `v33`+); do the end-to-end
  verification pass; fix any lint/format nits.
- **Done when:** full verification checklist passes; CI green.

---

## Critical files (by milestone)

| Milestone         | Files                                                                                                                                                                                                                   |
| ----------------- | ----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| 0 Plan            | `docs/redesign-warm-sand.md` (new)                                                                                                                                                                                      |
| 1 Tokens          | `static/css/variables.css`, `static/css/base.css`, `templates/index.html`                                                                                                                                               |
| 2 Sidebar         | `templates/partials/sidebar.html`←`header.html`, `static/css/sidebar.css`←`header.css`, `templates/index.html`, `static/js/stats.js`                                                                                    |
| 3 Filters         | `templates/partials/filters.html`, `static/css/filters.css`, `static/js/filters.js`, `static/js/stats.js`                                                                                                               |
| 4 Invoices/bulk   | `templates/partials/invoices.html`, `templates/partials/bulk-toolbar.html`, `static/css/invoices.css`, `static/js/render.js`, `static/js/bulk.js`                                                                       |
| 5 Stats           | `templates/partials/stats.html`, `static/css/stats.css`, `static/js/stats.js`, `static/js/state.js`                                                                                                                     |
| 6 Modals+combobox | `static/css/modals.css`, `templates/partials/modals/*`, `static/js/combobox.js`(new), `static/css/combobox.css`(new), `templates/partials/filters.html`, `static/js/api.js`, `static/js/dom.js`, `templates/index.html` |
| 7 Icons           | `generate_icons.py`, `static/icons/*`, `static/manifest.json`, `static/favicon.svg`                                                                                                                                     |
| 8 SW              | `static/sw.js`                                                                                                                                                                                                          |

## Reuse (don't reinvent)

- `dom.js`: `els()`, `escapeHtml`, `formatCurrency`, `formatDate`, `debounce`,
  `showToast`, `populateDropdown` (store filter keeps it).
- Existing delegation pattern + `setup*Listeners` — extend, no per-element listeners, no
  `window` globals.
- Existing collapsible mechanism (`grid-template-rows` + `.visible`) for the desktop panel.
- Existing bulk-toolbar `.visible` gate in `render.js` — keep; only harden CSS.
- Existing inline SVGs from partials — carry over (`stroke=currentColor`).

## Verification (end-to-end, run after Milestone 8; spot-check per milestone)

`uv run python -m summa` → `http://localhost:8000` (or `/run` skill).

1. Light warm-sand palette + Sora; terracotta theme-color/favicon.
2. Sidebar Invoices↔Statistics; active state; stats hides search/Filter.
3. Filters: pills + month nav; desktop panel opens (no overflow); Store/Category/From/To/
   Sort/Order apply + reload; badge counts; reset works.
4. Combobox: live filter; "Keine Kategorie" clears; "Neue Kategorie …" adds; value flows
   into add/edit, bulk-edit, and filter query (network params unchanged).
5. List: colored category badges; lazy item expand; edit/delete.
6. Bulk toolbar: hidden at 0 selected (no peeking on tall viewports); appears on select;
   Select all / Edit / Delete / ✕ work.
7. Stats: tinted cards, 55% donut + bars in warm palette, light tooltips, change badge.
8. Modals: add/edit, import (dropzone+paste), bulk-edit, confirm-delete.
9. ≤640px: bottom tab bar + New FAB + Import icon; collapsible filters; full-width bulk
   toolbar above tabs.
10. PWA: bumped `CACHE_NAME`; double reload → new assets cached (no stale dark theme);
    icons regenerated.
11. CI: `ruff check .` · `ruff format --check .` · `mypy` · `npm run lint` all green.
