# Plan: KI-gestützte Kategorie-Zuordnung für unkategorisierte Rechnungen

## Kontext

Aktuell haben viele Rechnungen keine Kategorie. Sie sollen per KI (Claude) anhand
von **Ladenname + gekauften Artikeln** eine Kategorie erhalten.

Die zwei ursprünglich erwogenen Optionen haben je einen durch den Code bestätigten Haken:

- **Export → KI → Reimport** scheitert am Import-Duplikat-Check
  (`summa/routes/invoices.py`, `import_invoices`): `date+store+total` matcht → korrigierte
  Kategorien würden als Duplikat verworfen. Ein Reimport kann bestehende Zeilen per Design
  nie _aktualisieren_. Zudem existiert gar kein Export-Endpoint.
- **KI schreibt direkt in die DB** umgeht den Review-Schritt und den bestehenden
  optimistischen Undo-Workflow (Commit #14).

**Gewählter Ansatz (dritte Option):** Ein Backend-Endpoint sammelt die unkategorisierten
Rechnungen des in der UI eingestellten Zeitraums, bündelt sie in **einem** Claude-Request
(strukturierte Ausgabe), und gibt **Vorschläge** ans Frontend zurück. Der Nutzer prüft/
bearbeitet die Vorschläge in einem Modal und bestätigt; das Schreiben läuft über den
**bereits vorhandenen** `bulk-update`-Pfad. So entfällt das Duplikat-Problem, die
menschliche Kontrolle bleibt, und es entsteht praktisch kein neuer Schreibpfad.

Entscheidungen:

- **Kategorien:** bestehende bevorzugen (`/api/categories` als Kontext), neue erlaubt
  und im Review als „neu" markiert.
- **Umfang:** nur unkategorisierte Rechnungen, gefiltert nach dem aktuell eingestellten
  Zeitraum/Filter.
- **Modell:** `claude-opus-4-8` über das offizielle `anthropic` Python-SDK.

## Backend

### 1. Dependency + Config

- `pyproject.toml`: `anthropic>=0.116,<1` zu `dependencies` hinzufügen, dann `uv sync`.
- API-Key über `ANTHROPIC_API_KEY` (Env-Var, analog zu `DATABASE_PATH`). Fehlt der Key,
  antwortet der Endpoint mit `error_response("AI categorization not configured", 503)` —
  kein Absturz, keine Hardcodes.
- `.env`/`docker-compose.yml`: Key als Env-Var durchreichen (nur dokumentieren, nicht committen).

### 2. Neues Modul `summa/ai.py`

Kapselt den Claude-Aufruf, damit die Route dünn bleibt und pure/testbar ist.

- Sphinx-Docstrings, vollständige Type-Annotations (auch lokale Variablen), Early-Return.
- Dataclass `CategorySuggestion { invoice_id: int, category: str | None, is_new: bool }`.
- Funktion `suggest_categories(invoices: list[dict[str, Any]], existing_categories: list[str]) -> list[CategorySuggestion]`:
  - Baut **einen** Request mit allen Rechnungen (id, store, items) als JSON im User-Turn.
  - Nutzt **strukturierte Ausgabe** via `client.messages.create(..., output_config={"format": {"type": "json_schema", "schema": …}})` — Schema = Array von `{invoice_id, category}`. `additionalProperties: false`, `required` gesetzt.
  - System-Prompt: „Ordne jeder Rechnung genau eine Ausgaben-Kategorie zu, basierend auf Ladenname und Artikeln. Bevorzuge eine Kategorie aus der übergebenen Liste; erfinde nur dann eine neue prägnante Kategorie, wenn keine passt." Die `existing_categories` werden im Prompt mitgegeben.
  - `model="claude-opus-4-8"`, `max_tokens` großzügig (z. B. 8000), `thinking={"type": "adaptive"}`.
  - `is_new` wird serverseitig berechnet: `category not in existing_categories` (case-insensitiv), damit das Flag verlässlich ist und nicht vom Modell abhängt.
  - Typische Exceptions gezielt fangen (`anthropic.APIError`, `anthropic.RateLimitError`) und als eigene Exception nach oben geben; die Route mappt auf `error_response`.

### 3. Neue Route in `summa/routes/invoices.py`

`POST /api/invoices/categorize-suggest` (read-only — schreibt **nichts**):

- Query-Params identisch zur Liste; **`_build_invoice_filter`** wiederverwenden und um
  `AND category IS NULL` ergänzen (nur Unkategorisierte). Kein neuer Filtercode nötig.
- Rechnungen laden **inkl. Items** (JOIN oder Nachladen wie in `get_invoice`).
  Sinnvolles Limit (z. B. max. 100 pro Durchlauf) gegen Token-/Kostenexplosion;
  Rest per Hinweis melden.
- Bestehende Kategorien über die vorhandene Query aus `get_categories` holen.
- `suggest_categories(...)` aufrufen, Ergebnis als JSON zurückgeben:
  `{"suggestions": [{invoice_id, store, category, is_new}, …], "count": n}`.
- `ApiResponse`-Rückgabetyp, `strip_text`-Konvention, gleiche try/except-Struktur wie die
  Nachbar-Handler. Der Endpoint schreibt nicht → kein `bulk-update` hier.

Kein `sw.js`/`CACHE_NAME`-Thema (reine API-Route).

## Frontend

Bewusst am bestehenden Muster orientiert (Vanilla-JS-Module unter `static/js/`,
`data-el`-Hooks, Combobox, optimistischer Undo-Toast).

### 4. Neues Modul `static/js/categorize.js`

- Button „Kategorien per KI vorschlagen" (im Header/Aktionsbereich), sichtbar wenn Filter
  aktiv. Ruft `POST /api/invoices/categorize-suggest?<buildFilterParams()>` auf —
  `buildFilterParams()` aus `static/js/api.js` exportieren und wiederverwenden.
- Antwort in einem **Review-Modal** rendern (analog Struktur der bestehenden Modals in
  `templates/index.html` + `static/css/modals.css`): pro Rechnung eine Zeile mit
  Store/Artikel-Kurzinfo und einer **editierbaren Category-Combobox** (bestehende
  `combobox.js`-Komponente, vorbefüllt mit dem Vorschlag). „Neu"-Vorschläge visuell markieren.
- Checkbox pro Zeile („übernehmen"), Bestätigen-Button.
- Beim Bestätigen: pro **eindeutiger Zielkategorie** die akzeptierten IDs sammeln und
  `PUT /api/invoices/bulk-update` aufrufen (ein Call je Kategorie), exakt wie
  `saveBulkEdit` in `static/js/bulk.js` — inklusive optimistischem Update + Undo-Toast,
  damit sich das Feature nahtlos einfügt.
- Danach `refreshAllData()` (`static/js/api.js`), damit neue Kategorien in den
  Comboboxen erscheinen.

### 5. Verdrahtung

- Modul in `app.js` importieren und Listener registrieren (wie `setupImportListeners`).
- Markup fürs Modal + Button in `templates/index.html`; `@html-eslint`-a11y beachten
  (semantische Elemente, `aria-*`/`id` nur wo Plattform es verlangt).
- CSS in bestehende `static/css/modals.css`/`components.css` einfügen (recess-order,
  kein `!important`, keine id-Selektoren).
- **`CACHE_NAME` in `static/sw.js` bumpen** (neue JS/CSS-Assets ändern Cache-Inhalt).
  Neue Dateien werden per `js-manifest`/`css-manifest` automatisch entdeckt — kein
  weiterer `sw.js`-Edit nötig.

## Tests

- `tests/`: Unit-Test für `summa/ai.py` mit **gemocktem** Anthropic-Client (kein echter
  API-Call) — prüft Prompt-Zusammenbau, `is_new`-Berechnung und Fehler-Mapping.
- Test für die neue Route: Filter (`category IS NULL` + Zeitraum), 503 ohne Key,
  Antwort-Shape. Anthropic-Aufruf in der Route mocken.

## Verifikation

1. `uv sync` (holt `anthropic`).
2. Lint/Types: `uv run ruff format .`, `uv run ruff check .`, `uv run mypy`, `npm run lint` (CI-Gate).
3. `uv run pytest` — neue Tests grün.
4. Runtime (verify-Skill): Dev-Server starten, `ANTHROPIC_API_KEY` gesetzt,
   ein paar unkategorisierte Test-Rechnungen anlegen, Zeitraum-Filter setzen,
   „Vorschlagen" klicken → Review-Modal prüfen (Vorschläge sinnvoll, „neu"-Markierung),
   bestätigen → Kategorien landen via `bulk-update`, Undo funktioniert, Filter/Stats
   aktualisieren sich. Playwright-Screenshot des Review-Modals.
5. Negativpfad: Ohne `ANTHROPIC_API_KEY` → 503 + saubere Fehlermeldung, kein Crash.
