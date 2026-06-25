# Refactor-Plan: `app.py` in ein `summa/`-Package mit App-Factory aufteilen

## Kontext

Das gesamte Backend liegt in einer einzigen `app.py` (591 Zeilen): App-Setup,
Logging, DB-Verbindung + Schema/Migrationen, Text-Helper und **alle** REST-Routen
(Invoices-CRUD, import, bulk, stores, categories, stats) plus die Web-View. Das
erschwert Lesbarkeit, Erweiterbarkeit und Navigation. Ziel ist eine Aufteilung
entlang der vorhandenen, klaren Verantwortungsgrenzen in ein idiomatisches
Python-Package `summa/` mit einer `create_app()`-Factory und Flask-Blueprints —
ohne Verhalten oder API-URLs zu ändern.

Gewählter Ansatz: **`summa/`-Package mit App-Factory, direkt im Repo-Root**
(neben `templates/`, `static/`, `pyproject.toml`). Das ändert das gunicorn/Flask-
Target und erfordert Anpassungen an Dockerfile, mypy-Config und den Template-/
Static-Pfaden (siehe unten).

## Zielstruktur

```
summa/
  __init__.py     # logging.basicConfig, create_app(), app = create_app(), main()
  __main__.py     # python -m summa -> main()  (Dev-Server-Entry)
  db.py           # DATABASE, get_db(), init_db()
  helpers.py      # ApiResponse-Typalias, strip_text()
  routes/
    __init__.py   # leer (Paket-Marker)
    web.py        # web_bp:      /            (index)
    invoices.py   # invoices_bp: /api/invoices*, /api/stores, /api/categories
    stats.py      # stats_bp:    /api/stats   + _calculate_comparison
templates/  static/   # bleiben am Repo-Root (unverändert)
```

Die **Route-Strings bleiben exakt identisch** (nur `@app.route` → `@<bp>.route`),
damit die API unverändert ist ("unrelated code nicht anfassen").

## Code-Mapping (aus `app.py`)

- **`summa/db.py`** — `DATABASE` (Zeile 24), `get_db()` (30–36), `init_db()`
  (47–99). Eigener `logger = logging.getLogger(__name__)`.
- **`summa/helpers.py`** — `ApiResponse`-Alias (27), `strip_text()` (39–44).
- **`summa/routes/web.py`** — `index()` (102–105) auf `web_bp`.
- **`summa/routes/invoices.py`** — `get_invoices` (108), `get_stores` (182),
  `get_categories` (195), `add_invoice` (209), `import_invoices` (246),
  `update_invoice` (310), `delete_invoice` (358), `bulk_update_invoices` (374),
  `bulk_delete_invoices` (426) auf `invoices_bp`. Importiert `get_db` aus
  `summa.db`, `strip_text`/`ApiResponse` aus `summa.helpers`. Eigener `logger`.
- **`summa/routes/stats.py`** — `_calculate_comparison` (461–499) + `get_stats`
  (502–578) auf `stats_bp`. Importiert `get_db`, `logger`.
- **`summa/__init__.py`** — `logging.basicConfig(...)` (16–19), `create_app()`:
  erzeugt `Flask(__name__, template_folder=..., static_folder=...)`, aktiviert
  CORS, registriert die drei Blueprints, ruft `init_db()`, gibt die App zurück.
  Modul-Level: `app = create_app()`. `main()` (581–583) ruft `app.run(debug=True,
  port=8000)`.

### Template-/Static-Pfade (wichtig bei Package-Layout)

Da die App-Instanz jetzt **innerhalb** von `summa/` erzeugt wird, zeigt Flasks
Default auf `summa/templates` bzw. `summa/static`. Die Ordner bleiben aber am
Repo-Root. Daher in `create_app()` explizit setzen (mit `pathlib`, gemäß
Code-Style):

```python
root: Path = Path(__file__).resolve().parent.parent
app = Flask(
    __name__,
    template_folder=str(root / "templates"),
    static_folder=str(root / "static"),
)
```

`static_url_path` bleibt Default `/static` → `url_for('static', ...)`, `sw.js`
und die Cache-Pfade bleiben unverändert.

## Deployment-/Config-Anpassungen (zwingend)

- **`Dockerfile`**
  - Zeile 52: `COPY app.py .` → `COPY summa/ summa/`
  - Zeile 67: `ENV FLASK_APP=app.py` → `ENV FLASK_APP=summa`
  - Zeile 81 (CMD): `"app:app"` → `"summa:app"`
- **`pyproject.toml`** — `[tool.mypy] files = ["app.py", "generate_icons.py"]`
  → `files = ["summa", "generate_icons.py"]`.
- **`CLAUDE.md`** — Hinweis im Architecture-Abschnitt aktualisieren: Backend ist
  jetzt das `summa/`-Package (Factory in `summa/__init__.py`, Blueprints in
  `summa/routes/`, DB-Layer in `summa/db.py`); `init_db()` läuft in `create_app()`.
  Dev-Command `uv run flask run` funktioniert via `FLASK_APP=summa` weiter.
- **`app.py` am Root löschen** (Inhalt vollständig migriert).

`ruff check .` / `ruff format` scannen das ganze Repo → keine Config-Änderung
nötig. `ci.yml` braucht keine Änderung (nutzt die obigen Configs).

## Verifikation (End-to-End)

1. `uv run ruff format .` und `uv run ruff check .` → sauber.
2. `uv run mypy` → strict-clean über das `summa`-Package.
3. Dev-Server: `FLASK_APP=summa uv run flask run --port 8000` (bzw.
   `uv run python -m summa`). Prüfen:
   - `GET /` rendert die Seite (Template-Pfad korrekt), `/static/...` lädt.
   - `GET /api/invoices`, `/api/stores`, `/api/categories`, `/api/stats`
     liefern wie zuvor.
   - `POST /api/invoices` anlegen, `PUT`/`DELETE`, `bulk-update`/`bulk-delete`,
     `import` durchspielen; DB `./invoices.db` wird via `init_db()` erzeugt.
4. `npm run lint` (Frontend unverändert, sollte unberührt grün bleiben).
5. Optional Container-Smoke: `docker compose up -d`, dann `curl localhost:8000/`
   und ein `GET /api/stats` → bestätigt `summa:app`-Target + COPY/ENV-Änderungen.
