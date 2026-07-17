# Summa

Invoice management and expense tracking web application.

## Requirements

- Python 3.12+
- [uv](https://docs.astral.sh/uv/) (for local development)
- Docker (optional)

## Quick Start

### Docker (Recommended)

```bash
docker compose up -d
```

The application runs at `http://localhost:8000`. Data is persisted in a Docker volume.

See [`docker-compose.yml`](docker-compose.yml) for the full configuration including health checks and volume setup.

### Local Development

```bash
# Install dependencies (creates .venv automatically)
uv sync

# Run the application
uv run flask run --port 8000
```

The application runs at `http://localhost:8000` with the database stored in `invoices.db`.

### Linting, Formatting & Type Checking

[ruff](https://docs.astral.sh/ruff/) is used for formatting and linting, and
[mypy](https://mypy-lang.org/) for static type checking:

```bash
uv run ruff format .   # format
uv run ruff check .    # lint
uv run mypy            # type check (files are configured in pyproject.toml)
```

## Configuration

| Environment Variable | Default       | Description                                                                                                      |
| -------------------- | ------------- | ---------------------------------------------------------------------------------------------------------------- |
| `DATABASE_PATH`      | `invoices.db` | Path to SQLite database                                                                                          |
| `FLASK_DEBUG`        | `0` (off)     | Set to `1` to enable the Flask/Werkzeug debugger on the dev server. Never enable in production — it allows RCE.  |

## API

The application provides REST endpoints at `/api/invoices` for managing invoice data.
