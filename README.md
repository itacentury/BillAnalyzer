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

The application runs at `http://localhost:8000`. Data is persisted in the
bind-mounted `./data` directory next to the compose file.

See [`docker-compose.yml`](docker-compose.yml) for the full configuration.

### Production Deployment

The app (both the dev server and gunicorn) serves **plain HTTP only** and performs
no TLS termination. For any deployment beyond `localhost`, run it **behind a
TLS-terminating reverse proxy** so traffic — which includes all of your financial
data — is encrypted in transit. Do not publish port `8000` directly to an untrusted
network without such a proxy in front.

Example nginx server block terminating TLS and forwarding to the container:

```nginx
server {
    listen 443 ssl;
    server_name summa.example.com;

    ssl_certificate     /etc/letsencrypt/live/summa.example.com/fullchain.pem;
    ssl_certificate_key /etc/letsencrypt/live/summa.example.com/privkey.pem;

    location / {
        proxy_pass http://127.0.0.1:8000;
        proxy_set_header Host              $host;
        proxy_set_header X-Forwarded-For   $proxy_add_x_forwarded_for;
        proxy_set_header X-Forwarded-Proto $scheme;
    }
}
```

Certificates come from certbot/ACME (Let's Encrypt) or your own CA. Bind the
container's published port to `127.0.0.1:8000:8000` so only the proxy can reach it.

Cross-origin browser access is **denied by default** — the PWA is served
same-origin and needs no CORS. To allow other origins (e.g. a native mobile
client), set `CORS_ALLOWED_ORIGINS` (see [Configuration](#configuration)).

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

## CI & Supply-Chain Security

The [`docker` workflow](.github/workflows/docker.yml) builds the image and runs two supply-chain checks against it before publishing:

- **Vulnerability scan & gate** — a single [`grype`](https://github.com/anchore/grype) run scans the image and emits three outputs from one pass: a full-inventory **table** in the job log (every finding with its _fixed in_ column — the place to look when a critical blocks the push), a **JSON** report, and a **SARIF** report. A following step **fails the build on any _fixable_ critical CVE**, deciding from the JSON's machine-readable `fix.state` field (grype's SARIF has no such field). Gating only on findings that have a released fix keeps the pipeline actionable: a critical CVE with no upstream fix (common for base-image OS packages Debian marks _won't fix_) can't be remediated by a rebuild, so it must not block the push indefinitely — the gate re-activates automatically once a fix ships.
- **Scan visibility** — the full SARIF report is uploaded via `github/codeql-action/upload-sarif`, so findings appear under **Security → Code scanning** and as annotations on pull requests. Because the SARIF is the full inventory, the Security tab lists unfixable criticals too; only the push gate is limited to fixable ones.
- **SBOM** — the published image carries an SPDX SBOM attached as an OCI attestation (`sbom: true` on the build-and-push step), so the bill of materials ships with the image rather than as a throwaway job artifact. Inspect it with `docker buildx imagetools inspect <user>/summa:latest --format '{{ json .SBOM }}'`.

Two operational caveats:

- **Pull requests from forks cannot upload SARIF.** GitHub grants fork PRs only a read-only token, so `upload-sarif` (which needs `security-events: write`) fails on them. The scan and its critical-CVE gate still run — only the Security-tab upload is skipped. Same-repo PRs are unaffected.
- **Code scanning must be enabled** for the repository, otherwise the SARIF upload succeeds but no alerts are surfaced. It is free on public repositories; private repositories require GitHub Advanced Security.
- **Pull-request findings are filtered by branch.** The **Security → Code scanning** view defaults to the default branch (`main`); a scan that ran on a PR only appears after switching the **Branch** filter to that PR's branch (or via the PR's own file annotations).

## AI Category Suggestions

Summa can suggest a spending category for uncategorized invoices using Claude.
From the categorize dialog you trigger a run over the currently filtered
uncategorized invoices; the model returns one category per invoice and you review
and confirm the suggestions before anything is written — the request itself never
mutates your data.

**Enabling it:** set `ANTHROPIC_API_KEY` in the server environment (get a key from
the [Anthropic Console](https://console.anthropic.com/)). Locally, copy
[`.env.example`](.env.example) to `.env` and fill it in; for Docker the
`env_file` in [`docker-compose.yml`](docker-compose.yml) picks the same `.env` up.
If `ENABLE_AI_SUGGESTIONS` is missing, the feature is disabled by default. Set
`ENABLE_AI_SUGGESTIONS=1` to enable it; set `0` to hard-disable it and hide the
trigger entirely. With the master switch enabled but no API key configured, the
trigger still renders and the endpoint returns `503`.

The model — Claude Haiku, Sonnet, or Opus — is chosen in the UI (default: Haiku)
and remembered per browser, so it is **not** an environment variable.

## Configuration

Copy [`.env.example`](.env.example) to `.env` and fill in the values you need.

| Environment Variable    | Default       | Description                                                                                                                   |
| ----------------------- | ------------- | ----------------------------------------------------------------------------------------------------------------------------- |
| `ANTHROPIC_API_KEY`     | _(unset)_     | Configures [AI category suggestions](#ai-category-suggestions). Required for requests once the master switch is enabled.      |
| `ENABLE_AI_SUGGESTIONS` | _(unset)_     | Master switch for AI category suggestions. Unset/`0` disables the feature; set to `1` to show the trigger and allow requests. |
| `DATABASE_PATH`         | `invoices.db` | Path to SQLite database                                                                                                       |
| `CORS_ALLOWED_ORIGINS`  | _(empty)_     | Cross-origin allowlist — a comma-separated list of origins, or `*` for the wildcard. Empty = same-origin only.                |
| `FLASK_DEBUG`           | `0` (off)     | Set to `1` to enable the Flask/Werkzeug debugger on the dev server. Never enable in production — it allows RCE.               |

## API

The application provides REST endpoints at `/api/invoices` for managing invoice data.
