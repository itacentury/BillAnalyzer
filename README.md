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
client), set `CORS_ALLOWED_ORIGINS` to a comma-separated allowlist, or to `*` to
re-enable the wildcard.

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

- **Vulnerability gate** — [`anchore/scan-action`](https://github.com/anchore/scan-action) scans the image and **fails the build on any critical CVE** (`fail-build: true`, `severity-cutoff: critical`), so a critical finding blocks the push.
- **Scan visibility** — the scan's SARIF report is uploaded via `github/codeql-action/upload-sarif`, so findings appear under **Security → Code scanning** and as annotations on pull requests.
- **SBOM** — the published image carries an SPDX SBOM attached as an OCI attestation (`sbom: true` on the build-and-push step), so the bill of materials ships with the image rather than as a throwaway job artifact. Inspect it with `docker buildx imagetools inspect <user>/summa:latest --format '{{ json .SBOM }}'`.

Two operational caveats:

- **Pull requests from forks cannot upload SARIF.** GitHub grants fork PRs only a read-only token, so `upload-sarif` (which needs `security-events: write`) fails on them. The scan and its critical-CVE gate still run — only the Security-tab upload is skipped. Same-repo PRs are unaffected.
- **Code scanning must be enabled** for the repository, otherwise the SARIF upload succeeds but no alerts are surfaced. It is free on public repositories; private repositories require GitHub Advanced Security.

## Configuration

| Environment Variable | Default       | Description                                                                                                     |
| -------------------- | ------------- | --------------------------------------------------------------------------------------------------------------- |
| `DATABASE_PATH`      | `invoices.db` | Path to SQLite database                                                                                         |
| `FLASK_DEBUG`        | `0` (off)     | Set to `1` to enable the Flask/Werkzeug debugger on the dev server. Never enable in production — it allows RCE. |

## API

The application provides REST endpoints at `/api/invoices` for managing invoice data.
