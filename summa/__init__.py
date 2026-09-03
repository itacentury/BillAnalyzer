"""Summa: invoice management and expense-tracking web application.

Exposes the application factory ``create_app()``. Importing this package has no
side effects; the eager WSGI ``app`` instance lives in :mod:`summa.wsgi`.
"""

import logging
import os
import secrets
from datetime import timedelta
from pathlib import Path
from typing import Final

from flask import Flask, Response
from flask_cors import CORS
from werkzeug.exceptions import RequestEntityTooLarge

from summa import config
from summa.auth import auth_config_warnings
from summa.db import init_db
from summa.helpers import ApiResponse, error_response
from summa.routes.auth import auth_bp
from summa.routes.invoices import invoices_bp
from summa.routes.stats import stats_bp
from summa.routes.web import web_bp

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
)

logger: logging.Logger = logging.getLogger(__name__)

# Content-Security-Policy. Chart.js and the web fonts are self-hosted (see
# SECURITY-TODO M2), so every fetch directive is locked to 'self' — no external
# origins are permitted.
_CSP_DIRECTIVES: list[str] = [
    "default-src 'self'",
    "script-src 'self'",
    "style-src 'self'",
    "font-src 'self'",
    "img-src 'self' data:",
    "connect-src 'self'",
    "frame-ancestors 'none'",
    "base-uri 'none'",
    "form-action 'self'",
    "object-src 'none'",
]
SECURITY_CSP: str = "; ".join(_CSP_DIRECTIVES)

# Cap the raw request body (SECURITY-TODO M3): every handler reads the whole
# JSON body into memory, so without this one request could exhaust it.
MAX_CONTENT_LENGTH: Final[int] = 5 * 1024 * 1024  # 5 MB


def _cors_origins() -> str | list[str]:
    """Resolve the CORS allowlist from CORS_ALLOWED_ORIGINS.

    Empty/unset -> no cross-origin access (the PWA is same-origin). The literal
    '*' is an explicit opt-in for native mobile clients; otherwise a
    comma-separated origin allowlist.
    """
    raw: str = os.environ.get("CORS_ALLOWED_ORIGINS", "").strip()
    if not raw:
        return []
    if raw == "*":
        return "*"
    return [origin.strip() for origin in raw.split(",") if origin.strip()]


def _configure_sessions(app: Flask) -> None:
    """Set the signing key and cookie attributes for the login session."""
    # Without a configured secret the key is per-process and per-restart, so
    # sessions silently stop working under `gunicorn --workers 2` (each worker
    # signs differently). Generating one anyway keeps cookies unforgeable; the
    # startup warning below is what tells the operator to fix the setup.
    app.secret_key = config.session_secret() or secrets.token_urlsafe(32)
    app.config["SESSION_COOKIE_NAME"] = "summa_session"
    app.config["SESSION_COOKIE_HTTPONLY"] = True
    app.config["SESSION_COOKIE_SECURE"] = config.cookie_secure()
    app.config["SESSION_COOKIE_SAMESITE"] = config.cookie_samesite()
    app.config["PERMANENT_SESSION_LIFETIME"] = timedelta(days=config.session_days())

    if not config.auth_enabled():
        return
    logger.info("Login gate enabled")
    for warning in auth_config_warnings():
        logger.warning("Auth configuration: %s", warning)


def create_app() -> Flask:
    """Build, configure and return the Flask application."""
    # Templates and static assets live at the repo root, not inside the package
    root: Path = Path(__file__).resolve().parent.parent
    app: Flask = Flask(
        __name__,
        template_folder=str(root / "templates"),
        static_folder=str(root / "static"),
    )
    # Cross-origin access is denied by default (the PWA is same-origin, so it
    # needs no CORS). Opt origins in via CORS_ALLOWED_ORIGINS; '*' re-enables the
    # wildcard for native mobile clients (SECURITY-TODO H1).
    CORS(app, origins=_cors_origins())
    app.config["MAX_CONTENT_LENGTH"] = MAX_CONTENT_LENGTH
    _configure_sessions(app)

    @app.after_request
    def set_security_headers(response: Response) -> Response:
        """Attach the CSP and hardening headers to every response."""
        response.headers["Content-Security-Policy"] = SECURITY_CSP
        response.headers["X-Content-Type-Options"] = "nosniff"
        response.headers["X-Frame-Options"] = "DENY"
        response.headers["Referrer-Policy"] = "no-referrer"
        return response

    @app.errorhandler(RequestEntityTooLarge)
    def handle_request_too_large(_: RequestEntityTooLarge) -> ApiResponse:
        """Return the body-size 413 as JSON, matching the API error convention."""
        return error_response("Request body too large", 413)

    app.register_blueprint(web_bp)
    app.register_blueprint(auth_bp)
    app.register_blueprint(invoices_bp)
    app.register_blueprint(stats_bp)

    # Initialize database on app creation (works with gunicorn and dev server)
    init_db()

    return app


def main() -> None:
    """Start the Flask development server."""
    app: Flask = create_app()
    # Debug is opt-in only: the Werkzeug debugger allows remote code execution,
    # so it must never default on (SECURITY-TODO M6).
    debug: bool = os.environ.get("FLASK_DEBUG") == "1"
    app.run(debug=debug, port=8000)
