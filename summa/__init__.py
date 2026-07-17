"""Summa: invoice management and expense-tracking web application.

Exposes the application factory ``create_app()``. Importing this package has no
side effects; the eager WSGI ``app`` instance lives in :mod:`summa.wsgi`.
"""

import logging
from pathlib import Path

from flask import Flask, Response
from flask_cors import CORS

from summa.db import init_db
from summa.routes.invoices import invoices_bp
from summa.routes.stats import stats_bp
from summa.routes.web import web_bp

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
)

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


def create_app() -> Flask:
    """Build, configure and return the Flask application."""
    # Templates and static assets live at the repo root, not inside the package
    root: Path = Path(__file__).resolve().parent.parent
    app: Flask = Flask(
        __name__,
        template_folder=str(root / "templates"),
        static_folder=str(root / "static"),
    )
    CORS(app)  # Enable CORS for all routes (required for native mobile apps)

    @app.after_request
    def set_security_headers(response: Response) -> Response:
        """Attach the CSP and hardening headers to every response."""
        response.headers["Content-Security-Policy"] = SECURITY_CSP
        response.headers["X-Content-Type-Options"] = "nosniff"
        response.headers["X-Frame-Options"] = "DENY"
        response.headers["Referrer-Policy"] = "no-referrer"
        return response

    app.register_blueprint(web_bp)
    app.register_blueprint(invoices_bp)
    app.register_blueprint(stats_bp)

    # Initialize database on app creation (works with gunicorn and dev server)
    init_db()

    return app


def main() -> None:
    """Start the Flask development server."""
    app: Flask = create_app()
    app.run(debug=True, port=8000)
