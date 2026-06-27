"""Summa: invoice management and expense-tracking web application.

Exposes the application factory ``create_app()``. Importing this package has no
side effects; the eager WSGI ``app`` instance lives in :mod:`summa.wsgi`.
"""

import logging
from pathlib import Path

from flask import Flask
from flask_cors import CORS

from summa.db import init_db
from summa.routes.invoices import invoices_bp
from summa.routes.stats import stats_bp
from summa.routes.web import web_bp

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
)


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
