"""Web interface routes."""

from pathlib import Path

from flask import Blueprint, Response, current_app, jsonify, render_template

web_bp: Blueprint = Blueprint("web", __name__)


@web_bp.route("/")
def index() -> str:
    """Render the main web interface."""
    return render_template("index.html")


def _asset_manifest(subdirectory: str, suffix: str) -> Response:
    """Return sorted URLs of static assets in ``subdirectory`` for the service worker to precache."""
    static_root: str | None = current_app.static_folder
    assert static_root is not None  # always set by create_app()
    asset_directory: Path = Path(static_root) / subdirectory
    urls: list[str] = sorted(
        f"/static/{subdirectory}/{path.name}"
        for path in asset_directory.glob(f"*{suffix}")
    )
    return jsonify(urls)


@web_bp.route("/static/js-manifest.json")
def js_manifest() -> Response:
    """Return the URLs of all frontend JS modules for the service worker to precache."""
    return _asset_manifest("js", ".js")


@web_bp.route("/static/css-manifest.json")
def css_manifest() -> Response:
    """Return the URLs of all frontend CSS files for the service worker to precache."""
    return _asset_manifest("css", ".css")


@web_bp.route("/static/fonts-manifest.json")
def fonts_manifest() -> Response:
    """Return the URLs of all self-hosted font files for the service worker to precache."""
    return _asset_manifest("fonts", ".woff2")
