"""Web interface routes."""

from pathlib import Path

from flask import Blueprint, Response, current_app, jsonify, render_template

web_bp: Blueprint = Blueprint("web", __name__)


@web_bp.route("/")
def index() -> str:
    """Render the main web interface."""
    return render_template("index.html")


@web_bp.route("/static/js-manifest.json")
def js_manifest() -> Response:
    """Return the URLs of all frontend JS modules for the service worker to precache."""
    static_root: str | None = current_app.static_folder
    assert static_root is not None  # always set by create_app()
    js_directory: Path = Path(static_root) / "js"
    urls: list[str] = sorted(
        f"/static/js/{path.name}" for path in js_directory.glob("*.js")
    )
    return jsonify(urls)
