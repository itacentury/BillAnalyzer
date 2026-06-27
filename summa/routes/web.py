"""Web interface routes."""

from flask import Blueprint, render_template

web_bp: Blueprint = Blueprint("web", __name__)


@web_bp.route("/")
def index() -> str:
    """Render the main web interface."""
    return render_template("index.html")
