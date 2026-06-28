"""Tests for the web interface routes and the service-worker JS manifest."""

from pathlib import Path

from flask.testing import FlaskClient


def test_js_manifest_matches_static_js_directory(client: FlaskClient) -> None:
    """The manifest lists exactly the JS files present under static/js/."""
    response = client.get("/static/js-manifest.json")
    assert response.status_code == 200

    manifest: list[str] = response.get_json()
    assert isinstance(manifest, list)

    js_directory: Path = Path(__file__).resolve().parent.parent / "static" / "js"
    expected: set[str] = {
        f"/static/js/{path.name}" for path in js_directory.glob("*.js")
    }
    assert set(manifest) == expected
