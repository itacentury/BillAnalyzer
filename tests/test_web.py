"""Tests for the web interface routes and the service-worker asset manifests."""

import re
from pathlib import Path

from flask.testing import FlaskClient

from summa.routes.invoices import DEFAULT_PAGE_SIZE


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


def test_css_manifest_matches_static_css_directory(client: FlaskClient) -> None:
    """The manifest lists exactly the CSS files present under static/css/."""
    response = client.get("/static/css-manifest.json")
    assert response.status_code == 200

    manifest: list[str] = response.get_json()
    assert isinstance(manifest, list)

    css_directory: Path = Path(__file__).resolve().parent.parent / "static" / "css"
    expected: set[str] = {
        f"/static/css/{path.name}" for path in css_directory.glob("*.css")
    }
    assert set(manifest) == expected


def test_frontend_default_page_size_matches_backend() -> None:
    """The UI's initial pageSize must equal the backend's DEFAULT_PAGE_SIZE."""
    state_js: Path = (
        Path(__file__).resolve().parent.parent / "static" / "js" / "state.js"
    )
    # Anchored + case-sensitive so it matches `pageSize:` (line 12), never the
    # `effectivePageSize:` line below it.
    match: re.Match[str] | None = re.search(
        r"^\s*pageSize:\s*(\d+)", state_js.read_text(), re.MULTILINE
    )
    assert match is not None, "pageSize default not found in state.js"
    frontend_default: int = int(match.group(1))
    assert frontend_default == DEFAULT_PAGE_SIZE


def test_security_headers_present_on_every_response(client: FlaskClient) -> None:
    """The after_request hook attaches CSP and hardening headers globally."""
    for path in ("/", "/api/invoices"):
        response = client.get(path)
        csp: str | None = response.headers.get("Content-Security-Policy")
        assert csp is not None
        assert "default-src 'self'" in csp
        assert "connect-src 'self'" in csp
        assert response.headers["X-Content-Type-Options"] == "nosniff"
        assert response.headers["X-Frame-Options"] == "DENY"
        assert response.headers["Referrer-Policy"] == "no-referrer"
