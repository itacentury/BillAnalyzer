"""Tests for the deny-by-default request gate."""

import pytest
from flask.testing import FlaskClient


def test_protected_route_is_rejected_without_a_session(
    gated_client: FlaskClient,
) -> None:
    """An API route answers 401 JSON when the client has no session."""
    response = gated_client.get("/api/invoices")

    assert response.status_code == 401
    assert response.get_json() == {
        "success": False,
        "error": "Authentication required",
    }


def test_protected_route_is_reachable_with_a_session(
    authed_client: FlaskClient,
) -> None:
    """The same route works once the client is logged in."""
    assert authed_client.get("/api/invoices").status_code == 200


def test_writes_are_gated_too(gated_client: FlaskClient) -> None:
    """The gate is not limited to reads."""
    response = gated_client.delete("/api/invoices/1")

    assert response.status_code == 401


def test_unknown_api_route_is_gated_before_routing(gated_client: FlaskClient) -> None:
    """A path that does not exist is rejected as 401, not disclosed as 404."""
    assert gated_client.get("/api/does-not-exist").status_code == 401


@pytest.mark.parametrize(
    "path",
    [
        "/",
        "/static/js-manifest.json",
        "/static/css-manifest.json",
        "/static/fonts-manifest.json",
        "/static/sw.js",
        "/static/manifest.json",
        "/static/icons/icon-192.png",
    ],
)
def test_public_paths_are_never_gated(gated_client: FlaskClient, path: str) -> None:
    """Everything needed to render the login screen stays reachable.

    Asserted as "not 401" rather than "200" so the test does not depend on the
    file existing in the checkout.
    """
    assert gated_client.get(path).status_code != 401


def test_login_endpoint_is_reachable_without_a_session(
    gated_client: FlaskClient,
) -> None:
    """The gate cannot lock the client out of the way to obtain a session."""
    response = gated_client.post("/api/auth/login", json={"password": "wrong"})

    assert response.status_code == 401
    assert response.get_json()["error"] == "Invalid password"


def test_security_headers_are_attached_to_a_rejection(
    gated_client: FlaskClient,
) -> None:
    """A 401 still passes through the after_request hardening."""
    response = gated_client.get("/api/invoices")

    assert response.headers["X-Content-Type-Options"] == "nosniff"
    assert "Content-Security-Policy" in response.headers


@pytest.mark.parametrize("path", ["/", "/api/invoices", "/api/stats"])
def test_nothing_is_gated_while_the_feature_is_off(
    client: FlaskClient, path: str
) -> None:
    """With the gate disabled the app behaves exactly as it did before."""
    assert client.get(path).status_code == 200
