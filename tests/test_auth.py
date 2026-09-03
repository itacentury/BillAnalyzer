"""Tests for the login, logout and session-status endpoints."""

import logging

import pytest
from flask.testing import FlaskClient
from werkzeug.test import TestResponse

from summa import auth, config, ratelimit
from tests.conftest import TEST_PASSWORD, TEST_PASSWORD_HASH

# Every shape a hash can arrive in that Werkzeug cannot read. The first four are
# what Docker Compose's '$' interpolation leaves behind; only the last two reach
# check_password_hash's ValueError, the rest come back as a plain False.
UNREADABLE_HASHES: list[str] = [
    "not-a-real-hash",
    "scrypt:32768:8:1$$",
    "scrypt:32768:8:1$$abcd",
    "$salt$abcd",
    "notamethod$salt$abcd",
    "scrypt:32768:8$salt$abcd",
]


def _session_cookie(response: TestResponse) -> str:
    """Return the Set-Cookie header that carries the session cookie."""
    headers: list[str] = response.headers.getlist("Set-Cookie")
    matching: list[str] = [value for value in headers if "summa_session=" in value]
    assert matching, f"no session cookie in {headers}"
    return matching[0]


def test_login_with_correct_password_starts_a_session(
    gated_client: FlaskClient,
) -> None:
    """The right password answers 200 and sets the session cookie."""
    response = gated_client.post("/api/auth/login", json={"password": TEST_PASSWORD})

    assert response.status_code == 200
    assert response.get_json() == {"success": True, "authed": True}
    assert "HttpOnly" in _session_cookie(response)


def test_login_with_wrong_password_is_rejected(gated_client: FlaskClient) -> None:
    """A wrong password answers 401 and sets no cookie."""
    response = gated_client.post("/api/auth/login", json={"password": "wrong"})

    assert response.status_code == 401
    assert response.get_json() == {"success": False, "error": "Invalid password"}
    assert not response.headers.getlist("Set-Cookie")


def test_login_requires_a_string_password(gated_client: FlaskClient) -> None:
    """A non-string password is a client error, not an auth failure."""
    response = gated_client.post("/api/auth/login", json={"password": 1234})

    assert response.status_code == 400


def test_login_requires_a_json_object(gated_client: FlaskClient) -> None:
    """A non-object body is rejected before the password is looked at."""
    response = gated_client.post("/api/auth/login", json=["not", "an", "object"])

    assert response.status_code == 400


def test_login_without_a_configured_hash_always_fails(
    gated_client: FlaskClient, monkeypatch: pytest.MonkeyPatch
) -> None:
    """A gate without a password hash fails closed rather than open."""
    monkeypatch.setenv(config.PASSWORD_HASH_ENV, "")

    response = gated_client.post("/api/auth/login", json={"password": TEST_PASSWORD})

    assert response.status_code == 401


def test_login_with_a_malformed_hash_fails_closed(
    gated_client: FlaskClient, monkeypatch: pytest.MonkeyPatch
) -> None:
    """A mangled hash yields 401, never a 500."""
    monkeypatch.setenv(config.PASSWORD_HASH_ENV, "not-a-real-hash")

    response = gated_client.post("/api/auth/login", json={"password": TEST_PASSWORD})

    assert response.status_code == 401


@pytest.mark.parametrize("configured_hash", UNREADABLE_HASHES)
def test_login_with_an_unreadable_hash_never_500s(
    gated_client: FlaskClient, monkeypatch: pytest.MonkeyPatch, configured_hash: str
) -> None:
    """Every unreadable hash shape is a 401 — including the ones that raise."""
    monkeypatch.setenv(config.PASSWORD_HASH_ENV, configured_hash)

    response = gated_client.post("/api/auth/login", json={"password": TEST_PASSWORD})

    assert response.status_code == 401


def test_login_with_an_unknown_hash_method_is_logged(
    gated_client: FlaskClient,
    monkeypatch: pytest.MonkeyPatch,
    caplog: pytest.LogCaptureFixture,
) -> None:
    """The shapes Werkzeug raises on say so in the log rather than 500-ing."""
    monkeypatch.setenv(config.PASSWORD_HASH_ENV, "notamethod$salt$abcd")

    with caplog.at_level(logging.WARNING, logger="summa.auth"):
        response = gated_client.post(
            "/api/auth/login", json={"password": TEST_PASSWORD}
        )

    assert response.status_code == 401
    assert "is not a valid hash" in caplog.text


def test_remember_makes_the_cookie_persistent(gated_client: FlaskClient) -> None:
    """With `remember` the cookie carries an expiry."""
    response = gated_client.post(
        "/api/auth/login", json={"password": TEST_PASSWORD, "remember": True}
    )

    assert "Expires=" in _session_cookie(response)


def test_a_remembered_session_is_not_renewed_by_use(gated_client: FlaskClient) -> None:
    """The expiry is fixed at login, so an active session still hard-expires."""
    gated_client.post(
        "/api/auth/login", json={"password": TEST_PASSWORD, "remember": True}
    )

    response = gated_client.get("/api/invoices")

    assert response.status_code == 200
    assert not response.headers.getlist("Set-Cookie")


def test_without_remember_the_cookie_is_session_scoped(
    gated_client: FlaskClient,
) -> None:
    """Without `remember` the browser drops the cookie on close."""
    response = gated_client.post(
        "/api/auth/login", json={"password": TEST_PASSWORD, "remember": False}
    )

    assert "Expires=" not in _session_cookie(response)


def test_logout_clears_the_session(authed_client: FlaskClient) -> None:
    """After logout the status endpoint reports an unauthenticated client."""
    response = authed_client.post("/api/auth/logout")

    assert response.status_code == 200
    assert response.get_json() == {"success": True, "authed": False}
    assert authed_client.get("/api/auth/me").get_json() == {
        "authed": False,
        "enabled": True,
    }


def test_logout_succeeds_without_a_session(gated_client: FlaskClient) -> None:
    """Logging out when not logged in is a no-op, not an error."""
    assert gated_client.post("/api/auth/logout").status_code == 200


def test_me_reports_an_authenticated_client(authed_client: FlaskClient) -> None:
    """A logged-in client is reported as authed with the gate enabled."""
    response = authed_client.get("/api/auth/me")

    assert response.status_code == 200
    assert response.get_json() == {"authed": True, "enabled": True}


def test_me_reports_a_missing_session_with_200(gated_client: FlaskClient) -> None:
    """The status endpoint answers 200 even when the client must log in."""
    response = gated_client.get("/api/auth/me")

    assert response.status_code == 200
    assert response.get_json() == {"authed": False, "enabled": True}


def test_me_reports_everyone_as_authed_when_the_gate_is_off(
    client: FlaskClient,
) -> None:
    """With the gate disabled the client is authed and the logout UI stays hidden."""
    response = client.get("/api/auth/me")

    assert response.status_code == 200
    assert response.get_json() == {"authed": True, "enabled": False}


def test_repeated_wrong_passwords_are_throttled(gated_client: FlaskClient) -> None:
    """Guessing is cut off once the failure window is full."""
    for _ in range(ratelimit.MAX_FAILURES):
        assert (
            gated_client.post("/api/auth/login", json={"password": "wrong"}).status_code
            == 401
        )

    response = gated_client.post("/api/auth/login", json={"password": "wrong"})

    assert response.status_code == 429
    assert response.get_json()["error"] == "Too many login attempts"
    assert int(response.headers["Retry-After"]) > 0


def test_throttling_outlasts_the_correct_password(gated_client: FlaskClient) -> None:
    """A locked-out client cannot get in by finally guessing right."""
    for _ in range(ratelimit.MAX_FAILURES):
        gated_client.post("/api/auth/login", json={"password": "wrong"})

    response = gated_client.post("/api/auth/login", json={"password": TEST_PASSWORD})

    assert response.status_code == 429


def test_successful_logins_are_not_counted(gated_client: FlaskClient) -> None:
    """Normal use never runs into the limit, because only failures are recorded."""
    for _ in range(ratelimit.MAX_FAILURES * 2):
        response = gated_client.post(
            "/api/auth/login", json={"password": TEST_PASSWORD}
        )
        assert response.status_code == 200


def test_a_sound_configuration_warns_about_nothing(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    """A set secret and a readable hash produce no startup warnings."""
    monkeypatch.setenv(config.SESSION_SECRET_ENV, "test-secret")
    monkeypatch.setenv(config.PASSWORD_HASH_ENV, TEST_PASSWORD_HASH)

    assert auth.auth_config_warnings() == []


def test_a_missing_hash_warns_that_none_is_configured(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    """An unset hash reads as "not configured", not as an unreadable one."""
    monkeypatch.setenv(config.SESSION_SECRET_ENV, "test-secret")

    warnings: list[str] = auth.auth_config_warnings()

    assert warnings == [
        f"No password configured — set {config.PASSWORD_HASH_ENV} "
        "(generate one with: uv run python -m summa.hashpw)"
    ]


@pytest.mark.parametrize("configured_hash", UNREADABLE_HASHES)
def test_an_unreadable_hash_warns_at_startup(
    monkeypatch: pytest.MonkeyPatch, configured_hash: str
) -> None:
    """The mangling the login path cannot report is caught where an operator looks."""
    monkeypatch.setenv(config.SESSION_SECRET_ENV, "test-secret")
    monkeypatch.setenv(config.PASSWORD_HASH_ENV, configured_hash)

    warnings: list[str] = auth.auth_config_warnings()

    assert len(warnings) == 1
    assert "is not a readable hash" in warnings[0]


def test_a_missing_session_secret_warns_at_startup(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    """An unsigned-key deployment is reported alongside a sound password."""
    monkeypatch.setenv(config.PASSWORD_HASH_ENV, TEST_PASSWORD_HASH)

    warnings: list[str] = auth.auth_config_warnings()

    assert len(warnings) == 1
    assert config.SESSION_SECRET_ENV in warnings[0]
