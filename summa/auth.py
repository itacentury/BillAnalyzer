"""Single-password authentication helpers.

The gate is stateless: a successful login marks Flask's own signed session
cookie, and every later request is authenticated by re-verifying that signature.
Nothing is stored server-side, so sessions cannot be revoked individually —
acceptable for a single-user deployment, where "log out everywhere" means
rotating ``SESSION_SECRET``.

Flask owns the cookie attributes (name, ``HttpOnly``, ``Secure``, ``SameSite``,
lifetime), configured once in :func:`summa.create_app`. That also means
:func:`end_session` clears the cookie with exactly the attributes it was set
with, which is the usual way a hand-rolled logout fails.
"""

import logging
from typing import Final

from flask import session
from werkzeug.security import check_password_hash

from summa import config

logger: logging.Logger = logging.getLogger(__name__)

# Marker stored in the signed session. The security comes from the signature and
# the cookie lifetime, not from this value.
_AUTHED_KEY: str = "authed"

# Reachable without a session. The app shell and the auth endpoints themselves
# must stay public — the login form is part of the bundle, so gating "/" would
# leave no UI to log in with, and the service worker (which holds "/"
# cache-first) would persist the rejection as the app shell.
_PUBLIC_EXACT_PATHS: Final[frozenset[str]] = frozenset(
    {"/", "/api/auth/login", "/api/auth/logout", "/api/auth/me"}
)

# Covers the static folder and the three asset manifests the service worker
# fetches at install time, which live under the same namespace.
_PUBLIC_PREFIXES: Final[tuple[str, ...]] = ("/static/",)


def is_public(path: str) -> bool:
    """Return whether a request path is reachable without authentication."""
    if path in _PUBLIC_EXACT_PATHS:
        return True
    return any(path.startswith(prefix) for prefix in _PUBLIC_PREFIXES)


def is_preflight(method: str, requested_method: str | None) -> bool:
    """Return whether this request is a CORS preflight.

    A preflight is sent without cookies, so it can never authenticate, and it
    returns no data — only the ``Allow`` and CORS headers. It has to pass the
    gate or a cross-origin client's real request is never sent at all.
    """
    return method == "OPTIONS" and bool(requested_method)


def verify_password(candidate: str) -> bool:
    """Check a candidate password against the configured hash.

    Returns ``False`` when no hash is configured, so a half-configured gate
    fails closed (nobody can log in) rather than open.
    """
    configured_hash: str = config.password_hash()
    if not configured_hash:
        return False
    try:
        return check_password_hash(configured_hash, candidate)
    except ValueError:
        # A malformed hash (commonly a Docker Compose-mangled '$') must not
        # surface as a 500 — fail closed and say so once in the log.
        logger.warning("Configured %s is not a valid hash", config.PASSWORD_HASH_ENV)
        return False


def start_session(remember: bool) -> None:
    """Mark the current session as authenticated.

    :param remember: when true the cookie persists for ``SESSION_DAYS``;
        otherwise the browser drops it on close.
    """
    session.clear()
    session[_AUTHED_KEY] = True
    session.permanent = remember


def end_session() -> None:
    """Drop the authenticated session, clearing the cookie in the browser."""
    session.clear()


def is_authenticated() -> bool:
    """Return whether the current request carries a valid session cookie."""
    return session.get(_AUTHED_KEY) is True


def auth_config_warnings() -> list[str]:
    """Return startup misconfiguration warnings (empty when the setup is sound).

    Only meaningful while :func:`summa.config.auth_enabled` is true.
    """
    warnings: list[str] = []
    if not config.session_secret():
        warnings.append(
            f"{config.SESSION_SECRET_ENV} is not set — sessions are signed with an "
            "ephemeral key, so every worker and every restart invalidates them"
        )
    if not config.password_hash():
        warnings.append(
            f"No password configured — set {config.PASSWORD_HASH_ENV} "
            "(generate one with: uv run python -m summa.hashpw)"
        )
    return warnings
