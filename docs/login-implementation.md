# Login Implementation — Complete Reference

A full, copy-paste-ready description of Tickr's authentication: a **single-password
gate** backed by a **stateless, signed `HttpOnly` session cookie**. Every code block is
copied from the current sources; JavaScript excerpts taken from inside a function body
are dedented to stand alone, but are otherwise unchanged.

This document is written to be **portable**. Blockquotes marked

> **Porting note:** …

flag everything that is Tickr-specific and say what to change for another project.

**Contents**

1. [Overview & threat model](#1-overview--threat-model)
2. [Architecture at a glance](#2-architecture-at-a-glance)
3. [Backend — core helpers](#3-backend--core-helpers)
4. [Backend — endpoints](#4-backend--endpoints)
5. [Backend — middleware gating](#5-backend--middleware-gating)
6. [Backend — error envelope](#6-backend--error-envelope)
7. [Backend — configuration](#7-backend--configuration)
8. [Backend — hardening around login](#8-backend--hardening-around-login)
9. [HTTP contract](#9-http-contract)
10. [Frontend — the login gate module](#10-frontend--the-login-gate-module)
11. [Frontend — the boot gate](#11-frontend--the-boot-gate)
12. [Frontend — styling](#12-frontend--styling)
13. [Frontend — logout flow](#13-frontend--logout-flow)
14. [Frontend — 401 handling in sync](#14-frontend--401-handling-in-sync)
15. [Frontend — service worker & offline](#15-frontend--service-worker--offline)
16. [Same-origin vs. cross-origin](#16-same-origin-vs-cross-origin)
17. [Tests](#17-tests)
18. [Porting checklist](#18-porting-checklist)
19. [Known limitations](#19-known-limitations)

---

## 1. Overview & threat model

The design in one paragraph: there is **one password** for the whole deployment. No user
table, no registration, no roles. On a successful login the server sets a cookie whose
value is a constant string signed with HMAC and stamped with a creation time
(`itsdangerous.URLSafeTimedSerializer`). Every subsequent request is authenticated by
re-verifying that signature and age — **nothing is stored server-side**. The cookie is
`HttpOnly`, so JavaScript never sees a token; the browser attaches it automatically.

What this buys you:

- **Zero session storage.** No table, no migration, no cleanup job, no Redis. Horizontal
  scaling works for auth (all instances sharing the same secret validate the same cookie).
- **No token handling in the client.** No refresh logic, no `Authorization` headers, no
  XSS-readable token in `localStorage`.
- **Tiny surface.** Two dependencies (`argon2-cffi`, `itsdangerous`), ~100 lines of
  backend code.

What it costs you:

- **No per-session revocation.** "Log out everywhere" is impossible without rotating
  `SESSION_SECRET` (which logs out _everyone_). Logout only deletes the browser cookie.
- **No multi-user, no audit trail.** Every request is "the user".
- **No sliding renewal.** The cookie is never re-issued, so a session hard-expires
  `TICKR_SESSION_DAYS` after login regardless of activity.

**Use this model when** the deployment serves one person or a small trusted group and the
goal is "keep the public internet out", not "distinguish between users". **Do not use it
when** you need per-user data, revocation, roles, or an audit trail — at that point you
need a real user table and server-side sessions, and only sections 5, 8, 9, 14 and 15 of
this document remain transferable.

---

## 2. Architecture at a glance

```
BOOT
  main.js  ──GET /api/v1/auth/me──▶  always 200 {authed, enabled}
     │
     ├─ authed:true  ──▶ startApp()  (RxDB + replication start here, not before)
     │
     └─ authed:false ──▶ renderLoginView()
                              │
                              └─POST /api/v1/auth/login {password, remember}
                                     │
                                     ├─401 {error:{code:"UNAUTHORIZED"}} ──▶ "Invalid password"
                                     └─200 + Set-Cookie: tickr_session=… ──▶ startApp()

STEADY STATE
  any protected request ──▶ auth_middleware ──▶ is_authenticated(cookie)?
                                                  ├─ yes ──▶ route handler
                                                  └─ no  ──▶ 401 UNAUTHORIZED

EXPIRY (mid-session)
  replication pull/push sees 401
     └─ handleAuthExpired()   [latched, fires once]
           ├─ pauseSSE()               (stop the reconnect storm)
           └─ authExpired$.next()
                 └─ main.js ──▶ renderLoginView(() => resumeReplication())
                                     └─ re-login ──▶ SSE reopened + rep.reSync(), NO reload

LOGOUT
  logoutBtn ──POST /api/v1/auth/logout──▶ delete_cookie ──▶ location.reload() ──▶ BOOT
```

### File inventory

| File                                                                 | Role                                                                             |
| -------------------------------------------------------------------- | -------------------------------------------------------------------------------- |
| `backend/auth.py`                                                    | Password verification, token create/verify, `is_authenticated`, startup warnings |
| `backend/routes/auth.py`                                             | `POST /login`, `POST /logout`, `GET /me`, cookie attributes                      |
| `backend/main.py`                                                    | `auth_middleware`, public path allow-list, rate limiter, security headers        |
| `backend/config.py`                                                  | All `TICKR_*` env vars                                                           |
| `backend/errors.py`                                                  | `ErrorCode.UNAUTHORIZED`, `AppError`, the JSON error envelope                    |
| `tests/conftest.py`                                                  | `auth_enabled` / `authed_client` fixtures                                        |
| `tests/test_auth.py`, `test_auth_middleware.py`, `test_sync_auth.py` | Backend auth tests                                                               |
| `frontend/src/auth.js`                                               | `getAuthStatus`, `logout`, `renderLoginView` (the whole login UI)                |
| `frontend/src/main.js`                                               | Boot gate + `authExpired$` subscription                                          |
| `frontend/src/bus.js`                                                | `authExpired$` Subject                                                           |
| `frontend/src/db/replication.js`                                     | 401 detection, SSE pause/resume, `resumeReplication`                             |
| `frontend/src/styles/auth.css`                                       | Login gate styling                                                               |
| `frontend/partials/modal-settings.html`                              | Sign-out button markup (`hidden` by default)                                     |
| `frontend/src/events/settings.js`                                    | Sign-out click handler                                                           |
| `frontend/public/sw.js`                                              | Never caches `/api/`; caches shells only when `response.ok`                      |
| `frontend/vite.config.js`                                            | `/api` dev proxy — what makes the cookie same-origin in dev                      |

Notably absent: there is **no** user model, **no** sessions table
(`backend/database.py` only has `lists`, `items`, `categories`, `history`, `settings`),
**no** registration, **no** admin bootstrap, **no** JWT, **no** refresh flow, and **no**
`Depends()` auth guard anywhere — gating is 100% middleware.

---

## 3. Backend — core helpers

`backend/auth.py`, complete and verbatim:

```python
"""Single-password authentication helpers.

Encapsulates password verification (argon2) and stateless, signed session
cookies (itsdangerous) so the middleware in ``main.py`` and the routes in
``routes/auth.py`` stay thin.

Design notes:
    Sessions are *stateless*: the cookie carries a signed, timestamped payload
    and is verified on every request. There is no server-side session store,
    so sessions cannot be revoked individually — acceptable for a single-user
    deployment. Logout simply clears the cookie in the browser.

    All configuration (secret, password, cookie flags) is read lazily from
    ``backend.config`` at call time rather than captured at import, so tests can
    monkeypatch the config module to exercise the authenticated paths.
"""

import secrets

from argon2 import PasswordHasher
from argon2.exceptions import VerifyMismatchError
from fastapi import Request
from itsdangerous import BadSignature, URLSafeTimedSerializer

from . import config
from .logging_config import get_logger

logger = get_logger(__name__)

SESSION_COOKIE_NAME = "tickr_session"

# Opaque marker stored in the signed cookie. The cookie's security comes from
# the signature + timestamp, not from this value.
_SESSION_PAYLOAD = "tickr"
_SESSION_SALT = "tickr-session"

_password_hasher = PasswordHasher()


def _serializer() -> URLSafeTimedSerializer:
    """Build a serializer from the current session secret."""
    return URLSafeTimedSerializer(config.SESSION_SECRET, salt=_SESSION_SALT)


def verify_password(password: str) -> bool:
    """Check a candidate password against the configured hash or plaintext.

    Prefers ``PASSWORD_HASH`` (argon2). Falls back to a constant-time
    comparison against ``PASSWORD_PLAINTEXT`` for local development.
    """
    if config.PASSWORD_HASH:
        try:
            return _password_hasher.verify(config.PASSWORD_HASH, password)
        except VerifyMismatchError:
            return False
        except Exception:
            logger.warning("auth_password_hash_invalid")
            return False

    if config.PASSWORD_PLAINTEXT:
        return secrets.compare_digest(password, config.PASSWORD_PLAINTEXT)

    return False


def create_session_token() -> str:
    """Create a signed, timestamped session token."""
    return _serializer().dumps(_SESSION_PAYLOAD)


def verify_session_token(token: str, max_age_seconds: int) -> bool:
    """Verify a session token's signature and age."""
    try:
        _serializer().loads(token, max_age=max_age_seconds)
        return True
    except BadSignature:
        # Covers SignatureExpired (a BadSignature subclass) too.
        return False


def is_authenticated(request: Request) -> bool:
    """Return whether the request carries a valid session cookie."""
    token: str | None = request.cookies.get(SESSION_COOKIE_NAME)
    if not token:
        return False
    return verify_session_token(token, config.SESSION_DAYS_DEFAULT * 86400)


def auth_config_warnings() -> list[str]:
    """Return startup misconfiguration warnings (empty when fine).

    Only meaningful when ``AUTH_ENABLED`` is true.
    """
    warnings: list[str] = []
    if not config.SESSION_SECRET:
        warnings.append("TICKR_SESSION_SECRET is not set — sessions cannot be signed")
    if not config.PASSWORD_HASH and not config.PASSWORD_PLAINTEXT:
        warnings.append("No password configured (set TICKR_PASSWORD_HASH or TICKR_PASSWORD)")
    elif not config.PASSWORD_HASH and config.PASSWORD_PLAINTEXT:
        warnings.append("Using plaintext TICKR_PASSWORD — set TICKR_PASSWORD_HASH for production")
    return warnings
```

### Details worth understanding before you copy this

**Lazy config reads are load-bearing.** Note `from . import config` followed by
`config.SESSION_SECRET` at call time — never `from .config import SESSION_SECRET`. The
module-level `import` form would freeze the value at import time, and since
`backend/config.py` reads `os.getenv` at import, tests could no longer flip auth on. The
`auth_enabled` fixture (§17) monkeypatches attributes on the live `config` module; that
only works because both `auth.py` and `main.py` dereference lazily.

**The payload is meaningless on purpose.** `_SESSION_PAYLOAD = "tickr"` is a constant. All
security comes from the HMAC signature over `payload + timestamp` and from the max-age
check on `loads()`. There is nothing to look up and nothing user-specific to leak.

**`BadSignature` covers expiry.** `itsdangerous.SignatureExpired` subclasses
`BadSignature`, so the single `except BadSignature` handles both "forged" and "too old".
Do not add a separate `SignatureExpired` branch expecting different behaviour — there is
none.

**Two password sources, hash wins.** If `PASSWORD_HASH` is set, the plaintext value is
never consulted. The bare `except Exception` around `verify()` catches
`InvalidHashError`/`InvalidHash` from a malformed hash string (a very common
configuration mistake — e.g. a Docker Compose-mangled `$`) and fails closed with a log
line instead of a 500. The plaintext branch uses `secrets.compare_digest`, not `==`, to
avoid a timing side channel.

**The server-side max age ignores "remember".** `is_authenticated` always validates
against `SESSION_DAYS_DEFAULT * 86400`. The `remember` checkbox only changes the
browser-side `Max-Age` (§4). A non-remember cookie that a browser happens to keep is
still accepted by the server for the full 30 days.

> **Porting note:** rename `SESSION_COOKIE_NAME`, `_SESSION_PAYLOAD`, `_SESSION_SALT` and
> the `TICKR_*` config names. The salt should be unique per application: it namespaces the
> signature so a cookie from another app sharing the same secret cannot be replayed.
> `logging_config.get_logger` is Tickr's structlog wrapper — swap in `logging.getLogger`
> if you have no structured logging.

---

## 4. Backend — endpoints

`backend/routes/auth.py`, complete and verbatim:

```python
"""Authentication endpoints: login, logout, and session status.

These routes are exempt from the auth middleware (see ``_PUBLIC_EXACT_PATHS``
in ``main.py``) and perform their own cookie checks where needed.
"""

from typing import Literal, TypedDict, cast

from fastapi import APIRouter, Request, Response
from pydantic import BaseModel

from .. import config
from ..auth import (
    SESSION_COOKIE_NAME,
    create_session_token,
    is_authenticated,
    verify_password,
)
from ..errors import AppError, ErrorCode

router = APIRouter(prefix="/api/v1/auth", tags=["auth"])


class LoginRequest(BaseModel):
    """Login payload."""

    password: str
    remember: bool = False


class _CookieAttrs(TypedDict):
    """Typed attributes for the session cookie."""

    httponly: bool
    secure: bool
    samesite: Literal["lax", "strict", "none"]
    path: str


def _cookie_attrs() -> _CookieAttrs:
    """Cookie attributes shared by the set and delete paths.

    Keeping a single source of truth ensures the logout ``delete_cookie`` call
    matches the attributes used when the cookie was set, so browsers reliably
    clear it.
    """
    return {
        "httponly": True,
        "secure": config.COOKIE_SECURE,
        "samesite": cast(Literal["lax", "strict", "none"], config.COOKIE_SAMESITE),
        "path": "/",
    }


def _set_session_cookie(response: Response, *, remember: bool) -> None:
    """Attach a signed session cookie to the response.

    With ``remember`` the cookie persists for ``SESSION_DAYS_DEFAULT`` days;
    otherwise it is a session cookie that the browser drops on close.
    """
    response.set_cookie(
        key=SESSION_COOKIE_NAME,
        value=create_session_token(),
        max_age=config.SESSION_DAYS_DEFAULT * 86400 if remember else None,
        **_cookie_attrs(),
    )


@router.post("/login")
def login(body: LoginRequest, response: Response) -> dict[str, bool]:
    """Verify the password and start a session on success."""
    if not verify_password(body.password):
        raise AppError(ErrorCode.UNAUTHORIZED, "Invalid password", 401)
    _set_session_cookie(response, remember=body.remember)
    return {"authed": True}


@router.post("/logout")
def logout(response: Response) -> dict[str, bool]:
    """Clear the session cookie. Always succeeds."""
    response.delete_cookie(key=SESSION_COOKIE_NAME, **_cookie_attrs())
    return {"authed": False}


@router.get("/me")
def me(request: Request) -> dict[str, bool]:
    """Report authentication status for the client-side login gate.

    Returns 200 in both cases (not 401) so the gate can poll without producing
    console errors. When auth is disabled the user is always considered authed.
    ``enabled`` lets the client decide whether to show a logout control.
    """
    authed: bool = not config.AUTH_ENABLED or is_authenticated(request)
    return {"authed": authed, "enabled": config.AUTH_ENABLED}
```

### Three decisions to carry over

**`_cookie_attrs()` is shared between set and delete.** This is the single most common
bug in cookie-session implementations: browsers match a deletion against the cookie's
`Path`, `Domain`, `Secure` and `SameSite`. If `delete_cookie` is called with different
attributes than `set_cookie` used, the browser keeps the original cookie and logout
silently does nothing. Deriving both from one function makes the mismatch impossible.

**`/me` never returns 401.** It reports status as a 200 payload. That means the boot gate
can call it on every page load without red errors in the console and without the browser
treating it as a failed request. It also carries `enabled`, which lets the client decide
whether the whole password feature exists at all — that is what hides the "Sign out"
button when the deployment runs unauthenticated.

**`remember` maps only to the browser `Max-Age`.** `max_age=None` produces a _session
cookie_ (dropped when the browser closes); `max_age=2592000` produces a persistent one.
The server-side validity is unaffected (§3).

---

## 5. Backend — middleware gating

Authentication is enforced by one middleware, **deny-by-default**, with an explicit
public allow-list. There are zero `Depends()` guards on individual routes — nothing can be
forgotten on a new endpoint, because new endpoints are protected automatically.

From `backend/main.py`, verbatim:

```python
# Routes reachable without a session. The app shell, PWA assets and the auth
# endpoints themselves must stay public — otherwise there is no UI to show the
# login, and the service worker would cache a 401 for "/".
_PUBLIC_EXACT_PATHS: frozenset[str] = frozenset(
    {
        "/",
        "/manifest.json",
        "/sw.js",
        "/circuit-breaker.js",
        "/api/v1/health",
        "/api/v1/auth/login",
        "/api/v1/auth/logout",
        "/api/v1/auth/me",
    }
)
_PUBLIC_PREFIXES: tuple[str, ...] = ("/assets/", "/static/", "/icons/")


def _is_public(path: str) -> bool:
    """Return whether a request path is reachable without authentication."""
    if path in _PUBLIC_EXACT_PATHS:
        return True
    return any(path.startswith(prefix) for prefix in _PUBLIC_PREFIXES)


# Declared first so it becomes the *innermost* middleware: rate limiting and
# access logging wrap it, so login attempts are rate-limited and 401s are logged
# and still receive security headers on the way out.
@app.middleware("http")
async def auth_middleware(request: Request, call_next: CallNext) -> Response:
    """Require a valid session cookie for protected (data/API) routes."""
    if not config.AUTH_ENABLED or _is_public(request.url.path):
        return await call_next(request)
    if is_authenticated(request):
        return await call_next(request)
    return JSONResponse(
        status_code=401,
        content=_error_body(ErrorCode.UNAUTHORIZED, "Authentication required", 401),
    )
```

And the startup warning hook, from the `lifespan` context manager:

```python
    if config.AUTH_ENABLED:
        for warning in auth_config_warnings():
            logger.warning("auth_config_warning", detail=warning)
        logger.info("auth_enabled")
```

### Middleware ordering (Starlette gotcha)

In Starlette/FastAPI, `@app.middleware("http")` **prepends** to the stack, so the
_declaration order is the reverse of the wrapping order_. Tickr declares, in order:
`auth_middleware`, `security_headers_middleware`, `rate_limit_middleware`,
`access_log_and_metrics_middleware`. Execution therefore wraps like this:

```
access_log_and_metrics   ← outermost (sees every response, incl. 429 and 401)
  └── rate_limit         ← login attempts are counted before auth runs
        └── security_headers  ← 401s still get CSP/HSTS/etc.
              └── auth   ← innermost
                    └── route handler
```

Getting this backwards is a real security bug: if auth were the _outermost_ middleware,
its 401s would bypass the rate limiter (unlimited brute-forcing) and the security headers.

### What is protected

Everything not in the allow-list: `/api/v1/lists`, `/api/v1/items`, `/api/v1/categories`,
`/api/v1/settings`, `/api/v1/history`, `/api/v1/sync/*`, `/api/v1/events`,
`/api/v1/metrics` — and the OpenAPI docs (`/api/docs`, `/api/redoc`,
`/api/openapi.json`), which is deliberate.

### Why the shell must stay public

`/`, `/manifest.json`, `/sw.js`, `/circuit-breaker.js`, `/assets/`, `/static/`, `/icons/`
are unauthenticated because **the login form is part of the app bundle**. If `/` returned
401 there would be no UI to log in with — and worse, the service worker would cache that
401 as the app shell (§15), bricking the PWA until the cache is cleared manually.
`/api/v1/health` stays public so uptime probes work without credentials.

> **Porting note:** the exact allow-list is Tickr's. The invariant to preserve is:
> _everything needed to render the login screen is public; everything that returns data is
> not._ For a server-rendered app the equivalent is the login page route plus its assets.

---

## 6. Backend — error envelope

All errors share one JSON shape, which the frontend keys off. From `backend/errors.py`:

```python
class ErrorCode(StrEnum):
    """Machine-readable error codes returned by the API."""

    # ... other codes ...
    UNAUTHORIZED = "UNAUTHORIZED"


class AppError(Exception):
    """Application-level error that renders as a structured JSON response.

    Args:
        code: A stable machine-readable error code from ``ErrorCode``.
        message: A human-readable description of the error.
        status_code: The HTTP status code for the response.
    """

    code: ErrorCode
    message: str
    status_code: int

    def __init__(self, code: ErrorCode, message: str, status_code: int) -> None:
        """Initialize the error with its code, message, and HTTP status."""
        self.code = code
        self.message = message
        self.status_code = status_code
        super().__init__(message)


def _error_body(code: str, message: str, status: int) -> dict[str, Any]:
    """Build the canonical error response body."""
    return {"error": {"code": code, "message": message, "status": status}}


async def _app_error_handler(_request: Request, exc: Exception) -> JSONResponse:
    """Handle AppError exceptions with structured JSON."""
    assert isinstance(exc, AppError)
    return JSONResponse(
        status_code=exc.status_code,
        content=_error_body(exc.code, exc.message, exc.status_code),
    )


def register_error_handlers(app: FastAPI) -> None:
    """Register all custom exception handlers on the FastAPI app."""
    app.add_exception_handler(AppError, _app_error_handler)
    app.add_exception_handler(RequestValidationError, _validation_error_handler)
```

Two paths produce a 401, and they share the envelope but not the mechanism:

- `routes/auth.py` **raises** `AppError(ErrorCode.UNAUTHORIZED, "Invalid password", 401)` —
  it runs inside the routing layer, so the registered exception handler converts it.
- `auth_middleware` **returns** a `JSONResponse` built from `_error_body(...)` directly —
  exception handlers do not apply to middleware that runs outside the router, so the
  middleware must construct the body itself. That is why `_error_body` is a separate,
  reusable function rather than logic hidden inside the handler.

`register_error_handlers(app)` is called once in `main.py` right after the app is created.

---

## 7. Backend — configuration

From `backend/config.py`, verbatim (the helpers plus the auth block):

```python
def _env_bool(name: str, default: bool) -> bool:
    """Read a boolean env var (``true``/``false``, case-insensitive)."""
    return os.getenv(name, str(default).lower()).lower() == "true"


def _env_int(name: str, default: int) -> int:
    """Read an integer env var, falling back to ``default``."""
    return int(os.getenv(name, str(default)))


# --- Authentication ---------------------------------------------------------
# Single-password gate. Defaults to OFF so the existing test suite (which never
# authenticates) and unconfigured deployments keep working unchanged.
AUTH_ENABLED: bool = _env_bool("TICKR_AUTH_ENABLED", False)

# Password source. The argon2 hash is preferred; the plaintext fallback exists
# only for local development and emits a startup warning when used.
PASSWORD_HASH: str = os.getenv("TICKR_PASSWORD_HASH", "")
PASSWORD_PLAINTEXT: str = os.getenv("TICKR_PASSWORD", "")

# Secret used to sign session cookies (itsdangerous). Required when auth is on.
SESSION_SECRET: str = os.getenv("TICKR_SESSION_SECRET", "")

# How long a "remember me" session stays valid, in days.
SESSION_DAYS_DEFAULT: int = _env_int("TICKR_SESSION_DAYS", 30)

# Cookie flags. `Secure` must be off for local dev over plain HTTP, on in prod.
COOKIE_SECURE: bool = _env_bool("TICKR_COOKIE_SECURE", True)
COOKIE_SAMESITE: str = os.getenv("TICKR_COOKIE_SAMESITE", "lax")
```

### Variables

| Variable                | Default   | Meaning                                | Dev          | Prod            |
| ----------------------- | --------- | -------------------------------------- | ------------ | --------------- |
| `TICKR_AUTH_ENABLED`    | `false`   | Master switch for the whole gate       | `true`       | `true`          |
| `TICKR_PASSWORD_HASH`   | _(empty)_ | argon2 hash of the password; preferred | _(empty)_    | **set**         |
| `TICKR_PASSWORD`        | _(empty)_ | Plaintext fallback; warns at startup   | `test1234`   | _(empty)_       |
| `TICKR_SESSION_SECRET`  | _(empty)_ | HMAC key for signing cookies           | `dev-secret` | **set, random** |
| `TICKR_SESSION_DAYS`    | `30`      | Session lifetime, days                 | `30`         | `30`            |
| `TICKR_COOKIE_SECURE`   | `true`    | `Secure` flag on the cookie            | `false`      | `true`          |
| `TICKR_COOKIE_SAMESITE` | `lax`     | `SameSite` flag                        | `lax`        | `lax`           |

Auth-adjacent: `TICKR_CORS_ORIGINS` (also drives the CSP `connect-src`),
`TICKR_TRUSTED_PROXIES` (uvicorn `--forwarded-allow-ips`, needed for correct client IPs in
the rate limiter and for `Secure` cookies behind a TLS-terminating proxy), and the
`TICKR_RATE_LIMIT_*` trio.

**Defaulting to OFF is deliberate.** It means an unconfigured deployment and the existing
test suite keep working unchanged, and auth is an opt-in feature rather than something
that breaks the app when a secret is missing. The trade-off — a forgotten
`TICKR_AUTH_ENABLED=true` silently leaves the app open — is mitigated by the startup
warnings from `auth_config_warnings()`.

### Generating the secrets

```bash
# Session secret
python -c "import secrets; print(secrets.token_urlsafe(32))"

# argon2 password hash (preferred over plaintext)
uv run python -c "from argon2 import PasswordHasher; print(PasswordHasher().hash('your-password'))"
```

> **Docker Compose & the `$` in argon2 hashes:** an argon2 hash (`$argon2id$v=19$m=...`)
> contains `$` characters that Docker Compose treats as variable references, so it mangles
> the hash and logs `WARN The "argon2id" variable is not set`. Disable interpolation using
> the long `env_file` syntax:
>
> ```yaml
> env_file:
>   - path: tickr.env
>     format: raw
> ```
>
> Alternatively double every `$` to `$$`. Verify with
> `docker exec tickr printenv TICKR_PASSWORD_HASH`.

Reference `tickr.env.example` block:

```bash
# Authentication (single password gate) — disabled by default
TICKR_AUTH_ENABLED=true
# Preferred: an argon2 hash (see README for how to generate one).
TICKR_PASSWORD_HASH=
# Dev-only fallback: plaintext password (logs a warning when used).
TICKR_PASSWORD=test1234
# Secret used to sign session cookies. Required when auth is enabled.
TICKR_SESSION_SECRET=dev-secret
# "Stay signed in" duration in days.
TICKR_SESSION_DAYS=30
# Set false for local plain-HTTP testing; keep true behind HTTPS.
TICKR_COOKIE_SECURE=false
TICKR_COOKIE_SAMESITE=lax
```

**Dependencies:** `argon2-cffi>=25.1.0` and `itsdangerous>=2.2.0`. That is the entire auth
dependency footprint.

> **Porting note:** the backend deliberately does **not** call `load_dotenv` — it reads the
> process environment only. Locally, run `uvicorn --env-file .env`. If you prefer
> `pydantic-settings` in the new project, keep the _lazy access_ property (§3): read
> settings through an object attribute at call time so tests can patch it.

---

## 8. Backend — hardening around login

### Rate limiting (brute-force protection)

Per-IP sliding window, in-process. The critical decision is in the exemption list:

```python
# Paths exempt from per-IP rate limiting: real-time streams, monitoring probes,
# and the static app shell + PWA assets. Auth endpoints are deliberately NOT
# exempt so login stays brute-force limited. Locking a client out of the shell
# (``/``, ``/sw.js``, ...) is the worst failure mode — a request storm could
# brick the PWA — so these cheap FileResponses bypass the limiter while every
# ``/api/...`` call stays counted.
_RATE_LIMIT_EXEMPT_PATHS: frozenset[str] = frozenset(
    {
        "/",
        "/manifest.json",
        "/sw.js",
        "/circuit-breaker.js",
        "/api/v1/events",
        "/api/v1/sync/stream",
        "/api/v1/health",
        "/api/v1/metrics",
    }
)
```

The limiter itself:

```python
@app.middleware("http")
async def rate_limit_middleware(request: Request, call_next: CallNext) -> Response:
    """Enforce per-IP sliding window rate limiting, excluding exempt paths."""
    if _is_rate_limit_exempt(request.url.path):
        return await call_next(request)

    client_ip: str = request.client.host if request.client else "unknown"
    now: float = time.time()

    with rate_limit_lock:
        cutoff: float = now - RATE_LIMIT_WINDOW
        timestamps: list[float] = [
            timestamp for timestamp in rate_limit_store[client_ip] if timestamp > cutoff
        ]
        rate_limit_store[client_ip] = timestamps

        if len(timestamps) >= RATE_LIMIT_REQUESTS:
            retry_after: int = max(1, int(timestamps[0] - cutoff) + 1)
            logger.warning(
                "rate_limit_exceeded", client_ip=client_ip, retry_after_seconds=retry_after
            )
            return JSONResponse(
                status_code=429,
                content={
                    "error": {
                        "code": ErrorCode.RATE_LIMITED,
                        "message": "Too many requests",
                        "status": 429,
                    }
                },
                headers={"Retry-After": str(retry_after)},
            )

        timestamps.append(now)

        # Evict only after appending so the current client's entry is non-empty
        # and won't be removed as "stale" by _evict_stale_entries.
        if len(rate_limit_store) > RATE_LIMIT_MAX_IPS:
            _evict_stale_entries(now)

    return await call_next(request)
```

Defaults: 100 requests / 60 s / 10 000 tracked IPs. Two caveats to carry over:

1. **Process-local state.** `rate_limit_store` is a `defaultdict` in memory guarded by a
   `threading.Lock`. Running `uvicorn --workers N` silently breaks it — each worker sees
   only its slice of traffic, so the effective limit becomes `N × 100`. Deploy single
   process, or move to `slowapi` + Redis.
2. **Client IP behind a proxy.** `request.client.host` is the immediate peer, which behind
   nginx/traefik collapses every client onto the proxy IP — making the limiter useless.
   Run uvicorn with `--proxy-headers --forwarded-allow-ips=<trusted proxy>`. uvicorn only
   honours `X-Forwarded-For` when the peer matches that list, so a misconfiguration fails
   _closed_ (everyone looks like the proxy) rather than letting an attacker spoof IPs.

> **Porting note:** 100/60s is a general API limit that happens to also cover login. If
> login is your only concern, a tighter dedicated bucket on `/auth/login` (e.g. 10/5 min)
> is stricter without throttling normal app traffic.

### Security headers and the CSP

```python
@app.middleware("http")
async def security_headers_middleware(request: Request, call_next: CallNext) -> Response:
    """Attach security headers to every response."""
    response: Response = await call_next(request)
    response.headers["Content-Security-Policy"] = (
        "default-src 'self'; "
        "style-src 'self' https://fonts.googleapis.com; "
        "font-src 'self' https://fonts.gstatic.com; "
        "img-src 'self' data:; "
        f"connect-src {CSP_CONNECT_SRC}"
    )
    response.headers["X-Content-Type-Options"] = "nosniff"
    response.headers["X-Frame-Options"] = "DENY"
    response.headers["Referrer-Policy"] = "strict-origin-when-cross-origin"
    response.headers["Permissions-Policy"] = "camera=(), microphone=(), geolocation=()"
    # HSTS only on HTTPS — sending it over plain HTTP is meaningless (browsers
    # ignore it) and misleading. `request.url.scheme` reflects
    # X-Forwarded-Proto when uvicorn runs with --proxy-headers. No `preload`:
    # that is a near-irreversible deployer decision, not ours.
    if request.url.scheme == "https":
        response.headers["Strict-Transport-Security"] = "max-age=31536000; includeSubDomains"
    return response
```

**This CSP is why the login form is built with `document.createElement` instead of a
template.** `default-src 'self'` with no `script-src 'unsafe-inline'` forbids inline
scripts and inline event handlers. `X-Frame-Options: DENY` additionally prevents
clickjacking of the login screen.

The CSP `connect-src` mirrors the CORS allow-list, because the browser enforces the page's
own CSP _before_ the server's CORS response is consulted — a permissive CORS policy
without a matching `connect-src` silently fails:

```python
CSP_CONNECT_SRC: str = " ".join(["'self'", *CORS_ORIGINS])
```

### CSRF posture

There is **no CSRF token**. The defence is:

1. `SameSite=lax` on the session cookie — the browser does not attach it to cross-site
   `POST` requests, which covers the state-changing endpoints.
2. An explicit CORS origin allow-list with `allow_credentials=True`:

```python
app.add_middleware(
    CORSMiddleware,
    allow_origins=CORS_ORIGINS,
    allow_credentials=True,
    allow_methods=["GET", "POST", "PUT", "DELETE", "OPTIONS"],
    allow_headers=["*"],
    expose_headers=["Content-Type"],
)
```

Note that CORS is _not_ a CSRF defence by itself (simple form posts are not preflighted) —
`SameSite=lax` is what actually does the work here.

> **Porting note:** the moment you need `SameSite=none` (a genuinely cross-site frontend),
> this posture collapses and you **must** add a real CSRF token — double-submit cookie or
> a synchronizer token. Also never set `allow_origins=["*"]` together with
> `allow_credentials=True`; browsers reject the combination, and it would be unsafe if they
> did not.

---

## 9. HTTP contract

This is the stack-agnostic middle of the document: implement this contract on the server
and the frontend in §10–§14 works unchanged, and vice versa.

### `POST /api/v1/auth/login`

```http
POST /api/v1/auth/login
Content-Type: application/json

{"password": "…", "remember": true}
```

`remember` is optional and defaults to `false`.

Success — **200**:

```http
HTTP/1.1 200 OK
Set-Cookie: tickr_session=InRpY2tyIg.apf_eQ.9EA8zR_jc7y1th7N0vofBzi_FGw; HttpOnly; Max-Age=2592000; Path=/; SameSite=lax
Content-Type: application/json

{"authed": true}
```

`Max-Age=2592000` (30 days) is present only when `remember` is `true`; otherwise the
`Set-Cookie` carries no `Max-Age` and the browser drops it on close. `Secure` is present
only when `TICKR_COOKIE_SECURE=true`.

Failure — **401**:

```json
{
  "error": {
    "code": "UNAUTHORIZED",
    "message": "Invalid password",
    "status": 401
  }
}
```

No cookie is set on failure.

### `POST /api/v1/auth/logout`

No body, no cookie required. Always **200**:

```http
HTTP/1.1 200 OK
Set-Cookie: tickr_session=""; expires=Wed, 02 Sep 2026 10:50:33 GMT; HttpOnly; Max-Age=0; Path=/; SameSite=lax

{"authed": false}
```

### `GET /api/v1/auth/me`

Always **200**, never 401:

```json
{ "authed": true, "enabled": true }
```

| `enabled` | `authed` | Meaning                                                             |
| --------- | -------- | ------------------------------------------------------------------- |
| `false`   | `true`   | Auth is off deployment-wide; everything is open, hide the logout UI |
| `true`    | `true`   | Auth is on and this client has a valid session                      |
| `true`    | `false`  | Auth is on and this client must log in                              |

`enabled: false, authed: false` is not reachable.

### Any protected route without a valid session — **401**

```json
{
  "error": {
    "code": "UNAUTHORIZED",
    "message": "Authentication required",
    "status": 401
  }
}
```

Note the message differs from the login failure (`"Invalid password"`) while the code is
the same — clients should branch on `status`/`code`, not on the message.

---

## 10. Frontend — the login gate module

`frontend/src/auth.js`, complete and verbatim. This single file is the entire login UI:
status check, logout helper, and a DOM-built login form.

```js
/**
 * Client-side login gate.
 *
 * The app shell is always served publicly; only the data/API routes are
 * protected server-side. This module checks the session, renders a login
 * view when needed (built via the DOM API — no inline scripts, since the CSP
 * forbids `unsafe-inline`), and exposes a logout helper.
 */

/**
 * Fetch the current auth status.
 *
 * @returns {Promise<{authed: boolean, enabled: boolean}>}
 *   `authed` — whether the user may use the app.
 *   `enabled` — whether the password gate is active at all (drives the logout UI).
 */
export async function getAuthStatus() {
  try {
    const response = await fetch("/api/v1/auth/me");
    if (!response.ok) return { authed: false, enabled: true };
    const data = await response.json();
    return { authed: data.authed === true, enabled: data.enabled === true };
  } catch {
    // Offline / network error: let the app try to start (offline-first PWA).
    // Protected requests will surface a 401 later and re-trigger the gate.
    return { authed: true, enabled: false };
  }
}

/**
 * Log out: clear the server session cookie.
 *
 * @returns {Promise<void>}
 */
export async function logout() {
  try {
    await fetch("/api/v1/auth/logout", { method: "POST" });
  } catch {
    // Ignore network errors — the cookie may already be gone.
  }
}

const EYE_OPEN =
  '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M2 12s3.5-7 10-7 10 7 10 7-3.5 7-10 7-10-7-10-7z"/><circle cx="12" cy="12" r="3"/></svg>';
const EYE_OFF =
  '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M9.9 4.24A9.12 9.12 0 0 1 12 4c6.5 0 10 7 10 7a13.16 13.16 0 0 1-1.67 2.68"/><path d="M6.61 6.61A13.526 13.526 0 0 0 2 12s3.5 7 10 7a9.74 9.74 0 0 0 5.39-1.61"/><path d="M14.12 14.12a3 3 0 1 1-4.24-4.24"/><line x1="2" y1="2" x2="22" y2="22"/></svg>';

/**
 * Build the password field: input plus a press-and-hold reveal toggle.
 *
 * @returns {{ wrap: HTMLDivElement, input: HTMLInputElement }}
 */
function buildPasswordField() {
  const wrap = document.createElement("div");
  wrap.className = "auth-password";

  const input = document.createElement("input");
  input.type = "password";
  input.className = "auth-input";
  input.placeholder = "Password";
  input.autocomplete = "current-password";
  input.required = true;

  const toggle = document.createElement("button");
  toggle.type = "button";
  toggle.className = "auth-password-toggle";
  toggle.setAttribute("aria-label", "Show password (press and hold)");
  toggle.innerHTML = EYE_OPEN;

  // Password is only visible while the toggle is held down.
  const show = () => {
    input.type = "text";
    toggle.innerHTML = EYE_OFF;
  };
  const hide = () => {
    input.type = "password";
    toggle.innerHTML = EYE_OPEN;
  };

  // Keep focus on the password input so Enter still submits the form
  // instead of re-triggering this toggle button.
  toggle.addEventListener("mousedown", (event) => event.preventDefault());

  toggle.addEventListener("pointerdown", show);
  for (const event of ["pointerup", "pointerleave", "pointercancel", "blur"]) {
    toggle.addEventListener(event, hide);
  }
  toggle.addEventListener("keydown", (event) => {
    if (event.key === " " || event.key === "Enter") {
      event.preventDefault();
      show();
    }
  });
  toggle.addEventListener("keyup", (event) => {
    if (event.key === " " || event.key === "Enter") {
      hide();
    }
  });

  wrap.appendChild(input);
  wrap.appendChild(toggle);
  return { wrap, input };
}

/**
 * Render the login view, replacing the app UI until login succeeds.
 *
 * @param {() => void} onSuccess - Called once after a successful login.
 */
export function renderLoginView(onSuccess) {
  // Hide the app shell while the gate is shown.
  const appEl = /** @type {HTMLElement | null} */ (
    document.querySelector(".app")
  );
  if (appEl) appEl.style.display = "none";

  // Reuse an existing gate if present (e.g. session expired mid-session).
  document.querySelector('[data-el="authGate"]')?.remove();

  const gate = document.createElement("div");
  gate.dataset.el = "authGate";
  gate.className = "auth-gate";

  const card = document.createElement("form");
  card.className = "auth-card";
  card.noValidate = true;

  const title = document.createElement("h1");
  title.className = "auth-title";
  title.textContent = "Tickr";

  const subtitle = document.createElement("p");
  subtitle.className = "auth-subtitle";
  subtitle.textContent = "Enter your password to continue";

  const { wrap: passwordWrap, input: passwordInput } = buildPasswordField();

  const rememberLabel = document.createElement("label");
  rememberLabel.className = "auth-remember";
  const rememberCheckbox = document.createElement("input");
  rememberCheckbox.type = "checkbox";
  const rememberText = document.createElement("span");
  rememberText.textContent = "Stay signed in for 30 days";
  rememberLabel.appendChild(rememberCheckbox);
  rememberLabel.appendChild(rememberText);

  const error = document.createElement("p");
  error.className = "auth-error";
  error.hidden = true;

  const submit = document.createElement("button");
  submit.type = "submit";
  submit.className = "auth-submit";
  submit.textContent = "Sign in";

  card.appendChild(title);
  card.appendChild(subtitle);
  card.appendChild(passwordWrap);
  card.appendChild(rememberLabel);
  card.appendChild(error);
  card.appendChild(submit);
  gate.appendChild(card);
  document.body.appendChild(gate);

  passwordInput.focus();

  card.addEventListener("submit", async (event) => {
    event.preventDefault();
    error.hidden = true;
    submit.disabled = true;
    submit.textContent = "Signing in…";

    try {
      const response = await fetch("/api/v1/auth/login", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
          password: passwordInput.value,
          remember: rememberCheckbox.checked,
        }),
      });

      if (!response.ok) {
        error.textContent = "Invalid password";
        error.hidden = false;
        passwordInput.select();
        return;
      }

      gate.remove();
      if (appEl) appEl.style.display = "";
      onSuccess();
    } catch {
      error.textContent = "Network error — please try again";
      error.hidden = false;
    } finally {
      submit.disabled = false;
      submit.textContent = "Sign in";
    }
  });
}
```

### The equivalent markup

If your project has no CSP restriction on inline scripts, you can ship the gate as a
template instead. This is exactly what the DOM code above produces:

```html
<div data-el="authGate" class="auth-gate">
  <form class="auth-card" novalidate>
    <h1 class="auth-title">Tickr</h1>
    <p class="auth-subtitle">Enter your password to continue</p>
    <div class="auth-password">
      <input
        type="password"
        class="auth-input"
        placeholder="Password"
        autocomplete="current-password"
        required
      />
      <button
        type="button"
        class="auth-password-toggle"
        aria-label="Show password (press and hold)"
      >
        <!-- EYE_OPEN svg -->
      </button>
    </div>
    <label class="auth-remember">
      <input type="checkbox" />
      <span>Stay signed in for 30 days</span>
    </label>
    <p class="auth-error" hidden></p>
    <button type="submit" class="auth-submit">Sign in</button>
  </form>
</div>
```

### Behavioural details that are easy to lose

- **`<form novalidate>` + `required`.** The form element gives you Enter-to-submit and
  browser password-manager integration for free. `noValidate` suppresses the native
  validation bubble so the inline `.auth-error` paragraph is the only error UI — but
  `required` stays on the input as an accessibility hint.
- **`autocomplete="current-password"`** is what makes password managers offer to fill and
  save. Without it many managers ignore the field entirely.
- **The reveal toggle is press-and-hold, not a latch.** `pointerdown` shows;
  `pointerup`, `pointerleave`, `pointercancel` and `blur` all hide. Four hide events
  rather than one because a pointer can leave the button, be cancelled by a gesture, or
  the element can lose focus — any of which would otherwise strand the password in plain
  text. Space/Enter `keydown`/`keyup` mirror the same behaviour for keyboard users.
- **`mousedown` is `preventDefault()`ed** on the toggle so focus never leaves the password
  input. Without it, pressing Enter after using the toggle would re-trigger the toggle
  button instead of submitting the form.
- **`passwordInput.select()` on failure** pre-selects the wrong password so retyping
  overwrites it.
- **The gate removes any pre-existing `[data-el="authGate"]` first**, so a mid-session
  expiry cannot stack two gates on top of each other.
- **`getAuthStatus()` fails open on network errors.** A rejected `fetch` yields
  `authed: true` so an offline PWA still boots from IndexedDB. A non-OK _HTTP_ response,
  by contrast, fails closed (`authed: false`). See §15 for the subtle interaction with
  the service worker.
- **The `.app` shell is hidden via inline `style.display`**, restored with
  `style.display = ""` on success — deliberately not a class, so it cannot collide with
  the app's own display rules.

> **Porting note:** rename the "Tickr" title, the "Stay signed in for 30 days" label (keep
> it in sync with `TICKR_SESSION_DAYS`), and the `.app` selector. The two SVG constants are
> Feather-style icons — replace freely.

---

## 11. Frontend — the boot gate

From `frontend/src/main.js`, verbatim:

```js
import { reportError } from "./error-reporting.js";
import { initApp } from "./app.js";
import { getAuthStatus, renderLoginView } from "./auth.js";
import { resumeReplication } from "./db/replication.js";
import { authExpired$ } from "./bus.js";
import { accountSettingGroup } from "./dom.js";

/** Start the app, reporting any init failure. */
function startApp() {
  initApp().catch((err) => {
    reportError("initialize app", err);
  });
}

if (!inRecovery) {
  // Auth gate: only initialize the app (and thus replication) once
  // authenticated. The logout control is revealed only when the password gate
  // is actually on.
  getAuthStatus()
    .then(({ authed, enabled }) => {
      const reveal = () => {
        if (enabled && accountSettingGroup) accountSettingGroup.hidden = false;
      };
      if (authed) {
        startApp();
        reveal();
      } else {
        renderLoginView(() => {
          startApp();
          reveal();
        });
      }
    })
    .catch((err) => reportError("check auth", err));

  // Session expired/revoked mid-session: drop back to the login view, then
  // resume sync in place once re-login succeeds — no full reload needed.
  authExpired$.subscribe(() => {
    renderLoginView(() => resumeReplication());
  });
}
```

`auth.css` is imported last among the stylesheets in `main.js`, so it wins over earlier
sheets at equal specificity.

### Reading the gate

**`initApp()` runs only after authentication.** This is the important structural
property: RxDB is created and replication is started _inside_ `initApp`, so an
unauthenticated visitor never opens IndexedDB and never fires a single sync request. The
login callback is the only other place `startApp` is called.

**Two different success callbacks.** Boot passes `startApp` (nothing exists yet);
expiry passes `resumeReplication` (RxDB and the collections already exist and must not be
re-created). Passing `startApp` on the expiry path would double-initialize the database.

**`enabled` drives only the logout UI.** When the deployment runs with
`TICKR_AUTH_ENABLED=false`, `authed` is `true` and `enabled` is `false`, so the app starts
and the "Sign out" group stays `hidden` — no dead button for a feature that is off.

**`inRecovery`** is unrelated to auth: `circuit-breaker.js` (a classic script running
before the bundle) sets `window.__tickrRecovery` when it detects a reload loop and takes
over the page. Both the auth gate and the service-worker registration are skipped in that
case so neither fights the recovery screen. Drop this guard if your project has no such
breaker.

---

## 12. Frontend — styling

`frontend/src/styles/auth.css`, complete and verbatim:

```css
/* Login gate */
.auth-gate {
  position: fixed;
  inset: 0;
  z-index: 1000;
  display: flex;
  align-items: center;
  justify-content: center;
  padding: 24px;
  background: var(--bg-primary);
}

.auth-card {
  display: flex;
  flex-direction: column;
  gap: 16px;
  width: 100%;
  max-width: 360px;
  padding: 32px;
  background: var(--bg-secondary);
  border: 1px solid var(--border);
  border-radius: var(--radius-lg);
  box-shadow: var(--shadow-lg);
}

.auth-title {
  margin: 0;
  font-size: 1.75rem;
  font-weight: 700;
  color: var(--text-primary);
  text-align: center;
}

.auth-subtitle {
  margin: 0 0 8px;
  font-size: 0.9rem;
  color: var(--text-secondary);
  text-align: center;
}

.auth-input {
  width: 100%;
  padding: 12px 14px;
  font-size: 1rem;
  color: var(--text-primary);
  background: var(--bg-tertiary);
  border: 1px solid var(--border-light);
  border-radius: var(--radius-md);
  transition: border-color var(--transition-fast);
}

.auth-input:focus {
  border-color: var(--accent);
}

.auth-password {
  position: relative;
  display: flex;
}

.auth-password .auth-input {
  padding-right: 44px;
}

.auth-password-toggle {
  position: absolute;
  top: 50%;
  right: 8px;
  display: flex;
  align-items: center;
  justify-content: center;
  padding: 4px;
  color: var(--text-secondary);
  cursor: pointer;
  background: transparent;
  border: none;
  transform: translateY(-50%);
  transition: color var(--transition-fast);
}

.auth-password-toggle:hover,
.auth-password-toggle:focus-visible {
  color: var(--accent);
}

.auth-password-toggle svg {
  width: 20px;
  height: 20px;
}

.auth-remember {
  display: flex;
  gap: 8px;
  align-items: center;
  font-size: 0.875rem;
  color: var(--text-secondary);
  cursor: pointer;
  user-select: none;
}

.auth-remember input {
  accent-color: var(--accent);
}

.auth-error {
  margin: 0;
  font-size: 0.875rem;
  color: var(--danger);
}

.auth-submit {
  padding: 12px;
  font-size: 1rem;
  font-weight: 600;
  color: #fff;
  cursor: pointer;
  background: var(--accent-hover);
  border: none;
  border-radius: var(--radius-md);
  transition: background var(--transition-fast);
}

.auth-submit:hover:not(:disabled) {
  background: var(--accent);
}

.auth-submit:disabled {
  cursor: default;
  opacity: 0.6;
}
```

### Custom properties consumed

Define these (in `:root`, or map them to your own tokens) or the sheet renders unstyled:

| Group          | Properties                                        |
| -------------- | ------------------------------------------------- |
| Surfaces       | `--bg-primary`, `--bg-secondary`, `--bg-tertiary` |
| Borders        | `--border`, `--border-light`                      |
| Radii / shadow | `--radius-md`, `--radius-lg`, `--shadow-lg`       |
| Text           | `--text-primary`, `--text-secondary`              |
| Accent / state | `--accent`, `--accent-hover`, `--danger`          |
| Motion         | `--transition-fast`                               |

Two layout notes: `.auth-gate` uses `position: fixed; inset: 0` with `z-index: 1000` so it
covers the app regardless of scroll position, and `.auth-password .auth-input` reserves
`padding-right: 44px` so the absolutely-positioned reveal button never overlaps typed text.

---

## 13. Frontend — logout flow

Markup, from `frontend/partials/modal-settings.html`:

```html
<div class="setting-group" data-el="accountSettingGroup" hidden>
  <label>Account</label>
  <button type="button" class="btn-danger" data-el="logoutBtn">Sign out</button>
</div>
```

DOM hooks, from `frontend/src/dom.js`:

```js
export const accountSettingGroup = el("accountSettingGroup");
export const logoutBtn = el("logoutBtn");
```

Handler, from `frontend/src/events/settings.js`:

```js
// Sign out: clear the server session, then reload so the auth gate re-renders.
dom.logoutBtn?.addEventListener("click", async () => {
  await logout();
  location.reload();
});
```

The flow: `POST /api/v1/auth/logout` (server `delete_cookie` with identical attributes) →
`location.reload()` → `main.js` re-runs `getAuthStatus()` → `{authed: false}` → login gate.

**Why logout hard-reloads but expiry does not.** Expiry is a _recoverable interruption_ —
the user is the same user, the local database is still theirs, so re-login should resume
sync in place without losing UI state. Logout is an intentional _end of session_: a full
reload is the cheapest way to guarantee every in-memory subscription, timer and cached
render is gone.

**Logout does not wipe local data.** The RxDB/IndexedDB contents survive, which is correct
for a single-user offline-first app (logging out on your own device should not destroy
your offline copy). Clearing local data is a separate, explicit action — the "Clear cache
& reload" button in the same modal, which calls `state.db.remove()`, deletes all
`caches`, unregisters the service workers and reloads.

> **Porting note:** on a shared device this is the wrong default. If the new project can be
> used by more than one person on one machine, wipe the local store inside the logout
> handler as well.

---

## 14. Frontend — 401 handling in sync

Once the app is running, the login gate must be able to come _back_ — the session can
expire, or the server can be restarted with a new secret. The pattern is: **one central
401 detector, a latch so it fires once, and a single bus event.**

### The bus event

`frontend/src/bus.js`:

```js
import { Subject } from "rxjs";

/**
 * Fires when a request is rejected with 401 (session expired/missing).
 *
 * Lives here to avoid a circular import between the data layer
 * (replication.js, which detects the 401) and the app entry (main.js,
 * which owns the login gate and re-renders the login view).
 */
export const authExpired$ = new Subject();
```

The event lives on the bus for a structural reason: `replication.js` (data layer) detects
the 401 and `main.js` (entry/view layer) owns the login UI. A direct import either way
would create a cycle. The bus keeps the dependency one-directional.

### The detector and the latch

From `frontend/src/db/replication.js`:

```js
/** Set once a 401 is seen, to stop the SSE reconnect storm. */
let sessionExpired = false;

/**
 * Handle a 401 from any sync request: stop SSE, halt reconnects, and notify
 * the app to show the login gate.
 */
function handleAuthExpired() {
  if (sessionExpired) return;
  sessionExpired = true;
  pauseSSE();
  authExpired$.next();
}
```

The latch matters because three collections (`lists`, `items`, `categories`) replicate
independently and will all 401 within milliseconds of each other. Without it you would get
three stacked login gates.

### The 401 branches

Pull handler:

```js
const response = await fetch(
  `/api/v1/sync/${collection}/pull?${params.toString()}`,
  { signal: AbortSignal.timeout(REPLICATION_FETCH_TIMEOUT_MS) },
);
if (response.status === 401) {
  handleAuthExpired();
  throw new Error(`Pull unauthorized for ${collection}`);
}
```

Push handler:

```js
const response = await fetch(`/api/v1/sync/${collection}/push`, {
  method: "POST",
  headers: { "Content-Type": "application/json" },
  body: JSON.stringify(body),
  signal: AbortSignal.timeout(REPLICATION_FETCH_TIMEOUT_MS),
});
if (response.status === 401) {
  handleAuthExpired();
  throw new Error(`Push unauthorized for ${collection}`);
}
if (!response.ok) {
  throw new Error(`Push failed for ${collection}: ${response.status}`);
}
```

**Both re-throw after signalling, and that is deliberate.** RxDB sees the handler reject
and retries on `retryTime` (5 s), which keeps the replication alive and the pending writes
queued. Swallowing the error would mark the batch as handled and silently drop unsynced
local changes.

### Pause, do not tear down

```js
/**
 * Close the shared SSE connection and stop reconnects, leaving the collection
 * subjects alive so the stream can be reopened later (e.g. after re-login).
 */
function pauseSSE() {
  if (reconnectTimeout) {
    clearTimeout(reconnectTimeout);
    reconnectTimeout = null;
  }
  if (staleTimeout) {
    clearTimeout(staleTimeout);
    staleTimeout = null;
  }
  if (!sharedEventSource) return;

  sharedEventSource.close();
  sharedEventSource = null;
}
```

`pauseSSE()` deliberately does **not** `complete()` the per-collection subjects — only
`cleanupSSE()` (run on unload) does. RxJS subjects cannot be reused after completion, so
completing them here would make resume-in-place impossible and force a page reload.

The SSE error listener honours the latch, which is what actually stops the reconnect
storm:

```js
eventSource.addEventListener("error", () => {
  eventSource.close();
  sharedEventSource = null;
  clearTimeout(staleTimeout);
  if (sessionExpired) return;
  clearTimeout(reconnectTimeout);
  reconnectTimeout = setTimeout(
    () => connectSharedStream(),
    SSE_RECONNECT_DELAY_MS,
  );
});
```

`EventSource` cannot read the HTTP status of a failed connection — a 401 on the stream is
indistinguishable from a network drop. That is precisely why the 401 must be detected on
the pull/push `fetch` calls and latched into `sessionExpired`; the stream then reads the
latch instead of guessing.

### Resume after re-login

```js
/**
 * Resume sync after a successful re-login: clear the expired latch, reopen the
 * SSE stream, and trigger an immediate re-sync of all replications.
 *
 * The collection subjects survive the expired window (only paused, not
 * completed), so replication can restart in place without a full page reload.
 */
export function resumeReplication() {
  if (!sessionExpired) return;
  sessionExpired = false;
  if (Object.keys(collectionSubjects).length > 0) {
    connectSharedStream();
  }
  for (const rep of activeReplications) {
    rep.reSync();
  }
}
```

Relevant constants from `frontend/src/timing.js`:

```js
export const SSE_RECONNECT_DELAY_MS = 3000;
export const SSE_STALE_TIMEOUT_MS = 40000;
export const REPLICATION_FETCH_TIMEOUT_MS = 15000;
export const REPLICATION_RETRY_MS = 5000;
```

### Known gap in the current implementation

**Only the sync pull and push handlers check for 401.** The other API calls —
`/api/v1/settings`, `/api/v1/lists/{id}/history`, `/api/v1/metrics`, `/api/v1/errors` —
do not, and fail silently or show a generic error toast. Re-arming the gate therefore
depends on replication traffic happening. In practice replication runs constantly so the
gap is small, but it is a real one: fix it in the new project by routing every request
through one wrapper.

> **Porting note — projects without RxDB or SSE:** the transferable pattern is the
> detector + latch + event, not the RxDB specifics. The minimal equivalent is a single
> fetch wrapper that every API call goes through:
>
> ```js
> import { authExpired$ } from "./bus.js";
>
> let sessionExpired = false;
>
> /**
>  * Fetch wrapper that centralizes 401 detection.
>  *
>  * @param {string} url - Request URL.
>  * @param {RequestInit} [options] - Fetch options.
>  * @returns {Promise<Response>}
>  */
> export async function apiFetch(url, options = {}) {
>   const response = await fetch(url, options);
>   if (response.status === 401 && !sessionExpired) {
>     sessionExpired = true;
>     authExpired$.next();
>   }
>   return response;
> }
>
> /** Clear the latch after a successful re-login. */
> export function clearAuthExpired() {
>   sessionExpired = false;
> }
> ```
>
> Then `authExpired$.subscribe(() => renderLoginView(clearAuthExpired))` in the entry
> point. If you have no RxJS, a plain callback registry or a `CustomEvent` on `window`
> works just as well.

---

## 15. Frontend — service worker & offline

Two rules in `frontend/public/sw.js` exist specifically to keep the service worker from
poisoning the auth flow.

### API requests are network-only

```js
// API requests - network only (RxDB handles offline data)
if (url.pathname.startsWith("/api/")) {
  event.respondWith(
    fetch(request).catch((error) => {
      if (error.name === "AbortError") {
        throw error;
      }
      return new Response(JSON.stringify({ error: "Offline" }), {
        status: 503,
        headers: { "Content-Type": "application/json" },
      });
    }),
  );
  return;
}
```

Nothing under `/api/` is ever cached, so a 401 can never be served from cache and an
authorized response can never be replayed after logout. The `AbortError` re-throw
preserves `AbortSignal.timeout` semantics for the replication fetches — swallowing it
would turn a timeout into a fake 503 and confuse RxDB's retry logic.

### HTML shells are cached only on success

```js
// HTML pages - network first, fall back to cache
if (
  request.headers.get("accept")?.includes("text/html") ||
  url.pathname === "/"
) {
  event.respondWith(
    fetch(request)
      .then((response) => {
        // Only cache successful shells; never persist a 429/503 as "/".
        if (response.ok) {
          const responseToCache = response.clone();
          caches
            .open(CACHE_NAME)
            .then((cache) => cache.put(request, responseToCache));
        }
        return response;
      })
      .catch(() =>
        caches.match(request).then(
          (cached) =>
            cached ||
            new Response("Offline - no cached version available", {
              status: 503,
              headers: { "Content-Type": "text/plain" },
            }),
        ),
      ),
  );
  return;
}
```

The `response.ok` check pairs with the backend keeping `/` in `_PUBLIC_EXACT_PATHS` (§5).
Together they guarantee the app shell in the cache is always a real shell — never a 401,
429 or 503 — so there is always a UI capable of rendering the login form.

### The subtle offline consequence

`getAuthStatus()` has two failure modes and they behave differently:

| Situation               | `fetch` result   | `getAuthStatus()` returns        | Effect                        |
| ----------------------- | ---------------- | -------------------------------- | ----------------------------- |
| No SW, offline          | rejects          | `{authed: true, enabled: false}` | App boots offline (fail open) |
| SW controlling, offline | resolves **503** | `{authed: false, enabled: true}` | Login form is shown offline   |

With the service worker installed, an offline `/api/v1/auth/me` never rejects — the SW
synthesizes a 503 — so the `!response.ok` branch fires and the user sees a login screen
they cannot get past while offline. This is a wart, not a design goal.

> **Porting note:** if you want a genuinely offline-capable gate, treat 503 like a network
> error in `getAuthStatus()`:
>
> ```js
> if (response.status === 503) return { authed: true, enabled: false };
> if (!response.ok) return { authed: false, enabled: true };
> ```
>
> Or exclude the auth endpoints from the SW's network-only branch and let them reject
> naturally.

---

## 16. Same-origin vs. cross-origin

**No `credentials` option appears anywhere in this codebase.** Every API call uses a
relative URL, so `fetch`'s default of `same-origin` already attaches the cookie. The same
applies to `new EventSource("/api/v1/sync/stream")`.

This holds in both environments:

- **Production:** FastAPI serves the built SPA from `static/dist`, so app and API share an
  origin.
- **Development:** the Vite dev server proxies `/api` to FastAPI, so the browser still sees
  one origin:

```js
  server: {
    proxy: {
      "/api": {
        target: "http://localhost:8000",
        changeOrigin: true,
      },
    },
  },
```

**Dev gotcha:** set `TICKR_COOKIE_SECURE=false` for local plain-HTTP work, otherwise the
`Secure` cookie is dropped over `http://localhost:5173` and login appears to succeed while
every subsequent request 401s.

### If the new project splits frontend and API across origins

All of the following become mandatory:

| Concern       | Change                                                                      |
| ------------- | --------------------------------------------------------------------------- |
| Every `fetch` | add `credentials: "include"`                                                |
| `EventSource` | `new EventSource(url, { withCredentials: true })`                           |
| Cookie        | `SameSite=none; Secure` (which requires HTTPS on both ends)                 |
| CORS          | `allow_credentials=True` **plus** an explicit origin list — never `["*"]`   |
| CSP           | `connect-src` must list the API origin (the browser checks CSP before CORS) |
| CSRF          | now genuinely required — `SameSite=none` removes the only CSRF defence (§8) |

Behind a TLS-terminating reverse proxy, also forward `X-Forwarded-Proto` and run uvicorn
with `--proxy-headers`, so `request.url.scheme` is `https` and `Secure` cookies plus HSTS
behave correctly.

---

## 17. Tests

### Fixtures

From `tests/conftest.py`, verbatim:

```python
# The suite runs with AUTH_ENABLED=false by default (see backend.config), so the
# existing tests need no changes. The fixtures below opt specific tests into the
# authenticated paths by monkeypatching the live config module — both the
# middleware (main.py) and auth helpers (auth.py) read config attributes lazily.
TEST_PASSWORD = "test-password"


@pytest.fixture()
def auth_enabled(monkeypatch) -> str:
    """Turn on auth with a known plaintext password and an insecure cookie.

    Returns the configured password so tests can log in.
    """
    monkeypatch.setattr(config, "AUTH_ENABLED", True)
    monkeypatch.setattr(config, "SESSION_SECRET", "test-secret")
    monkeypatch.setattr(config, "PASSWORD_HASH", "")
    monkeypatch.setattr(config, "PASSWORD_PLAINTEXT", TEST_PASSWORD)
    monkeypatch.setattr(config, "COOKIE_SECURE", False)
    return TEST_PASSWORD


@pytest.fixture()
def authed_client(auth_enabled) -> TestClient:
    """A TestClient with auth enabled and a valid session cookie."""
    test_client = TestClient(app, raise_server_exceptions=False)
    resp = test_client.post(
        "/api/v1/auth/login", json={"password": auth_enabled, "remember": False}
    )
    assert resp.status_code == 200
    return test_client
```

Three things make this work and are worth replicating:

1. **Auth is off by default**, so the entire pre-existing test suite needed no changes when
   auth was introduced. Tests opt _in_ by requesting the fixture.
2. **Monkeypatching the live config module** is only possible because of the lazy reads
   (§3). Freeze the config at import and this fixture becomes impossible.
3. **`COOKIE_SECURE=False`** in tests, because `TestClient` speaks plain HTTP and would
   otherwise discard the cookie — the same trap as local dev.

### Cases to re-create

`tests/test_auth.py`:

- Correct password → 200 `{"authed": true}` and the cookie is set.
- Wrong password → 401 with `error.code == "UNAUTHORIZED"`, no cookie set.
- `remember: true` → `Max-Age=2592000` in `Set-Cookie`.
- `remember: false` → no `Max-Age` in `Set-Cookie`.
- Logout → 200 `{"authed": false}`, and a following `/me` reports
  `{"authed": false, "enabled": true}`.
- `/me` for a logged-in client → `{"authed": true, "enabled": true}`.
- `/me` with auth on but no session → 200 `{"authed": false, "enabled": true}`.
- `/me` with auth off → 200 `{"authed": true, "enabled": false}`.

`tests/test_auth_middleware.py`:

- A protected route (`/api/v1/lists`) → 401 `UNAUTHORIZED` without a session.
- `/api/v1/metrics` → 401 without a session, 200 with one.
- Parametrized public paths — `/`, `/api/v1/health`, `/manifest.json`, `/sw.js`,
  `/assets/anything.js`, `/icons/favicon.ico` — must **never** be 401 (they may 404; the
  assertion is `!= 401`, deliberately, so the test does not depend on files existing).
- The login endpoint itself is reachable without a session.
- With `AUTH_ENABLED=false`, every route is open.

`tests/test_sync_auth.py`:

- Sync pull and push → 401 without a session; pull → 200 with one.

Example, verbatim:

```python
@pytest.mark.parametrize(
    "path",
    [
        "/",
        "/api/v1/health",
        "/manifest.json",
        "/sw.js",
        "/assets/anything.js",
        "/icons/favicon.ico",
    ],
)
def test_public_paths_are_not_blocked(unauth_client, path) -> None:
    """Public asset and health paths are never gated by the auth middleware."""
    # The file may or may not exist in the test env, but it must never be a 401.
    resp = unauth_client.get(path)
    assert resp.status_code != 401
```

Run them with:

```bash
uv run pytest tests/test_auth.py tests/test_auth_middleware.py tests/test_sync_auth.py -v
```

### Frontend coverage: none

There are **no frontend auth tests** in Tickr. `auth.js`, the boot gate in `main.js` and
the `authExpired$` / `resumeReplication` path are entirely untested, and
`replication.test.js` covers converters, the stale-checkpoint guard and SSE staleness but
not the 401 branch.

> **Porting note:** this is the clearest gap to close in the new project. With vitest +
> jsdom the high-value cases are: `getAuthStatus()` for 200/non-OK/reject; `handleAuthExpired`
> firing `authExpired$` exactly once across three concurrent 401s; and `renderLoginView`
> submitting the right body and not stacking gates when called twice.

---

## 18. Porting checklist

Ordered so each step compiles and can be verified before the next.

**Backend**

1. **Add dependencies:** `argon2-cffi`, `itsdangerous`.
2. **Copy `auth.py`** (§3). Rename `SESSION_COOKIE_NAME` (`tickr_session`),
   `_SESSION_SALT` (`tickr-session`) and `_SESSION_PAYLOAD` (`tickr`). Replace
   `get_logger` with your logger. Keep the lazy `config.X` access.
3. **Add config** (§7): `AUTH_ENABLED`, `PASSWORD_HASH`, `PASSWORD_PLAINTEXT`,
   `SESSION_SECRET`, `SESSION_DAYS_DEFAULT`, `COOKIE_SECURE`, `COOKIE_SAMESITE`. Rename the
   `TICKR_` env prefix.
4. **Copy the error envelope** (§6) — `AppError`, `ErrorCode.UNAUTHORIZED`, `_error_body`,
   and register the handlers — or adapt `routes/auth.py` and the middleware to your own
   error shape.
5. **Copy `routes/auth.py`** (§4) and mount the router. Rename the `/api/v1/auth` prefix if
   your API is versioned differently.
6. **Add the auth middleware and the public allow-list** (§5). Declare it **first** so it
   ends up innermost. Adjust the allow-list to your own shell/asset paths — everything
   needed to render the login screen must be public.
7. **Add the startup warnings** from `auth_config_warnings()` to your lifespan/startup
   hook.
8. **Check the ordering of your other middleware:** rate limiting and security headers must
   wrap the auth middleware, not the other way round (§5).
9. **Generate the secret and the password hash** (§7) and put them in your env file — never
   in version control. Watch the Docker Compose `$` trap.

**Frontend**

10. **Copy `auth.js`** (§10). Rename the app title, the "Stay signed in for 30 days" label
    (keep it in sync with `SESSION_DAYS`), the `.app` selector, and the endpoint paths.
11. **Copy `auth.css`** (§12) and define the custom properties it consumes, or map them to
    your design tokens.
12. **Add the boot gate** to your entry point (§11): `getAuthStatus()` → `startApp()` or
    `renderLoginView(startApp)`. Make sure nothing that talks to the API runs before it.
13. **Add the logout UI** (§13): a button hidden unless `enabled === true`, wired to
    `await logout(); location.reload();`. On shared devices, also clear local storage here.
14. **Add the 401 detector** (§14): either the RxDB pull/push branches, or the `apiFetch`
    wrapper from the porting note. Wire `authExpired$` to `renderLoginView`, and clear the
    latch in the success callback.
15. **Set the service worker rules** (§15) if you have one: never cache `/api/`, cache
    shells only when `response.ok`, and decide how you want offline `/auth/me` to behave.
16. **Configure the dev proxy** so the frontend and API are same-origin in dev, and set
    `COOKIE_SECURE=false` locally (§16).

**Verify**

17. **Port the backend tests** (§17) — they are the cheapest way to prove the whole gate
    actually gates. Then add the frontend tests Tickr is missing.

### Rename reference

| Tickr value                  | Where                                             |
| ---------------------------- | ------------------------------------------------- |
| `tickr_session`              | `SESSION_COOKIE_NAME` in `auth.py`                |
| `tickr-session`              | `_SESSION_SALT` in `auth.py`                      |
| `tickr`                      | `_SESSION_PAYLOAD` in `auth.py`                   |
| `TICKR_*`                    | every env var in `config.py`                      |
| `/api/v1/auth`               | router prefix and the three frontend `fetch` URLs |
| `Tickr`                      | the `.auth-title` text                            |
| `Stay signed in for 30 days` | the `.auth-remember` label                        |
| `.app`                       | the shell selector hidden by `renderLoginView`    |
| `authGate`                   | the `data-el` hook on the gate element            |

---

## 19. Known limitations

Deliberate trade-offs of this design, listed so you can decide consciously rather than
discover them later:

- **No per-session revocation.** Sessions are stateless signed cookies. "Sign out
  everywhere" requires rotating `SESSION_SECRET`, which invalidates every session
  including your own. If you need selective revocation you need server-side session
  storage — at which point most of §3 is replaced.
- **No sliding renewal.** The cookie is never re-issued, so an active user is logged out
  exactly `SESSION_DAYS` after login. Adding renewal means re-setting the cookie on
  requests past a threshold age.
- **Single password, single user.** No accounts, no roles, no audit trail. Everyone with
  the password is the same principal.
- **Process-local rate limiting.** Multiple uvicorn workers multiply the effective login
  attempt limit. Single-process deployment only, or move to Redis-backed limiting.
- **No CSRF token.** Protection rests entirely on `SameSite=lax` plus the CORS origin
  list. Any move to `SameSite=none` requires adding one.
- **Only sync requests detect 401.** Other API calls fail silently; gate re-arming depends
  on replication traffic (§14).
- **Offline gate behaviour with the service worker.** An offline `/auth/me` resolves as a
  synthetic 503 and shows an unusable login screen (§15).
- **No frontend auth tests.** The entire client-side gate is untested (§17).
