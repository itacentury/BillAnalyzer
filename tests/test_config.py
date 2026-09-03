"""Tests for the cookie attribute accessors in :mod:`summa.config`."""

import pytest

from summa import config

# Every SameSite spelling an operator might set, and what it normalizes to.
# The last three are the silent fallback: an unrecognized value must not reach
# Flask, which would raise on it and take the whole app down.
SAMESITE_VALUES: list[tuple[str, str]] = [
    ("lax", "Lax"),
    ("strict", "Strict"),
    ("none", "None"),
    ("STRICT", "Strict"),
    ("  Strict  ", "Strict"),
    ("", "Lax"),
    ("sometimes", "Lax"),
    ("true", "Lax"),
]

# COOKIE_SECURE is a boolean read through _env_bool(), which treats only these
# spellings as off — anything else set is on, so a typo fails safe.
SECURE_VALUES: list[tuple[str, bool]] = [
    ("0", False),
    ("false", False),
    ("FALSE", False),
    ("no", False),
    ("off", False),
    ("", False),
    ("1", True),
    ("true", True),
    ("yes", True),
]


@pytest.mark.parametrize(("configured", "expected"), SAMESITE_VALUES)
def test_cookie_samesite_normalizes_the_configured_value(
    monkeypatch: pytest.MonkeyPatch, configured: str, expected: str
) -> None:
    """Case and padding are the operator's business; Flask wants one spelling."""
    monkeypatch.setenv(config.COOKIE_SAMESITE_ENV, configured)

    assert config.cookie_samesite() == expected


def test_cookie_samesite_defaults_to_lax() -> None:
    """An unset SameSite is Lax, the browser default the gate relies on."""
    assert config.cookie_samesite() == config.DEFAULT_COOKIE_SAMESITE


@pytest.mark.parametrize(("configured", "expected"), SECURE_VALUES)
def test_cookie_secure_reads_the_configured_flag(
    monkeypatch: pytest.MonkeyPatch, configured: str, expected: bool
) -> None:
    """Only the documented falsy spellings turn the Secure attribute off."""
    monkeypatch.setenv(config.COOKIE_SECURE_ENV, configured)

    assert config.cookie_secure() is expected


def test_cookie_secure_defaults_to_on() -> None:
    """Unset means Secure: a deployment must opt out of HTTPS-only, not into it."""
    assert config.cookie_secure() is True
