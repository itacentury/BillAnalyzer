"""Unit tests for the pure helper functions in :mod:`summa.helpers`."""

from typing import Any

import pytest

from summa.helpers import escape_like, strip_text


@pytest.mark.parametrize(
    ("value", "expected"),
    [
        (None, None),
        ("  hello  ", "hello"),
        ("hello", "hello"),
        ("", None),
        ("   ", None),
        ("\t\n", None),
        (5, "5"),
        (0, "0"),
        (3.5, "3.5"),
    ],
)
def test_strip_text(value: Any, expected: str | None) -> None:
    """strip_text normalizes whitespace and maps empty input to None."""
    assert strip_text(value) == expected


@pytest.mark.parametrize(
    ("value", "expected"),
    [
        ("ab", "ab"),
        ("a_b", "a\\_b"),
        ("50%", "50\\%"),
        ("a\\b", "a\\\\b"),
    ],
)
def test_escape_like(value: str, expected: str) -> None:
    """escape_like backslash-escapes LIKE wildcards and the escape char."""
    assert escape_like(value) == expected
