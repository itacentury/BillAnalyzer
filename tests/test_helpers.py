"""Unit tests for the pure helper functions in :mod:`summa.helpers`."""

from typing import Any

import pytest

from summa.helpers import strip_text


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
