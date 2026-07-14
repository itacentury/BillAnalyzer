"""Unit tests for the pure helper functions in :mod:`summa.helpers`."""

from typing import Any

import pytest

from summa.helpers import (
    ValidationError,
    escape_like,
    parse_invoice,
    parse_invoice_batch,
    strip_text,
)


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


def test_parse_invoice_batch_all_valid() -> None:
    """A fully valid batch parses every entry and reports no errors."""
    result = parse_invoice_batch(
        [
            {"date": "2024-01-01", "store": "A", "total": 1.0, "items": []},
            {"date": "2024-01-02", "store": "B", "total": 2.0, "items": []},
        ]
    )
    assert [invoice.store for invoice in result.invoices] == ["A", "B"]
    assert result.errors == []


def test_parse_invoice_batch_collects_errors_with_index_and_field() -> None:
    """Invalid entries are collected with their array index, field and raw value."""
    raw_bad = {"date": "2024-01-01", "store": "Bad", "total": "abc"}
    result = parse_invoice_batch(
        [
            {"date": "2024-01-01", "store": "Good", "total": 1.0, "items": []},
            {"store": "NoDate", "total": 1.0},
            raw_bad,
        ]
    )
    assert [invoice.store for invoice in result.invoices] == ["Good"]
    assert [(error.index, error.field) for error in result.errors] == [
        (1, "date"),
        (2, "total"),
    ]
    assert result.errors[1].message == "Field 'total' must be a number"
    assert result.errors[1].value == raw_bad


def test_parse_invoice_non_list_items_raises() -> None:
    """A truthy non-list `items` raises ValidationError instead of a bare TypeError."""
    with pytest.raises(ValidationError, match="Field 'items' must be a list") as info:
        parse_invoice({"date": "2024-01-01", "store": "A", "total": 1.0, "items": 5})
    assert info.value.field == "items"


def test_parse_invoice_batch_non_list_items_does_not_abort_batch() -> None:
    """A bad `items` field is collected per-entry; sibling valid entries still parse."""
    result = parse_invoice_batch(
        [
            {"date": "2024-01-01", "store": "Good", "total": 1.0, "items": []},
            {"date": "2024-01-02", "store": "Bad", "total": 2.0, "items": 5},
        ]
    )
    assert [invoice.store for invoice in result.invoices] == ["Good"]
    assert [(error.index, error.field) for error in result.errors] == [(1, "items")]


def test_parse_invoice_batch_non_list_raises() -> None:
    """A non-list payload raises, so the endpoint can answer 400."""
    with pytest.raises(ValidationError, match="Expected a list of invoices"):
        parse_invoice_batch({"date": "2024-01-01"})


def test_validation_error_exposes_field() -> None:
    """ValidationError carries the offending field for machine-readable reporting."""
    assert ValidationError("bad", field="total").field == "total"
    assert ValidationError("generic").field is None
