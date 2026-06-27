"""Shared types and helper functions for the Summa backend."""

from dataclasses import dataclass
from typing import Any

from flask import Response

# Type alias for API responses that may include HTTP status codes
ApiResponse = Response | tuple[Response, int]


class ValidationError(Exception):
    """Raised when client-supplied invoice data is malformed."""


@dataclass
class InvoiceItem:
    """A single validated invoice line item."""

    item_name: str | None
    item_price: float


@dataclass
class Invoice:
    """A validated invoice parsed from a request payload."""

    date: str | None
    store: str | None
    category: str | None
    total: float
    items: list[InvoiceItem]


def strip_text(value: Any) -> str | None:
    """Strip whitespace from text values, returning None for empty strings."""
    if value is None:
        return None
    stripped: str = str(value).strip()
    return stripped if stripped else None


def _require(data: dict[str, Any], key: str) -> Any:
    """Return data[key], raising ValidationError if the key is absent."""
    if key not in data:
        raise ValidationError(f"Missing required field: {key}")
    return data[key]


def _parse_float(value: Any, field: str) -> float:
    """Convert value to float, raising ValidationError if it is not numeric."""
    try:
        return float(value)
    except (TypeError, ValueError):
        raise ValidationError(f"Field '{field}' must be a number") from None


def parse_invoice(data: Any) -> Invoice:
    """Validate and parse a single invoice payload into an Invoice."""
    if not isinstance(data, dict):
        raise ValidationError("Invoice must be a JSON object")

    items: list[InvoiceItem] = []
    for raw_item in data.get("items", []):
        if not isinstance(raw_item, dict):
            raise ValidationError("Each item must be a JSON object")
        name: str | None = strip_text(_require(raw_item, "item_name"))
        price: float = _parse_float(_require(raw_item, "item_price"), "item_price")
        items.append(InvoiceItem(item_name=name, item_price=price))

    return Invoice(
        date=strip_text(_require(data, "date")),
        store=strip_text(_require(data, "store")),
        category=strip_text(data.get("category")),
        total=_parse_float(_require(data, "total"), "total"),
        items=items,
    )


def parse_invoice_list(data: Any) -> list[Invoice]:
    """Validate and parse a list of invoice payloads (for /import)."""
    if not isinstance(data, list):
        raise ValidationError("Expected a list of invoices")
    return [parse_invoice(item) for item in data]
