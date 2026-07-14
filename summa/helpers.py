"""Shared types and helper functions for the Summa backend."""

from dataclasses import dataclass
from typing import Any

from flask import Response, jsonify

# Type alias for API responses that may include HTTP status codes
ApiResponse = Response | tuple[Response, int]


def error_response(message: str, status: int) -> tuple[Response, int]:
    """Build a standard {success: False, error: …} JSON error response."""
    return jsonify({"success": False, "error": message}), status


class ValidationError(Exception):
    """Raised when client-supplied invoice data is malformed."""

    def __init__(self, message: str, field: str | None = None) -> None:
        """:param field: the offending field name, if the error is field-specific."""
        super().__init__(message)
        self.field: str | None = field


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


def escape_like(value: str) -> str:
    """Escape LIKE wildcards so a search term matches literally."""
    # Escape the escape char first, then the two LIKE wildcards.
    return value.replace("\\", "\\\\").replace("%", "\\%").replace("_", "\\_")


def parse_bounded_int(value: Any, default: int, minimum: int, maximum: int) -> int:
    """Parse value to an int clamped to [minimum, maximum], falling back to default."""
    try:
        parsed: int = int(value)
    except (TypeError, ValueError):
        return default
    return max(minimum, min(parsed, maximum))


def _require(data: dict[str, Any], key: str) -> Any:
    """Return data[key], raising ValidationError if the key is absent."""
    if key not in data:
        raise ValidationError(f"Missing required field: {key}", field=key)
    return data[key]


def _parse_float(value: Any, field: str) -> float:
    """Convert value to float, raising ValidationError if it is not numeric."""
    try:
        return float(value)
    except (TypeError, ValueError):
        raise ValidationError(
            f"Field '{field}' must be a number", field=field
        ) from None


def parse_invoice(data: Any) -> Invoice:
    """Validate and parse a single invoice payload into an Invoice."""
    if not isinstance(data, dict):
        raise ValidationError("Invoice must be a JSON object")

    items: list[InvoiceItem] = []
    for raw_item in data.get("items") or []:
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


@dataclass
class ImportEntryError:
    """One invalid entry from a batch import: its position, field and reason."""

    index: int
    field: str | None
    message: str
    value: Any  # the raw entry, so the client can prefill an editor


@dataclass
class ImportValidation:
    """The outcome of validating a batch: the valid invoices plus per-entry errors."""

    invoices: list[Invoice]
    errors: list[ImportEntryError]


def parse_invoice_batch(data: Any) -> ImportValidation:
    """Validate a list of invoice payloads, collecting per-entry errors (for /import).

    Unlike :func:`parse_invoice`, a single malformed entry does not abort the batch:
    valid entries are still parsed and each failure is recorded with its array index.
    """
    if not isinstance(data, list):
        raise ValidationError("Expected a list of invoices")

    invoices: list[Invoice] = []
    errors: list[ImportEntryError] = []
    for index, item in enumerate(data):
        try:
            invoices.append(parse_invoice(item))
        except ValidationError as error:
            errors.append(ImportEntryError(index, error.field, str(error), item))
    return ImportValidation(invoices=invoices, errors=errors)
