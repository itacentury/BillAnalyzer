"""Shared types and helper functions for the Summa backend."""

from typing import Any

from flask import Response

# Type alias for API responses that may include HTTP status codes
ApiResponse = Response | tuple[Response, int]


def strip_text(value: Any) -> str | None:
    """Strip whitespace from text values, returning None for empty strings."""
    if value is None:
        return None
    stripped: str = str(value).strip()
    return stripped if stripped else None
