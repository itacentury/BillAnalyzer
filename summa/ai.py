"""Claude-backed category suggestions for uncategorized invoices.

Wraps a single Anthropic request behind :func:`suggest_categories` so the route
layer stays thin and the logic is testable without a live API call. The request
uses structured output (a JSON schema), so the response is guaranteed-parseable.
"""

import json
import logging
import os
from dataclasses import dataclass
from typing import Any, Final

import anthropic

logger: logging.Logger = logging.getLogger(__name__)

API_KEY_ENV: Final[str] = "ANTHROPIC_API_KEY"
MODEL: Final[str] = "claude-opus-4-8"
# The output is tiny (one id + category per invoice), so a non-streaming request
# stays well under the SDK's timeout guard even for a full 100-invoice batch.
MAX_TOKENS: Final[int] = 8000

SYSTEM_PROMPT: Final[str] = (
    "You assign exactly one spending category to each invoice, based on its store "
    "name and purchased items. Prefer a category from the provided list of existing "
    "categories; only invent a new, concise category when none of them fit. Keep "
    "category names short and general (e.g. 'Groceries', 'Electronics', 'Finance')."
)

# Structured-output schema: an object wrapping the suggestion array. Every object
# sets `additionalProperties: false` and `required`, so the model cannot drift.
OUTPUT_SCHEMA: Final[dict[str, Any]] = {
    "type": "object",
    "properties": {
        "suggestions": {
            "type": "array",
            "items": {
                "type": "object",
                "properties": {
                    "invoice_id": {"type": "integer"},
                    "category": {"type": "string"},
                },
                "required": ["invoice_id", "category"],
                "additionalProperties": False,
            },
        }
    },
    "required": ["suggestions"],
    "additionalProperties": False,
}


class AiCategorizationError(Exception):
    """Raised when the Claude request fails or returns an unusable response."""


@dataclass
class CategorySuggestion:
    """A single suggested category for one invoice.

    :param is_new: whether ``category`` is absent from the existing category list
        (computed server-side, case-insensitively, so it never depends on the model).
    """

    invoice_id: int
    category: str | None
    is_new: bool


def api_key_configured() -> bool:
    """Return whether an Anthropic API key is present in the environment."""
    return bool(os.environ.get(API_KEY_ENV))


def _extract_text(message: anthropic.types.Message) -> str:
    """Return the first text block of a response, or raise if none is present."""
    for block in message.content:
        if block.type == "text":
            return block.text
    raise AiCategorizationError("Claude response contained no text block")


def suggest_categories(
    invoices: list[dict[str, Any]], existing_categories: list[str]
) -> list[CategorySuggestion]:
    """Ask Claude to categorize each invoice in a single structured request.

    :param invoices: dicts with ``id``, ``store`` and an ``items`` list
        (``item_name`` / ``item_price``), as assembled by the route.
    :param existing_categories: the categories already in use, given to the model
        as context and used to compute :attr:`CategorySuggestion.is_new`.
    :raises AiCategorizationError: on any API error or an unparseable response.
    """
    if not invoices:
        return []

    client: anthropic.Anthropic = anthropic.Anthropic()
    user_payload: dict[str, Any] = {
        "existing_categories": existing_categories,
        "invoices": invoices,
    }

    try:
        message: anthropic.types.Message = client.messages.create(
            model=MODEL,
            max_tokens=MAX_TOKENS,
            thinking={"type": "adaptive"},
            system=SYSTEM_PROMPT,
            output_config={"format": {"type": "json_schema", "schema": OUTPUT_SCHEMA}},
            messages=[{"role": "user", "content": json.dumps(user_payload)}],
        )
    except anthropic.APIError as error:
        logger.error("Claude categorization request failed: %s", error)
        raise AiCategorizationError(str(error)) from error

    raw: str = _extract_text(message)
    try:
        parsed: Any = json.loads(raw)
    except json.JSONDecodeError as error:
        raise AiCategorizationError("Claude response was not valid JSON") from error

    # Case-insensitive lookup so a suggestion that only differs in casing from an
    # existing category is not falsely flagged as new.
    existing_lower: set[str] = {category.lower() for category in existing_categories}

    suggestions: list[CategorySuggestion] = []
    for entry in parsed.get("suggestions", []):
        category: str | None = entry.get("category") or None
        is_new: bool = category is not None and category.lower() not in existing_lower
        suggestions.append(
            CategorySuggestion(
                invoice_id=entry["invoice_id"],
                category=category,
                is_new=is_new,
            )
        )
    return suggestions
