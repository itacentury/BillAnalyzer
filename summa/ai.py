"""Claude-backed category suggestions for uncategorized invoices.

Wraps a single Anthropic request behind :func:`suggest_categories` so the route
layer stays thin and the logic is testable without a live API call. The request
uses structured output (a JSON schema), so the response is guaranteed-parseable.

Prompt-injection note: the request embeds user-controlled store and item names,
so a crafted name could try to steer the model. The blast radius is bounded — the
output schema locks the shape to ``invoice_id`` + ``category``, and each returned
category is length-capped and whitespace-normalized via
:func:`summa.helpers.clean_category` before it can be surfaced or persisted.
"""

import json
import logging
import os
from dataclasses import dataclass
from typing import Any, Final, Literal

import anthropic
from anthropic.types import (
    OutputConfigParam,
    ThinkingConfigAdaptiveParam,
    ThinkingConfigDisabledParam,
)

from summa.helpers import MAX_CATEGORY_LENGTH, clean_category

logger: logging.Logger = logging.getLogger(__name__)

API_KEY_ENV: Final[str] = "ANTHROPIC_API_KEY"

# User-selectable models, keyed by the short identifier the client sends. The
# resolver below maps a key to its real model id, so an arbitrary client string
# never reaches the API.
MODELS: Final[dict[str, str]] = {
    "haiku": "claude-haiku-4-5",
    "sonnet": "claude-sonnet-5",
    "opus": "claude-opus-4-8",
}
DEFAULT_MODEL_KEY: Final[str] = "haiku"

# Models supporting adaptive thinking + the effort parameter (4.6+). The default
# claude-haiku-4-5 is an older model that rejects both, so it runs with thinking
# disabled instead.
ADAPTIVE_MODELS: Final[frozenset[str]] = frozenset({MODELS["sonnet"], MODELS["opus"]})

# The output is one id + category per invoice, so it stays tiny; the budget is
# sized to the batch mostly to leave room for thinking tokens (which count
# against max_tokens) without truncating the JSON. The cap keeps the request
# under the SDK's ~10-minute non-streaming timeout guard, so no streaming needed.
_TOKEN_BUDGET_BASE: Final[int] = 4096
_TOKENS_PER_INVOICE: Final[int] = 64
_TOKEN_BUDGET_CAP: Final[int] = 15000

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
                    "category": {"type": "string", "maxLength": MAX_CATEGORY_LENGTH},
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


def resolve_model(key: str | None) -> str:
    """Map a user-selected model key to its real model id.

    :param key: the short key from the client (``haiku`` / ``sonnet`` / ``opus``).
    :returns: the model id, falling back to the default for unknown or missing keys.
    """
    return MODELS.get(key or DEFAULT_MODEL_KEY, MODELS[DEFAULT_MODEL_KEY])


def _max_tokens_for(invoice_count: int) -> int:
    """Size the output budget to the batch so thinking + JSON don't truncate."""
    budget: int = _TOKEN_BUDGET_BASE + invoice_count * _TOKENS_PER_INVOICE
    return min(_TOKEN_BUDGET_CAP, budget)


def _thinking_config(
    model: str,
) -> tuple[
    ThinkingConfigAdaptiveParam | ThinkingConfigDisabledParam, Literal["low"] | None
]:
    """Return the ``(thinking, effort)`` pair tuned to the model's capabilities.

    :param model: the resolved model id (see :func:`resolve_model`).
    :returns: the ``thinking`` request value and an optional ``effort`` level;
        adaptive models get low effort, others run with thinking disabled.
    """
    if model in ADAPTIVE_MODELS:
        return {"type": "adaptive"}, "low"
    return {"type": "disabled"}, None


def _extract_text(message: anthropic.types.Message) -> str:
    """Return the first text block of a response, or raise if none is present."""
    for block in message.content:
        if block.type == "text":
            return block.text
    raise AiCategorizationError("Claude response contained no text block")


def suggest_categories(
    invoices: list[dict[str, Any]], existing_categories: list[str], model: str
) -> list[CategorySuggestion]:
    """Ask Claude to categorize each invoice in a single structured request.

    :param invoices: dicts with ``id``, ``store`` and an ``items`` list
        (``item_name`` / ``item_price``), as assembled by the route.
    :param existing_categories: the categories already in use, given to the model
        as context and used to compute :attr:`CategorySuggestion.is_new`.
    :param model: the resolved model id (see :func:`resolve_model`).
    :raises AiCategorizationError: on any API error or an unparseable response.
    """
    if not invoices:
        return []

    client: anthropic.Anthropic = anthropic.Anthropic()
    user_payload: dict[str, Any] = {
        "existing_categories": existing_categories,
        "invoices": invoices,
    }

    thinking, effort = _thinking_config(model)
    output_config: OutputConfigParam = {
        "format": {"type": "json_schema", "schema": OUTPUT_SCHEMA}
    }
    if effort is not None:
        output_config["effort"] = effort

    try:
        message: anthropic.types.Message = client.messages.create(
            model=model,
            max_tokens=_max_tokens_for(len(invoices)),
            thinking=thinking,
            system=SYSTEM_PROMPT,
            output_config=output_config,
            messages=[{"role": "user", "content": json.dumps(user_payload)}],
        )
    except anthropic.APIError as error:
        logger.error("Claude categorization request failed: %s", error)
        raise AiCategorizationError(str(error)) from error

    # A truncated response would otherwise fail JSON parsing below with a
    # misleading "not valid JSON"; surface the real cause instead.
    if message.stop_reason == "max_tokens":
        raise AiCategorizationError(
            "Claude response was truncated (token budget exceeded) — "
            "try a smaller batch"
        )

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
        category: str | None = clean_category(entry.get("category"))
        is_new: bool = category is not None and category.lower() not in existing_lower
        suggestions.append(
            CategorySuggestion(
                invoice_id=entry["invoice_id"],
                category=category,
                is_new=is_new,
            )
        )
    return suggestions
