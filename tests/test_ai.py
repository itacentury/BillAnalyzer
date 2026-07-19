"""Unit tests for :mod:`summa.ai`, with the Anthropic client fully mocked."""

import json
from typing import Any

import anthropic
import httpx
import pytest

from summa import ai


class _FakeBlock:
    """Minimal stand-in for an Anthropic text content block."""

    def __init__(self, text: str) -> None:
        self.type: str = "text"
        self.text: str = text


class _FakeMessage:
    """Minimal stand-in for an Anthropic message response."""

    def __init__(self, text: str) -> None:
        self.content: list[_FakeBlock] = [_FakeBlock(text)]


class _FakeMessages:
    """Records the create() call and returns a canned response (or raises)."""

    def __init__(self, response_text: str | None, error: Exception | None) -> None:
        self._response_text: str | None = response_text
        self._error: Exception | None = error
        self.last_kwargs: dict[str, Any] = {}

    def create(self, **kwargs: Any) -> _FakeMessage:
        self.last_kwargs = kwargs
        if self._error is not None:
            raise self._error
        assert self._response_text is not None
        return _FakeMessage(self._response_text)


class _FakeClient:
    """Stand-in for anthropic.Anthropic exposing a `.messages` attribute."""

    def __init__(self, messages: _FakeMessages) -> None:
        self.messages: _FakeMessages = messages


def _patch_client(
    monkeypatch: pytest.MonkeyPatch,
    *,
    response_text: str | None = None,
    error: Exception | None = None,
) -> _FakeMessages:
    """Patch anthropic.Anthropic to return a fake client; return its messages stub."""
    fake_messages: _FakeMessages = _FakeMessages(response_text, error)
    monkeypatch.setattr(anthropic, "Anthropic", lambda: _FakeClient(fake_messages))
    return fake_messages


def test_suggest_categories_builds_request_with_all_invoices(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    """Every invoice and the existing categories are sent in one request."""
    response: str = json.dumps(
        {
            "suggestions": [
                {"invoice_id": 1, "category": "Groceries"},
                {"invoice_id": 2, "category": "Gaming"},
            ]
        }
    )
    fake_messages = _patch_client(monkeypatch, response_text=response)

    invoices: list[dict[str, Any]] = [
        {"id": 1, "store": "Bakery", "items": [{"item_name": "Bread"}]},
        {"id": 2, "store": "PlayStation", "items": []},
    ]
    ai.suggest_categories(invoices, ["Groceries"])

    sent: dict[str, Any] = json.loads(
        fake_messages.last_kwargs["messages"][0]["content"]
    )
    assert sent["existing_categories"] == ["Groceries"]
    assert [invoice["id"] for invoice in sent["invoices"]] == [1, 2]
    assert fake_messages.last_kwargs["model"] == ai.MODEL


def test_is_new_is_computed_case_insensitively(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    """A suggestion matching an existing category (any case) is not flagged new."""
    response: str = json.dumps(
        {
            "suggestions": [
                {"invoice_id": 1, "category": "groceries"},
                {"invoice_id": 2, "category": "Gaming"},
            ]
        }
    )
    _patch_client(monkeypatch, response_text=response)

    results = ai.suggest_categories(
        [{"id": 1, "store": "A", "items": []}, {"id": 2, "store": "B", "items": []}],
        ["Groceries"],
    )

    by_id = {result.invoice_id: result for result in results}
    assert by_id[1].is_new is False  # differs only in casing
    assert by_id[2].is_new is True


def test_api_error_is_wrapped(monkeypatch: pytest.MonkeyPatch) -> None:
    """An Anthropic APIError surfaces as AiCategorizationError."""
    request: httpx.Request = httpx.Request("POST", "https://api.anthropic.com")
    _patch_client(
        monkeypatch,
        error=anthropic.APIConnectionError(message="boom", request=request),
    )

    with pytest.raises(ai.AiCategorizationError):
        ai.suggest_categories([{"id": 1, "store": "A", "items": []}], [])


def test_invalid_json_is_wrapped(monkeypatch: pytest.MonkeyPatch) -> None:
    """A non-JSON response surfaces as AiCategorizationError."""
    _patch_client(monkeypatch, response_text="not json")

    with pytest.raises(ai.AiCategorizationError):
        ai.suggest_categories([{"id": 1, "store": "A", "items": []}], [])


def test_empty_invoices_short_circuits(monkeypatch: pytest.MonkeyPatch) -> None:
    """No invoices means no request is made and the result is empty."""
    fake_messages = _patch_client(monkeypatch, response_text="{}")

    assert ai.suggest_categories([], []) == []
    assert fake_messages.last_kwargs == {}
