"""Integration tests for the AI categorize-suggest route (Anthropic call mocked)."""

from typing import Any

import pytest
from flask.testing import FlaskClient

from summa import ai
from summa.routes import invoices as invoices_route
from tests.conftest import SeedInvoice


def _stub_suggestions(
    monkeypatch: pytest.MonkeyPatch,
) -> list[list[dict[str, Any]]]:
    """Patch suggest_categories to echo a category per invoice; capture its input.

    Returns a one-element list that will hold the invoices the route passed, so a
    test can assert on exactly which invoices were sent to the model.
    """
    captured: list[list[dict[str, Any]]] = []

    def _fake(
        invoices: list[dict[str, Any]], existing_categories: list[str]
    ) -> list[ai.CategorySuggestion]:
        captured.append(invoices)
        return [
            ai.CategorySuggestion(
                invoice_id=invoice["id"], category="Guessed", is_new=True
            )
            for invoice in invoices
        ]

    monkeypatch.setattr(invoices_route, "suggest_categories", _fake)
    return captured


def test_returns_503_without_api_key(
    client: FlaskClient, monkeypatch: pytest.MonkeyPatch
) -> None:
    """With no ANTHROPIC_API_KEY the endpoint reports it is not configured."""
    monkeypatch.delenv("ANTHROPIC_API_KEY", raising=False)

    response = client.post("/api/invoices/categorize-suggest")

    assert response.status_code == 503
    assert response.get_json()["error"] == "AI categorization not configured"


def test_only_uncategorized_are_sent(
    client: FlaskClient,
    seed_invoice: SeedInvoice,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    """Only invoices with a NULL category are collected for suggestion."""
    monkeypatch.setenv("ANTHROPIC_API_KEY", "test-key")
    captured = _stub_suggestions(monkeypatch)

    uncategorized_id = seed_invoice(
        store="Bakery",
        category=None,
        total=14.5,
        items=[{"item_name": "Bread", "item_price": 3.9}],
    )
    seed_invoice(store="Shop", category="Groceries", total=5.0)

    response = client.post("/api/invoices/categorize-suggest")

    assert response.status_code == 200
    body = response.get_json()
    assert body["count"] == 1
    assert body["total"] == 1
    suggestion = body["suggestions"][0]
    assert suggestion["invoice_id"] == uncategorized_id
    assert suggestion["store"] == "Bakery"
    assert suggestion["total"] == 14.5
    assert suggestion["category"] == "Guessed"
    assert suggestion["is_new"] is True
    assert suggestion["items"] == [{"item_name": "Bread", "item_price": 3.9}]
    # Only the uncategorized invoice reached the model.
    assert [invoice["id"] for invoice in captured[0]] == [uncategorized_id]


def test_period_filter_is_applied(
    client: FlaskClient,
    seed_invoice: SeedInvoice,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    """The shared date filter narrows which uncategorized invoices are collected."""
    monkeypatch.setenv("ANTHROPIC_API_KEY", "test-key")
    _stub_suggestions(monkeypatch)

    in_range = seed_invoice(store="InRange", category=None, date="2024-03-10")
    seed_invoice(store="OutOfRange", category=None, date="2024-05-10")

    response = client.post(
        "/api/invoices/categorize-suggest?date_from=2024-03-01&date_to=2024-03-31"
    )

    body = response.get_json()
    assert body["count"] == 1
    assert body["suggestions"][0]["invoice_id"] == in_range


def test_run_is_capped_and_reports_total(
    client: FlaskClient,
    seed_invoice: SeedInvoice,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    """More than the cap: only the first 100 are sent, `total` reports the full set."""
    monkeypatch.setenv("ANTHROPIC_API_KEY", "test-key")
    monkeypatch.setattr(invoices_route, "CATEGORIZE_SUGGEST_LIMIT", 3)
    captured = _stub_suggestions(monkeypatch)

    for index in range(5):
        seed_invoice(store=f"Store {index}", category=None)

    response = client.post("/api/invoices/categorize-suggest")

    body = response.get_json()
    assert body["total"] == 5
    assert body["count"] == 3
    assert len(captured[0]) == 3


def test_no_uncategorized_returns_empty(
    client: FlaskClient,
    seed_invoice: SeedInvoice,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    """With nothing to categorize the endpoint returns an empty result, no API call."""
    monkeypatch.setenv("ANTHROPIC_API_KEY", "test-key")
    captured = _stub_suggestions(monkeypatch)
    seed_invoice(store="Shop", category="Groceries")

    response = client.post("/api/invoices/categorize-suggest")

    assert response.get_json() == {"suggestions": [], "count": 0, "total": 0}
    assert captured == []  # suggest_categories was never called


def test_list_endpoint_reports_uncategorized_count(
    client: FlaskClient, seed_invoice: SeedInvoice
) -> None:
    """GET /api/invoices exposes the uncategorized count that backs the trigger badge."""
    seed_invoice(store="A", category=None)
    seed_invoice(store="B", category=None)
    seed_invoice(store="C", category="Groceries")

    body = client.get("/api/invoices").get_json()

    assert body["uncategorized_count"] == 2
