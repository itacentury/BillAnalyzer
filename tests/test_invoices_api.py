"""Integration tests for the invoice CRUD, bulk and lookup endpoints."""

from typing import Any

from flask.testing import FlaskClient

from summa import db
from tests.conftest import SeedInvoice


def _get_json(response: Any) -> Any:
    """Return the parsed JSON body of a test-client response."""
    return response.get_json()


# --- POST /api/invoices -------------------------------------------------------


def test_add_invoice_creates_invoice_with_items(client: FlaskClient) -> None:
    """A valid payload creates the invoice and its items and returns the new id."""
    response = client.post(
        "/api/invoices",
        json={
            "date": "2024-03-01",
            "store": "Grocer",
            "category": "Food",
            "total": 25.5,
            "items": [
                {"item_name": "Milk", "item_price": 1.5},
                {"item_name": "Bread", "item_price": 2.0},
            ],
        },
    )

    assert response.status_code == 200
    body = _get_json(response)
    assert body["success"] is True
    assert isinstance(body["id"], int)

    listed = _get_json(client.get("/api/invoices"))
    assert len(listed) == 1
    assert listed[0]["store"] == "Grocer"
    assert len(listed[0]["items"]) == 2


def test_add_invoice_strips_whitespace(client: FlaskClient) -> None:
    """strip_text normalizes surrounding whitespace on text fields."""
    client.post(
        "/api/invoices",
        json={
            "date": "  2024-03-01  ",
            "store": "  Spaced Store  ",
            "category": "  Cat  ",
            "total": 5,
            "items": [],
        },
    )

    invoice = _get_json(client.get("/api/invoices"))[0]
    assert invoice["date"] == "2024-03-01"
    assert invoice["store"] == "Spaced Store"
    assert invoice["category"] == "Cat"


def test_add_invoice_allows_empty_items(client: FlaskClient) -> None:
    """An invoice without items is accepted."""
    response = client.post(
        "/api/invoices",
        json={"date": "2024-03-01", "store": "NoItems", "total": 3.0},
    )
    assert response.status_code == 200
    assert _get_json(client.get("/api/invoices"))[0]["items"] == []


def test_add_invoice_missing_field_returns_400(client: FlaskClient) -> None:
    """A missing required key is rejected with the JSON 400 envelope."""
    response = client.post("/api/invoices", json={"store": "NoDate", "total": 1.0})
    assert response.status_code == 400
    body = _get_json(response)
    assert body["success"] is False
    assert body["error"] == "Missing required field: date"


def test_add_invoice_non_numeric_total_returns_400(client: FlaskClient) -> None:
    """A non-numeric total is rejected with the JSON 400 envelope."""
    response = client.post(
        "/api/invoices",
        json={"date": "2024-03-01", "store": "Shop", "total": "abc"},
    )
    assert response.status_code == 400
    body = _get_json(response)
    assert body["success"] is False
    assert body["error"] == "Field 'total' must be a number"


# --- GET /api/invoices --------------------------------------------------------


def test_get_invoices_excludes_soft_deleted(
    client: FlaskClient, seed_invoice: SeedInvoice
) -> None:
    """Soft-deleted invoices never appear in the listing."""
    seed_invoice(store="Visible")
    seed_invoice(store="Gone", deleted=True)

    stores = [invoice["store"] for invoice in _get_json(client.get("/api/invoices"))]
    assert stores == ["Visible"]


def test_get_invoices_filters_by_store_and_category(
    client: FlaskClient, seed_invoice: SeedInvoice
) -> None:
    """The store and category query parameters filter on exact matches."""
    seed_invoice(store="Aldi", category="Food")
    seed_invoice(store="Rewe", category="Food")
    seed_invoice(store="Aldi", category="Drinks")

    by_store = _get_json(client.get("/api/invoices?store=Aldi"))
    assert len(by_store) == 2

    by_category = _get_json(client.get("/api/invoices?category=Drinks"))
    assert len(by_category) == 1
    assert by_category[0]["store"] == "Aldi"


def test_get_invoices_filters_by_date_range(
    client: FlaskClient, seed_invoice: SeedInvoice
) -> None:
    """date_from and date_to bound the results inclusively."""
    seed_invoice(date="2024-01-01", store="Jan")
    seed_invoice(date="2024-02-01", store="Feb")
    seed_invoice(date="2024-03-01", store="Mar")

    result = _get_json(
        client.get("/api/invoices?date_from=2024-02-01&date_to=2024-02-28")
    )
    assert [invoice["store"] for invoice in result] == ["Feb"]


def test_get_invoices_search_matches_store_and_item(
    client: FlaskClient, seed_invoice: SeedInvoice
) -> None:
    """search matches on the store name as well as item names."""
    seed_invoice(store="Banana Stand")
    seed_invoice(store="Other", items=[{"item_name": "banana bread", "item_price": 3}])
    seed_invoice(store="Unrelated", items=[{"item_name": "apple", "item_price": 1}])

    result = _get_json(client.get("/api/invoices?search=banana"))
    assert len(result) == 2


def test_get_invoices_search_escapes_like_wildcards(
    client: FlaskClient, seed_invoice: SeedInvoice
) -> None:
    """LIKE wildcards in the search term match literally, not as wildcards."""
    seed_invoice(store="Sale")
    seed_invoice(store="S_le")

    underscore = _get_json(client.get("/api/invoices?search=S_le"))
    assert [invoice["store"] for invoice in underscore] == ["S_le"]

    seed_invoice(store="50")
    seed_invoice(store="50%")

    percent = _get_json(client.get("/api/invoices?search=50%25"))
    assert [invoice["store"] for invoice in percent] == ["50%"]


def test_get_invoices_default_sort_is_date_desc(
    client: FlaskClient, seed_invoice: SeedInvoice
) -> None:
    """Without sort_by the listing defaults to date descending."""
    seed_invoice(date="2024-02-01", store="Feb")
    seed_invoice(date="2024-01-01", store="Jan")
    seed_invoice(date="2024-03-01", store="Mar")

    dates = [invoice["date"] for invoice in _get_json(client.get("/api/invoices"))]
    assert dates == ["2024-03-01", "2024-02-01", "2024-01-01"]


def test_get_invoices_sorting(client: FlaskClient, seed_invoice: SeedInvoice) -> None:
    """sort_by/sort_order order the results by the chosen column."""
    seed_invoice(store="A", total=30.0)
    seed_invoice(store="B", total=10.0)
    seed_invoice(store="C", total=20.0)

    ascending = _get_json(client.get("/api/invoices?sort_by=total&sort_order=asc"))
    assert [invoice["total"] for invoice in ascending] == [10.0, 20.0, 30.0]

    descending = _get_json(client.get("/api/invoices?sort_by=total&sort_order=desc"))
    assert [invoice["total"] for invoice in descending] == [30.0, 20.0, 10.0]


def test_get_invoices_ignores_unknown_sort_column(
    client: FlaskClient, seed_invoice: SeedInvoice
) -> None:
    """An out-of-whitelist sort_by is ignored rather than injected into SQL."""
    seed_invoice(store="A", total=1.0)
    seed_invoice(store="B", total=2.0)

    response = client.get("/api/invoices?sort_by=total;DROP TABLE invoices")
    assert response.status_code == 200
    assert len(_get_json(response)) == 2


# --- GET /api/stores and /api/categories --------------------------------------


def test_get_stores_distinct_sorted_without_deleted(
    client: FlaskClient, seed_invoice: SeedInvoice
) -> None:
    """Stores are distinct, alphabetically sorted and exclude deleted rows."""
    seed_invoice(store="Rewe")
    seed_invoice(store="Aldi")
    seed_invoice(store="Aldi")
    seed_invoice(store="Hidden", deleted=True)

    assert _get_json(client.get("/api/stores")) == ["Aldi", "Rewe"]


def test_get_categories_excludes_null_and_deleted(
    client: FlaskClient, seed_invoice: SeedInvoice
) -> None:
    """Categories are distinct, sorted, and skip NULL and deleted rows."""
    seed_invoice(store="A", category="Food")
    seed_invoice(store="B", category=None)
    seed_invoice(store="C", category="Drinks")
    seed_invoice(store="D", category="Hidden", deleted=True)

    assert _get_json(client.get("/api/categories")) == ["Drinks", "Food"]


# --- POST /api/invoices/import ------------------------------------------------


def test_import_skips_duplicates(
    client: FlaskClient, seed_invoice: SeedInvoice
) -> None:
    """Import inserts new rows and skips ones matching date+store+total."""
    seed_invoice(date="2024-01-01", store="Dup", total=10.0)

    response = client.post(
        "/api/invoices/import",
        json=[
            {"date": "2024-01-01", "store": "Dup", "total": 10.0, "items": []},
            {"date": "2024-01-02", "store": "New", "total": 5.0, "items": []},
        ],
    )

    body = _get_json(response)
    assert response.status_code == 200
    assert body == {"success": True, "imported": 1, "skipped": 1}
    assert len(_get_json(client.get("/api/invoices"))) == 2


def test_import_non_numeric_total_returns_400(client: FlaskClient) -> None:
    """A non-numeric total in any imported invoice is rejected with 400."""
    response = client.post(
        "/api/invoices/import",
        json=[{"date": "2024-01-01", "store": "Bad", "total": "abc"}],
    )
    assert response.status_code == 400
    assert _get_json(response)["error"] == "Field 'total' must be a number"
    assert _get_json(client.get("/api/invoices")) == []


def test_import_non_list_payload_returns_400(client: FlaskClient) -> None:
    """A payload that is not a list of invoices is rejected with 400."""
    response = client.post(
        "/api/invoices/import",
        json={"date": "2024-01-01", "store": "Single", "total": 1.0},
    )
    assert response.status_code == 400
    assert _get_json(response)["error"] == "Expected a list of invoices"
    assert _get_json(client.get("/api/invoices")) == []


# --- PUT /api/invoices/<id> ---------------------------------------------------


def test_update_invoice_replaces_items(
    client: FlaskClient, seed_invoice: SeedInvoice
) -> None:
    """Updating an invoice overwrites its fields and fully replaces its items."""
    invoice_id = seed_invoice(
        store="Old", items=[{"item_name": "old item", "item_price": 1.0}]
    )

    response = client.put(
        f"/api/invoices/{invoice_id}",
        json={
            "date": "2024-04-01",
            "store": "New",
            "category": "Updated",
            "total": 42.0,
            "items": [{"item_name": "new item", "item_price": 5.0}],
        },
    )

    assert response.status_code == 200
    invoice = _get_json(client.get("/api/invoices"))[0]
    assert invoice["store"] == "New"
    assert invoice["total"] == 42.0
    assert [item["item_name"] for item in invoice["items"]] == ["new item"]


def test_update_invoice_missing_field_returns_400(
    client: FlaskClient, seed_invoice: SeedInvoice
) -> None:
    """A missing required key on update is rejected with the JSON 400 envelope."""
    invoice_id = seed_invoice(store="Keep")

    response = client.put(
        f"/api/invoices/{invoice_id}",
        json={"store": "NoDate", "total": 1.0},
    )
    assert response.status_code == 400
    assert _get_json(response)["error"] == "Missing required field: date"


def test_update_nonexistent_invoice_reports_success(client: FlaskClient) -> None:
    """Updating an unknown id succeeds silently and creates nothing."""
    response = client.put(
        "/api/invoices/999999",
        json={"date": "2024-04-01", "store": "Ghost", "total": 1.0, "items": []},
    )
    assert response.status_code == 200
    assert _get_json(response) == {"success": True}
    assert _get_json(client.get("/api/invoices")) == []


# --- DELETE /api/invoices/<id> ------------------------------------------------


def test_delete_invoice_is_soft(client: FlaskClient, seed_invoice: SeedInvoice) -> None:
    """Delete sets deleted_at but keeps the row physically present."""
    invoice_id = seed_invoice(store="ToDelete")

    response = client.delete(f"/api/invoices/{invoice_id}")
    assert response.status_code == 200
    assert _get_json(client.get("/api/invoices")) == []

    conn = db.get_db()
    try:
        row = conn.execute(
            "SELECT deleted_at FROM invoices WHERE id = ?", (invoice_id,)
        ).fetchone()
    finally:
        conn.close()
    assert row is not None
    assert row["deleted_at"] is not None


# --- PUT /api/invoices/bulk-update --------------------------------------------


def test_bulk_update_store_and_category(
    client: FlaskClient, seed_invoice: SeedInvoice
) -> None:
    """Bulk update applies a new store/category to the given ids."""
    first = seed_invoice(store="Old", category="A")
    second = seed_invoice(store="Old", category="A")

    response = client.put(
        "/api/invoices/bulk-update",
        json={"ids": [first, second], "store": "Renamed", "category": "B"},
    )

    body = _get_json(response)
    assert body == {"success": True, "updated": 2}
    stores = {invoice["store"] for invoice in _get_json(client.get("/api/invoices"))}
    assert stores == {"Renamed"}


def test_bulk_update_empty_category_clears_it(
    client: FlaskClient, seed_invoice: SeedInvoice
) -> None:
    """An empty-string category is stored as NULL via strip_text."""
    invoice_id = seed_invoice(store="Shop", category="ToClear")

    client.put(
        "/api/invoices/bulk-update",
        json={"ids": [invoice_id], "category": ""},
    )

    assert _get_json(client.get("/api/invoices"))[0]["category"] is None


def test_bulk_update_missing_ids_returns_400(client: FlaskClient) -> None:
    """An empty ids list is rejected with 400."""
    response = client.put("/api/invoices/bulk-update", json={"ids": [], "store": "X"})
    assert response.status_code == 400
    assert _get_json(response)["error"] == "Missing ids"


def test_bulk_update_missing_fields_returns_400(
    client: FlaskClient, seed_invoice: SeedInvoice
) -> None:
    """Providing neither store nor category is rejected with 400."""
    invoice_id = seed_invoice()
    response = client.put("/api/invoices/bulk-update", json={"ids": [invoice_id]})
    assert response.status_code == 400
    assert _get_json(response)["error"] == "Missing store or category"


# --- POST /api/invoices/bulk-delete -------------------------------------------


def test_bulk_delete_soft_deletes_many(
    client: FlaskClient, seed_invoice: SeedInvoice
) -> None:
    """Bulk delete soft-deletes every given id and reports the count."""
    first = seed_invoice(store="A")
    second = seed_invoice(store="B")

    response = client.post("/api/invoices/bulk-delete", json={"ids": [first, second]})

    assert _get_json(response) == {"success": True, "deleted": 2}
    assert _get_json(client.get("/api/invoices")) == []


def test_bulk_delete_missing_ids_returns_400(client: FlaskClient) -> None:
    """An empty ids list is rejected with 400."""
    response = client.post("/api/invoices/bulk-delete", json={"ids": []})
    assert response.status_code == 400
    assert _get_json(response)["error"] == "Missing ids"
