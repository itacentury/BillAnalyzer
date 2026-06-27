"""Shared pytest fixtures providing an isolated, temp-backed Flask test client.

Importing :mod:`summa` runs ``create_app()`` (and therefore ``init_db()``) at
module import time. To keep that import-time side effect from creating a file, a
throwaway in-memory ``DATABASE_PATH`` is set *before* ``summa`` is first imported
here. Each test then gets its own fresh on-disk database via the :func:`client`
fixture.
"""

import os
from collections.abc import Callable
from pathlib import Path
from typing import Any

import pytest

# Redirect the import-time database to an in-memory SQLite database before
# importing summa, so the import-time create_app()/init_db() leaves no file
# behind. Each test gets its own on-disk DB via the client fixture's monkeypatch.
os.environ.setdefault("DATABASE_PATH", ":memory:")

from flask.testing import FlaskClient  # noqa: E402

from summa import (
    create_app,  # noqa: E402
    db,  # noqa: E402
)

SeedInvoice = Callable[..., int]


@pytest.fixture
def client(tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> FlaskClient:
    """Return a Flask test client backed by a fresh, per-test SQLite database."""
    db_path: Path = tmp_path / "test.db"
    # get_db() reads this module global on every call, so patching it redirects
    # every connection (including the one init_db() opens) to the temp database.
    monkeypatch.setattr(db, "DATABASE", str(db_path))
    app = create_app()
    app.config["TESTING"] = True
    # Keep producing real 500 responses for uncaught exceptions (as gunicorn does)
    # instead of letting TESTING re-raise them into the test.
    app.config["PROPAGATE_EXCEPTIONS"] = False
    return app.test_client()


@pytest.fixture
def seed_invoice(client: FlaskClient) -> SeedInvoice:
    """Return a helper that inserts an invoice (and items) straight into the DB."""

    def _seed(
        date: str = "2024-01-15",
        store: str = "Test Store",
        category: str | None = None,
        total: float = 10.0,
        items: list[dict[str, Any]] | None = None,
        deleted: bool = False,
    ) -> int:
        conn = db.get_db()
        cursor = conn.cursor()
        cursor.execute(
            "INSERT INTO invoices (date, store, category, total) VALUES (?, ?, ?, ?)",
            (date, store, category, total),
        )
        invoice_id: int | None = cursor.lastrowid
        assert invoice_id is not None  # AUTOINCREMENT always yields a rowid
        for item in items or []:
            cursor.execute(
                "INSERT INTO invoice_items (invoice_id, item_name, item_price) "
                "VALUES (?, ?, ?)",
                (invoice_id, item["item_name"], item["item_price"]),
            )
        if deleted:
            cursor.execute(
                "UPDATE invoices SET deleted_at = CURRENT_TIMESTAMP WHERE id = ?",
                (invoice_id,),
            )
        conn.commit()
        conn.close()
        return invoice_id

    return _seed
