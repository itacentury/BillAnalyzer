"""Tests for schema creation, migrations and connection setup in :mod:`summa.db`."""

import sqlite3
from pathlib import Path

import pytest

from summa import db


def _columns(conn: sqlite3.Connection, table: str) -> list[str]:
    """Return the column names of ``table`` via PRAGMA table_info."""
    cursor = conn.execute(f"PRAGMA table_info({table})")
    return [row[1] for row in cursor.fetchall()]


@pytest.fixture
def temp_db(tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> Path:
    """Point summa.db at a fresh, empty database file for the duration of a test."""
    db_path: Path = tmp_path / "schema.db"
    monkeypatch.setattr(db, "DATABASE", str(db_path))
    return db_path


def test_init_db_creates_tables(temp_db: Path) -> None:
    """init_db creates both tables with the expected columns."""
    db.init_db()

    conn = db.get_db()
    try:
        invoice_columns: list[str] = _columns(conn, "invoices")
        item_columns: list[str] = _columns(conn, "invoice_items")
    finally:
        conn.close()

    assert {
        "id",
        "date",
        "store",
        "category",
        "total",
        "created_at",
        "deleted_at",
    } <= set(invoice_columns)
    assert {"id", "invoice_id", "item_name", "item_price"} <= set(item_columns)


def test_init_db_is_idempotent(temp_db: Path) -> None:
    """Running init_db repeatedly does not raise and keeps the schema stable."""
    db.init_db()
    db.init_db()

    conn = db.get_db()
    try:
        columns: list[str] = _columns(conn, "invoices")
    finally:
        conn.close()

    assert "deleted_at" in columns
    assert "category" in columns


def test_init_db_migrates_legacy_table(temp_db: Path) -> None:
    """An old invoices table without deleted_at/category gets both columns added."""
    conn = db.get_db()
    conn.execute(
        """
        CREATE TABLE invoices (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            date TEXT NOT NULL,
            store TEXT NOT NULL,
            total REAL NOT NULL,
            created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
        )
        """
    )
    conn.commit()
    conn.close()

    db.init_db()

    conn = db.get_db()
    try:
        columns: list[str] = _columns(conn, "invoices")
    finally:
        conn.close()

    assert "deleted_at" in columns
    assert "category" in columns


def test_get_db_uses_row_factory(temp_db: Path) -> None:
    """get_db returns rows that support mapping-style access by column name."""
    db.init_db()

    conn = db.get_db()
    try:
        conn.execute(
            "INSERT INTO invoices (date, store, total) VALUES (?, ?, ?)",
            ("2024-01-01", "Shop", 9.99),
        )
        conn.commit()
        row = conn.execute("SELECT store, total FROM invoices").fetchone()
    finally:
        conn.close()

    assert row["store"] == "Shop"
    assert row["total"] == 9.99


def test_get_db_enables_wal(temp_db: Path) -> None:
    """get_db enables WAL journal mode on the connection."""
    db.init_db()

    conn = db.get_db()
    try:
        mode = conn.execute("PRAGMA journal_mode").fetchone()[0]
    finally:
        conn.close()

    assert mode.lower() == "wal"
