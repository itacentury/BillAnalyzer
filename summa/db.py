"""Database connection management and schema initialization."""

import logging
import os
import sqlite3
from typing import Final

logger: logging.Logger = logging.getLogger(__name__)

DATABASE: Final[str] = os.environ.get("DATABASE_PATH", "invoices.db")


def get_db() -> sqlite3.Connection:
    """Create and return a database connection with WAL mode enabled."""
    conn: sqlite3.Connection = sqlite3.connect(DATABASE, timeout=30.0)
    conn.row_factory = sqlite3.Row
    # Enable WAL mode for better concurrency
    conn.execute("PRAGMA journal_mode=WAL")
    return conn


def init_db() -> None:
    """Initialize the database schema and apply migrations if needed."""
    conn: sqlite3.Connection = get_db()
    cursor: sqlite3.Cursor = conn.cursor()

    cursor.execute(
        """
        CREATE TABLE IF NOT EXISTS invoices (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            date TEXT NOT NULL,
            store TEXT NOT NULL,
            category TEXT DEFAULT NULL,
            total REAL NOT NULL,
            created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
            deleted_at TIMESTAMP DEFAULT NULL
        )
    """
    )

    cursor.execute(
        """
        CREATE TABLE IF NOT EXISTS invoice_items (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            invoice_id INTEGER NOT NULL,
            item_name TEXT NOT NULL,
            item_price REAL NOT NULL,
            FOREIGN KEY (invoice_id) REFERENCES invoices (id) ON DELETE CASCADE
        )
    """
    )

    # Migration: Add deleted_at column if it doesn't exist (for existing databases)
    cursor.execute("PRAGMA table_info(invoices)")
    columns: list[str] = [column[1] for column in cursor.fetchall()]
    if "deleted_at" not in columns:
        try:
            cursor.execute(
                "ALTER TABLE invoices ADD COLUMN deleted_at TIMESTAMP DEFAULT NULL"
            )
            logger.info("Migration applied: added 'deleted_at' column")
        except sqlite3.OperationalError:
            logger.debug("Column 'deleted_at' already exists, skipping migration")

    if "category" not in columns:
        try:
            cursor.execute("ALTER TABLE invoices ADD COLUMN category TEXT DEFAULT NULL")
            logger.info("Migration applied: added 'category' column")
        except sqlite3.OperationalError:
            logger.debug("Column 'category' already exists, skipping migration")

    conn.commit()
    conn.close()
    logger.info("Database initialized successfully")
