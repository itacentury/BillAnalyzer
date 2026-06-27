"""REST API routes for invoice CRUD, bulk operations, stores and categories."""

import logging
import sqlite3
from typing import Any

from flask import Blueprint, Response, jsonify, request

from summa.db import get_db
from summa.helpers import ApiResponse, strip_text

logger: logging.Logger = logging.getLogger(__name__)

invoices_bp: Blueprint = Blueprint("invoices", __name__)


@invoices_bp.route("/api/invoices", methods=["GET"])
def get_invoices() -> Response:
    """Retrieve all invoices with optional filtering and sorting."""
    conn: sqlite3.Connection = get_db()
    cursor: sqlite3.Cursor = conn.cursor()

    # Get filter parameters
    filters: dict[str, str] = {
        "search": request.args.get("search", ""),
        "store": request.args.get("store", ""),
        "category": request.args.get("category", ""),
        "date_from": request.args.get("date_from", ""),
        "date_to": request.args.get("date_to", ""),
        "sort_by": request.args.get("sort_by", "date"),
        "sort_order": request.args.get("sort_order", "desc"),
    }

    # Build query - exclude soft-deleted invoices
    query: str = "SELECT * FROM invoices WHERE deleted_at IS NULL"
    params: list[str] = []

    if filters["search"]:
        query += (
            " AND (store LIKE ? OR id IN "
            "(SELECT invoice_id FROM invoice_items WHERE item_name LIKE ?))"
        )
        params.extend([f"%{filters['search']}%", f"%{filters['search']}%"])

    if filters["store"]:
        query += " AND store = ?"
        params.append(filters["store"])

    if filters["category"]:
        query += " AND category = ?"
        params.append(filters["category"])

    if filters["date_from"]:
        query += " AND date >= ?"
        params.append(filters["date_from"])

    if filters["date_to"]:
        query += " AND date <= ?"
        params.append(filters["date_to"])

    # Sorting
    if filters["sort_by"] in ["date", "store", "total"]:
        order: str = "DESC" if filters["sort_order"] == "desc" else "ASC"
        query += f" ORDER BY {filters['sort_by']} {order}"

    cursor.execute(query, params)

    result: list[dict[str, Any]] = []
    for invoice in cursor.fetchall():
        cursor.execute(
            "SELECT * FROM invoice_items WHERE invoice_id = ?", (invoice["id"],)
        )
        result.append(
            {
                "id": invoice["id"],
                "date": invoice["date"],
                "store": invoice["store"],
                "category": invoice["category"],
                "total": invoice["total"],
                "items": [
                    {"item_name": item["item_name"], "item_price": item["item_price"]}
                    for item in cursor.fetchall()
                ],
            }
        )

    conn.close()
    return jsonify(result)


@invoices_bp.route("/api/stores", methods=["GET"])
def get_stores() -> Response:
    """Return a list of all unique store names."""
    conn: sqlite3.Connection = get_db()
    cursor: sqlite3.Cursor = conn.cursor()
    cursor.execute(
        "SELECT DISTINCT store FROM invoices WHERE deleted_at IS NULL ORDER BY store"
    )
    stores: list[str] = [row["store"] for row in cursor.fetchall()]
    conn.close()
    return jsonify(stores)


@invoices_bp.route("/api/categories", methods=["GET"])
def get_categories() -> Response:
    """Return a list of all unique invoice categories."""
    conn: sqlite3.Connection = get_db()
    cursor: sqlite3.Cursor = conn.cursor()
    cursor.execute(
        "SELECT DISTINCT category FROM invoices "
        "WHERE deleted_at IS NULL AND category IS NOT NULL ORDER BY category"
    )
    categories: list[str] = [row["category"] for row in cursor.fetchall()]
    conn.close()
    return jsonify(categories)


@invoices_bp.route("/api/invoices", methods=["POST"])
def add_invoice() -> Response:
    """Create a new invoice with its associated items."""
    data: Any = request.json
    conn: sqlite3.Connection = get_db()
    cursor: sqlite3.Cursor = conn.cursor()

    cursor.execute(
        "INSERT INTO invoices (date, store, category, total) VALUES (?, ?, ?, ?)",
        (
            strip_text(data["date"]),
            strip_text(data["store"]),
            strip_text(data.get("category")),
            float(data["total"]),
        ),
    )
    invoice_id: int | None = cursor.lastrowid

    for item in data.get("items", []):
        cursor.execute(
            "INSERT INTO invoice_items (invoice_id, item_name, item_price) VALUES (?, ?, ?)",
            (invoice_id, strip_text(item["item_name"]), float(item["item_price"])),
        )

    conn.commit()
    conn.close()
    item_count: int = len(data.get("items", []))
    logger.info(
        "Invoice created: id=%s, store='%s', total=%.2f, items=%d",
        invoice_id,
        data.get("store"),
        float(data["total"]),
        item_count,
    )
    return jsonify({"success": True, "id": invoice_id})


@invoices_bp.route("/api/invoices/import", methods=["POST"])
def import_invoices() -> ApiResponse:
    """Bulk import invoices, skipping duplicates based on date, store, and total."""
    data: Any = request.json
    conn: sqlite3.Connection = get_db()
    cursor: sqlite3.Cursor = conn.cursor()

    imported_count: int = 0
    skipped_count: int = 0

    try:
        for invoice_data in data:
            store: str | None = strip_text(invoice_data["store"])
            date: str | None = strip_text(invoice_data["date"])
            category: str | None = strip_text(invoice_data.get("category"))
            total: float = float(invoice_data["total"])

            # Duplicate check: same combination of date, store and total amount
            cursor.execute(
                "SELECT id FROM invoices WHERE date = ? AND store = ? AND total = ?",
                (date, store, total),
            )
            existing: Any = cursor.fetchone()

            if existing:
                skipped_count += 1
                continue

            cursor.execute(
                "INSERT INTO invoices (date, store, category, total) VALUES (?, ?, ?, ?)",
                (date, store, category, total),
            )
            invoice_id: int | None = cursor.lastrowid

            for item in invoice_data.get("items", []):
                cursor.execute(
                    "INSERT INTO invoice_items "
                    "(invoice_id, item_name, item_price) VALUES (?, ?, ?)",
                    (
                        invoice_id,
                        strip_text(item["item_name"]),
                        float(item["item_price"]),
                    ),
                )
            imported_count += 1

        conn.commit()
        logger.info(
            "Import completed: imported=%d, skipped=%d (of %d total)",
            imported_count,
            skipped_count,
            len(data),
        )
        return jsonify(
            {"success": True, "imported": imported_count, "skipped": skipped_count}
        )
    except sqlite3.Error as e:
        conn.rollback()
        logger.error("Import failed: %s", e)
        return jsonify({"success": False, "error": str(e)}), 500
    finally:
        conn.close()


@invoices_bp.route("/api/invoices/<int:invoice_id>", methods=["PUT"])
def update_invoice(invoice_id: int) -> ApiResponse:
    """Update an existing invoice and replace all its items."""
    data: Any = request.json
    conn: sqlite3.Connection = get_db()
    cursor: sqlite3.Cursor = conn.cursor()

    try:
        # Update invoice
        cursor.execute(
            "UPDATE invoices SET date = ?, store = ?, category = ?, total = ? WHERE id = ?",
            (
                strip_text(data["date"]),
                strip_text(data["store"]),
                strip_text(data.get("category")),
                float(data["total"]),
                invoice_id,
            ),
        )

        # Delete existing items
        cursor.execute("DELETE FROM invoice_items WHERE invoice_id = ?", (invoice_id,))

        # Insert new items
        for item in data.get("items", []):
            cursor.execute(
                "INSERT INTO invoice_items (invoice_id, item_name, item_price) VALUES (?, ?, ?)",
                (invoice_id, strip_text(item["item_name"]), float(item["item_price"])),
            )

        conn.commit()
        item_count: int = len(data.get("items", []))
        logger.info(
            "Invoice updated: id=%d, store='%s', total=%.2f, items=%d",
            invoice_id,
            data.get("store"),
            float(data["total"]),
            item_count,
        )
        return jsonify({"success": True})
    except sqlite3.Error as e:
        conn.rollback()
        logger.error("Failed to update invoice id=%d: %s", invoice_id, e)
        return jsonify({"success": False, "error": str(e)}), 500
    finally:
        conn.close()


@invoices_bp.route("/api/invoices/<int:invoice_id>", methods=["DELETE"])
def delete_invoice(invoice_id: int) -> Response:
    """Soft-delete an invoice by setting its deleted_at timestamp."""
    conn: sqlite3.Connection = get_db()
    cursor: sqlite3.Cursor = conn.cursor()
    # Soft delete: set deleted_at timestamp instead of removing from database
    cursor.execute(
        "UPDATE invoices SET deleted_at = CURRENT_TIMESTAMP WHERE id = ?",
        (invoice_id,),
    )
    conn.commit()
    conn.close()
    logger.info("Invoice soft-deleted: id=%d", invoice_id)
    return jsonify({"success": True})


@invoices_bp.route("/api/invoices/bulk-update", methods=["PUT"])
def bulk_update_invoices() -> ApiResponse:
    """Update store name and/or category for multiple invoices at once."""
    data: Any = request.json
    invoice_ids: list[int] = data.get("ids", [])
    new_store: str | None = strip_text(data.get("store"))
    new_category: str | None = data.get("category")

    if not invoice_ids:
        return jsonify({"success": False, "error": "Missing ids"}), 400

    if not new_store and new_category is None:
        return jsonify({"success": False, "error": "Missing store or category"}), 400

    conn: sqlite3.Connection = get_db()
    cursor: sqlite3.Cursor = conn.cursor()

    try:
        placeholders: str = ",".join("?" * len(invoice_ids))
        set_clauses: list[str] = []
        params: list[str | int | None] = []

        if new_store:
            set_clauses.append("store = ?")
            params.append(new_store)

        if new_category is not None:
            set_clauses.append("category = ?")
            # Empty string means remove category (set to NULL)
            params.append(strip_text(new_category))

        params.extend(invoice_ids)
        cursor.execute(
            f"UPDATE invoices SET {', '.join(set_clauses)} WHERE id IN ({placeholders})",
            params,
        )
        updated_count: int = cursor.rowcount
        conn.commit()
        logger.info(
            "Bulk update completed: %d invoices updated (ids=%s)",
            updated_count,
            invoice_ids,
        )
        return jsonify({"success": True, "updated": updated_count})
    except sqlite3.Error as e:
        conn.rollback()
        logger.error("Bulk update failed for ids=%s: %s", invoice_ids, e)
        return jsonify({"success": False, "error": str(e)}), 500
    finally:
        conn.close()


@invoices_bp.route("/api/invoices/bulk-delete", methods=["POST"])
def bulk_delete_invoices() -> ApiResponse:
    """Soft-delete multiple invoices at once."""
    data: Any = request.json
    invoice_ids: list[int] = data.get("ids", [])

    if not invoice_ids:
        return jsonify({"success": False, "error": "Missing ids"}), 400

    conn: sqlite3.Connection = get_db()
    cursor: sqlite3.Cursor = conn.cursor()

    try:
        placeholders: str = ",".join("?" * len(invoice_ids))
        # Soft delete: set deleted_at timestamp instead of removing from database
        cursor.execute(
            f"UPDATE invoices SET deleted_at = CURRENT_TIMESTAMP WHERE id IN ({placeholders})",
            invoice_ids,
        )
        deleted_count: int = cursor.rowcount
        conn.commit()
        logger.info(
            "Bulk soft-delete completed: %d invoices deleted (ids=%s)",
            deleted_count,
            invoice_ids,
        )
        return jsonify({"success": True, "deleted": deleted_count})
    except sqlite3.Error as e:
        conn.rollback()
        logger.error("Bulk delete failed for ids=%s: %s", invoice_ids, e)
        return jsonify({"success": False, "error": str(e)}), 500
    finally:
        conn.close()
