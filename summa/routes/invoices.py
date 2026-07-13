"""REST API routes for invoice CRUD, bulk operations, stores and categories."""

import logging
import sqlite3
from typing import Any, Final

from flask import Blueprint, Response, jsonify, request

from summa.db import chunked, db_cursor, insert_invoice_items, placeholders_for
from summa.helpers import (
    ApiResponse,
    Invoice,
    ValidationError,
    error_response,
    escape_like,
    parse_bounded_int,
    parse_invoice,
    parse_invoice_list,
    strip_text,
)

logger: logging.Logger = logging.getLogger(__name__)

invoices_bp: Blueprint = Blueprint("invoices", __name__)

DEFAULT_PAGE_SIZE: Final[int] = 50
MAX_PAGE_SIZE: Final[int] = 200


def _build_invoice_filter(args: Any) -> tuple[str, list[str]]:
    """Build the shared WHERE clause and params for invoice list filtering.

    Excludes soft-deleted invoices. Kept separate from ORDER BY/LIMIT so the same
    clause and params can drive the list, count/sum and id-list queries alike.
    """
    where: str = "WHERE deleted_at IS NULL"
    params: list[str] = []

    search: str = args.get("search", "")
    if search:
        where += (
            " AND (store LIKE ? ESCAPE '\\' OR id IN "
            "(SELECT invoice_id FROM invoice_items WHERE item_name LIKE ? ESCAPE '\\'))"
        )
        escaped: str = escape_like(search)
        params.extend([f"%{escaped}%", f"%{escaped}%"])

    store: str = args.get("store", "")
    if store:
        where += " AND store = ?"
        params.append(store)

    category: str = args.get("category", "")
    if category:
        where += " AND category = ?"
        params.append(category)

    date_from: str = args.get("date_from", "")
    if date_from:
        where += " AND date >= ?"
        params.append(date_from)

    date_to: str = args.get("date_to", "")
    if date_to:
        where += " AND date <= ?"
        params.append(date_to)

    return where, params


@invoices_bp.route("/api/invoices", methods=["GET"])
def get_invoices() -> Response:
    """Retrieve a page of invoices with optional filtering and sorting."""
    where, params = _build_invoice_filter(request.args)

    # Sorting. Always include `id` as a unique tie-breaker so rows sharing a
    # sort value keep a stable relative order across LIMIT/OFFSET page
    # boundaries (otherwise paging can skip or duplicate rows).
    sort_by: str = request.args.get("sort_by", "date")
    if sort_by in ["date", "store", "total"]:
        direction: str = (
            "DESC" if request.args.get("sort_order", "desc") == "desc" else "ASC"
        )
        order: str = f" ORDER BY {sort_by} {direction}, id DESC"
    else:
        order = " ORDER BY id DESC"

    # Pagination
    page: int = parse_bounded_int(request.args.get("page"), 1, 1, 1_000_000)
    page_size: int = parse_bounded_int(
        request.args.get("page_size"), DEFAULT_PAGE_SIZE, 1, MAX_PAGE_SIZE
    )
    offset: int = (page - 1) * page_size

    with db_cursor() as cursor:
        cursor.execute(
            f"SELECT COUNT(*) AS total_count, COALESCE(SUM(total), 0) AS total_sum "
            f"FROM invoices {where}",
            params,
        )
        totals: sqlite3.Row = cursor.fetchone()
        total_count: int = totals["total_count"]
        total_sum: float = totals["total_sum"]

        cursor.execute(
            f"SELECT * FROM invoices {where}{order} LIMIT ? OFFSET ?",
            [*params, page_size, offset],
        )
        invoices: list[sqlite3.Row] = cursor.fetchall()
        items_by_invoice: dict[int, list[dict[str, Any]]] = _fetch_items_by_invoice(
            cursor, [invoice["id"] for invoice in invoices]
        )

    result: list[dict[str, Any]] = [
        {
            "id": invoice["id"],
            "date": invoice["date"],
            "store": invoice["store"],
            "category": invoice["category"],
            "total": invoice["total"],
            "items": items_by_invoice.get(invoice["id"], []),
        }
        for invoice in invoices
    ]
    return jsonify(
        {
            "invoices": result,
            "page": page,
            "page_size": page_size,
            "total_count": total_count,
            "total_sum": total_sum,
        }
    )


@invoices_bp.route("/api/invoices/ids", methods=["GET"])
def get_invoice_ids() -> Response:
    """Return the ids of every invoice matching the current filters.

    Backs cross-page "select all": the client needs the full filtered id set,
    which the paginated list endpoint does not expose. Reuses the same filter
    clause so both endpoints always agree on what "matching" means.
    """
    where, params = _build_invoice_filter(request.args)
    with db_cursor() as cursor:
        cursor.execute(f"SELECT id FROM invoices {where} ORDER BY id", params)
        ids: list[int] = [row["id"] for row in cursor.fetchall()]
    return jsonify({"ids": ids})


def _fetch_items_by_invoice(
    cursor: sqlite3.Cursor, invoice_ids: list[int]
) -> dict[int, list[dict[str, Any]]]:
    """Fetch all items for the given invoices, grouped by invoice id."""
    items_by_invoice: dict[int, list[dict[str, Any]]] = {}
    for chunk in chunked(invoice_ids):
        cursor.execute(
            f"SELECT * FROM invoice_items WHERE invoice_id IN ({placeholders_for(len(chunk))})",
            chunk,
        )
        for item in cursor.fetchall():
            items_by_invoice.setdefault(item["invoice_id"], []).append(
                {"item_name": item["item_name"], "item_price": item["item_price"]}
            )
    return items_by_invoice


@invoices_bp.route("/api/stores", methods=["GET"])
def get_stores() -> Response:
    """Return a list of all unique store names."""
    with db_cursor() as cursor:
        cursor.execute(
            "SELECT DISTINCT store FROM invoices WHERE deleted_at IS NULL ORDER BY store"
        )
        stores: list[str] = [row["store"] for row in cursor.fetchall()]
    return jsonify(stores)


@invoices_bp.route("/api/categories", methods=["GET"])
def get_categories() -> Response:
    """Return a list of all unique invoice categories."""
    with db_cursor() as cursor:
        cursor.execute(
            "SELECT DISTINCT category FROM invoices "
            "WHERE deleted_at IS NULL AND category IS NOT NULL ORDER BY category"
        )
        categories: list[str] = [row["category"] for row in cursor.fetchall()]
    return jsonify(categories)


@invoices_bp.route("/api/invoices", methods=["POST"])
def add_invoice() -> ApiResponse:
    """Create a new invoice with its associated items."""
    data: Any = request.json
    try:
        invoice: Invoice = parse_invoice(data)
    except ValidationError as e:
        return error_response(str(e), 400)

    try:
        with db_cursor() as cursor:
            cursor.execute(
                "INSERT INTO invoices (date, store, category, total) VALUES (?, ?, ?, ?)",
                (invoice.date, invoice.store, invoice.category, invoice.total),
            )
            invoice_id: int | None = cursor.lastrowid
            insert_invoice_items(cursor, invoice_id, invoice.items)
        logger.info(
            "Invoice created: id=%s, store='%s', total=%.2f, items=%d",
            invoice_id,
            invoice.store,
            invoice.total,
            len(invoice.items),
        )
        return jsonify({"success": True, "id": invoice_id})
    except sqlite3.Error as e:
        logger.error("Failed to create invoice: %s", e)
        return error_response(str(e), 500)


@invoices_bp.route("/api/invoices/import", methods=["POST"])
def import_invoices() -> ApiResponse:
    """Bulk import invoices, skipping duplicates based on date, store, and total."""
    data: Any = request.json
    try:
        invoices: list[Invoice] = parse_invoice_list(data)
    except ValidationError as e:
        return error_response(str(e), 400)

    imported_count: int = 0
    skipped_count: int = 0

    try:
        with db_cursor() as cursor:
            for invoice in invoices:
                # Duplicate check: same combination of date, store and total amount
                cursor.execute(
                    "SELECT id FROM invoices WHERE date = ? AND store = ? AND total = ?",
                    (invoice.date, invoice.store, invoice.total),
                )
                existing: Any = cursor.fetchone()

                if existing:
                    skipped_count += 1
                    continue

                cursor.execute(
                    "INSERT INTO invoices (date, store, category, total) VALUES (?, ?, ?, ?)",
                    (invoice.date, invoice.store, invoice.category, invoice.total),
                )
                insert_invoice_items(cursor, cursor.lastrowid, invoice.items)
                imported_count += 1

        logger.info(
            "Import completed: imported=%d, skipped=%d (of %d total)",
            imported_count,
            skipped_count,
            len(invoices),
        )
        return jsonify(
            {"success": True, "imported": imported_count, "skipped": skipped_count}
        )
    except sqlite3.Error as e:
        logger.error("Import failed: %s", e)
        return error_response(str(e), 500)


@invoices_bp.route("/api/invoices/<int:invoice_id>", methods=["PUT"])
def update_invoice(invoice_id: int) -> ApiResponse:
    """Update an existing invoice and replace all its items."""
    data: Any = request.json
    try:
        invoice: Invoice = parse_invoice(data)
    except ValidationError as e:
        return error_response(str(e), 400)

    try:
        with db_cursor() as cursor:
            cursor.execute(
                "UPDATE invoices SET date = ?, store = ?, category = ?, total = ? WHERE id = ?",
                (
                    invoice.date,
                    invoice.store,
                    invoice.category,
                    invoice.total,
                    invoice_id,
                ),
            )
            # Replace all items: remove the old ones, then insert the new set
            cursor.execute(
                "DELETE FROM invoice_items WHERE invoice_id = ?", (invoice_id,)
            )
            insert_invoice_items(cursor, invoice_id, invoice.items)
        logger.info(
            "Invoice updated: id=%d, store='%s', total=%.2f, items=%d",
            invoice_id,
            invoice.store,
            invoice.total,
            len(invoice.items),
        )
        return jsonify({"success": True})
    except sqlite3.Error as e:
        logger.error("Failed to update invoice id=%d: %s", invoice_id, e)
        return error_response(str(e), 500)


@invoices_bp.route("/api/invoices/<int:invoice_id>", methods=["DELETE"])
def delete_invoice(invoice_id: int) -> ApiResponse:
    """Soft-delete an invoice by setting its deleted_at timestamp."""
    try:
        with db_cursor() as cursor:
            # Soft delete: set deleted_at timestamp instead of removing from database
            cursor.execute(
                "UPDATE invoices SET deleted_at = CURRENT_TIMESTAMP WHERE id = ?",
                (invoice_id,),
            )
        logger.info("Invoice soft-deleted: id=%d", invoice_id)
        return jsonify({"success": True})
    except sqlite3.Error as e:
        logger.error("Failed to delete invoice id=%d: %s", invoice_id, e)
        return error_response(str(e), 500)


@invoices_bp.route("/api/invoices/bulk-update", methods=["PUT"])
def bulk_update_invoices() -> ApiResponse:
    """Update store name and/or category for multiple invoices at once."""
    data: Any = request.json
    invoice_ids: list[int] = data.get("ids", [])
    new_store: str | None = strip_text(data.get("store"))
    new_category: str | None = data.get("category")

    if not invoice_ids:
        return error_response("Missing ids", 400)

    if not new_store and new_category is None:
        return error_response("Missing store or category", 400)

    set_clauses: list[str] = []
    params: list[str | int | None] = []

    if new_store:
        set_clauses.append("store = ?")
        params.append(new_store)

    if new_category is not None:
        set_clauses.append("category = ?")
        # Empty string means remove category (set to NULL)
        params.append(strip_text(new_category))

    try:
        with db_cursor() as cursor:
            updated_count: int = 0
            for chunk in chunked(invoice_ids):
                cursor.execute(
                    f"UPDATE invoices SET {', '.join(set_clauses)} "
                    f"WHERE id IN ({placeholders_for(len(chunk))})",
                    [*params, *chunk],
                )
                updated_count += cursor.rowcount
        logger.info(
            "Bulk update completed: %d invoices updated (ids=%s)",
            updated_count,
            invoice_ids,
        )
        return jsonify({"success": True, "updated": updated_count})
    except sqlite3.Error as e:
        logger.error("Bulk update failed for ids=%s: %s", invoice_ids, e)
        return error_response(str(e), 500)


@invoices_bp.route("/api/invoices/bulk-delete", methods=["POST"])
def bulk_delete_invoices() -> ApiResponse:
    """Soft-delete multiple invoices at once."""
    data: Any = request.json
    invoice_ids: list[int] = data.get("ids", [])

    if not invoice_ids:
        return error_response("Missing ids", 400)

    try:
        with db_cursor() as cursor:
            # Soft delete: set deleted_at timestamp instead of removing from database
            deleted_count: int = 0
            for chunk in chunked(invoice_ids):
                cursor.execute(
                    "UPDATE invoices SET deleted_at = CURRENT_TIMESTAMP "
                    f"WHERE id IN ({placeholders_for(len(chunk))})",
                    chunk,
                )
                deleted_count += cursor.rowcount
        logger.info(
            "Bulk soft-delete completed: %d invoices deleted (ids=%s)",
            deleted_count,
            invoice_ids,
        )
        return jsonify({"success": True, "deleted": deleted_count})
    except sqlite3.Error as e:
        logger.error("Bulk delete failed for ids=%s: %s", invoice_ids, e)
        return error_response(str(e), 500)
