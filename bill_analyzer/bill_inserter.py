"""
Main logic for inserting bill data into ODS files
"""

# pyright: reportGeneralTypeIssues=false

import datetime as dt
from dataclasses import dataclass
from typing import Any

import dateutil.parser as dparser
from odf import table
from odf.namespaces import TABLENS
from odf.opendocument import load

from .config import COL_DATE, COL_ITEM, COL_PRICE, COL_STORE, COL_TOTAL, ODS_FILE
from .ods_cells import clear_cell_completely, set_cell_value
from .ods_rows import (
    create_blank_separator_row,
    create_item_row,
    restore_row_as_new,
    save_existing_row_data,
)
from .ods_sheets import (
    create_new_date_row,
    find_date_row,
    find_sheet_by_name,
    has_existing_data,
)
from .utils import (
    create_backup,
    extract_price_number,
    is_number,
    parse_date,
    remove_backup,
    restore_from_backup,
)

# Export the duplicate check function for use in other modules
__all__ = ["process_multiple_bills", "insert_bill_into_ods", "check_duplicate_bill"]


@dataclass
class _BillSearchCriteria:
    """Criteria for searching bills during duplicate detection."""

    date_str: str  # ISO format date string (YYYY-MM-DD)
    store_normalized: str  # Lowercase, trimmed store name
    total: float  # Total bill amount
    epsilon: float = 0.01  # Float comparison tolerance


def _has_matching_date(cells: list[table.TableCell], target_date_str: str) -> bool:
    """Check if a row contains a date matching the target date.

    This function extracts the date value from the date column and compares it
    to the target date. It supports both ISO format dates (YYYY-MM-DD) and
    non-ISO formats (e.g., "10.12.25") by attempting ISO parsing first, then
    falling back to flexible date parsing with day-first convention.

    :param cells: Row cells to check
    :param target_date_str: Target date in ISO format (YYYY-MM-DD)
    :return: True if the row's date matches the target date, False otherwise
    """
    if len(cells) <= COL_DATE:
        return False

    # Get date value
    date_value: str = str(cells[COL_DATE])

    # Parse and check date
    if not date_value:
        return False

    row_date: str | None = parse_date(date_value)

    if row_date is None:
        return False

    # Check if date matches
    return row_date == target_date_str


def _get_store_from_bill_start(
    rows: list[table.TableRow], start_idx: int, store: str
) -> int | None:
    """Search for a matching store name within a bill group starting from a given row.

    This function scans rows starting from start_idx to find a row containing
    the specified store name. It stops when it encounters another date entry
    (indicating the start of a different bill) or reaches the end of rows.

    :param rows: All rows in the sheet
    :param start_idx: Row index to start searching from
    :param store: Normalized store name to search for (lowercase, trimmed)
    :return: Row index containing the matching store, or None if not found
    """
    for idx in range(start_idx, len(rows)):
        cells: list[table.TableCell] = rows[idx].getElementsByType(table.TableCell)

        if len(cells) <= COL_STORE:
            continue

        if start_idx != idx and str(cells[COL_DATE]):
            return None

        found_store: str = str(cells[COL_STORE])
        if not found_store:
            continue

        found_store = found_store.strip().lower()
        if found_store != store:
            continue

        return idx

    return None


def _find_total_in_bill_group(
    rows: list[table.TableRow], start_idx: int, total: float
) -> int | None:
    """Search for a matching total price within a bill group starting from a given row.

    This function scans rows starting from start_idx to find a row containing
    the specified total amount. It stops when it encounters a new bill marker
    (date or store entry on a different row) or reaches the end of rows.

    :param rows: All rows in the sheet
    :param start_idx: Row index to start searching from
    :param total: Total price to search for (exact float match)
    :return: Row index containing the matching total, or None if not found or mismatch
    """
    for idx in range(start_idx, len(rows)):
        cells: list[table.TableCell] = rows[idx].getElementsByType(table.TableCell)

        if len(cells) <= COL_TOTAL:
            continue

        if start_idx != idx and (str(cells[COL_DATE]) or str(cells[COL_STORE])):
            return None

        found_total: str = str(cells[COL_TOTAL])
        if not found_total:
            continue

        found_total = extract_price_number(found_total)
        if not is_number(found_total):
            continue

        if float(found_total) != total:
            return None

        return idx

    return None


def _log_duplicate_mismatch(
    idx: int, found_store: str, found_total: float, criteria: _BillSearchCriteria
) -> None:
    """Log details about bills that partially match search criteria.

    :param idx: Row index of the found bill
    :param found_store: Store name found in the row
    :param found_total: Total price found in the bill
    :param criteria: Search criteria containing target values
    """
    store_matches = found_store == criteria.store_normalized
    diff = abs(found_total - criteria.total)
    total_matches = diff < criteria.epsilon

    if store_matches and not total_matches:
        print(
            f"  [Duplicate Check] Row {idx}: Same store but different total "
            f"({found_total}€, diff: {diff:.2f}€)"
        )
    elif not store_matches and total_matches:
        print(
            f"  [Duplicate Check] Row {idx}: Same total but different store "
            f"('{found_store}' vs '{criteria.store_normalized}')"
        )


def _check_bill_match(
    found_store: str, found_total: float, criteria: _BillSearchCriteria
) -> bool:
    """Check if a found bill matches the target bill.

    :param found_store: Store name from the found bill
    :param found_total: Total price from the found bill
    :param criteria: Search criteria containing target values
    :return: True if both store and total match
    """
    store_matches = found_store == criteria.store_normalized
    diff = abs(found_total - criteria.total)
    total_matches = diff < criteria.epsilon
    return store_matches and total_matches


def _process_bill_row_for_duplicate(
    rows: list[table.TableRow],
    idx: int,
    criteria: _BillSearchCriteria,
) -> bool:
    """Process a single row to check if it represents a duplicate bill.

    This function checks if the row at idx contains a bill that matches the
    search criteria (date, store, and total). It uses helper functions to
    search within the bill group for matching store and total values.

    :param rows: All rows in the sheet
    :param idx: Current row index to check
    :param criteria: Search criteria containing target date, store, and total
    :return: True if a matching duplicate bill is found, False otherwise
    """
    cells: list[table.TableCell] = rows[idx].getElementsByType(table.TableCell)

    # Check if this row has the target date
    if not _has_matching_date(cells, criteria.date_str):
        return False

    new_idx: int | None = None

    while new_idx is None:
        new_idx = _get_store_from_bill_start(rows, idx, criteria.store_normalized)
        if new_idx is None:
            return False

        idx = new_idx
        new_idx = _find_total_in_bill_group(rows, idx, criteria.total)
        idx += 1

    return True


def _check_duplicate_bill(
    doc: Any, bill_data: dict[str, Any], verbose: bool = False
) -> bool:
    """Check if a bill with same store, date, and total already exists.

    This function searches for bill entries that span multiple rows.
    Each bill starts with a row containing date and store, followed by
    item rows, with the total appearing in the last row.

    :param doc: ODS document object
    :type doc: Any
    :param bill_data: Bill data dictionary
    :type bill_data: dict[str, Any]
    :param verbose: Whether to print debug information
    :type verbose: bool
    :return: True if duplicate exists, False otherwise
    :rtype: bool
    """
    # Parse date and determine sheet name
    date_parsed: dt.datetime = dparser.parse(bill_data["date"], dayfirst=True)
    sheet_name: str = f"{date_parsed.strftime('%b')} {date_parsed.strftime('%y')}"

    # Find sheet
    target_sheet: table.Table | None = find_sheet_by_name(doc, sheet_name)
    if not target_sheet:
        if verbose:
            print(
                f"  [Duplicate Check] Sheet '{sheet_name}' not found "
                "- not a duplicate"
            )
        return False

    # Prepare search criteria
    criteria: _BillSearchCriteria = _BillSearchCriteria(
        date_str=date_parsed.date().strftime("%Y-%m-%d"),
        store_normalized=bill_data["store"].strip().lower(),
        total=bill_data["total"],
    )

    if verbose:
        print(
            f"  [Duplicate Check] Searching for: "
            f"{bill_data['store']} | {criteria.date_str} | {criteria.total}€"
        )

    # Search for bills on the target date
    rows: list[table.TableRow] = target_sheet.getElementsByType(table.TableRow)
    bills_on_date: list[tuple[int, str, float]] = []

    for idx in range(len(rows)):
        if _process_bill_row_for_duplicate(rows, idx, criteria):
            return True  # Duplicate found

    # No duplicate found
    if verbose:
        if bills_on_date:
            print(
                f"  [Duplicate Check] Found {len(bills_on_date)} bill(s) on {criteria.date_str}, "
                "but none matched both store and total"
            )
        print("  [Duplicate Check] No match found - not a duplicate")
    return False


def _find_target_sheet_and_row(
    doc: Any, bill_data: dict[str, Any], verbose: bool
) -> tuple[table.Table | None, int | None]:
    """Find or create the target sheet and row for bill insertion.

    :param doc: ODS document object
    :type doc: Any
    :param bill_data: Bill data dictionary
    :type bill_data: dict[str, Any]
    :param verbose: Whether to print messages
    :type verbose: bool
    :return: Tuple of (target_sheet, target_row_idx) or (None, None) if sheet not found
    :rtype: tuple[table.Table | None, int | None]
    """
    # Parse date and determine sheet name
    date_parsed: Any = dparser.parse(bill_data["date"], dayfirst=True)
    month: str = date_parsed.strftime("%b")
    year: str = date_parsed.strftime("%y")
    sheet_name: str = f"{month} {year}"

    # Find sheet
    target_sheet: table.Table | None = find_sheet_by_name(doc, sheet_name)
    if not target_sheet:
        if verbose:
            print(f"⚠ Sheet '{sheet_name}' not found - skipping bill")
        return None, None

    # Find or create date row
    target_row_idx: int | None = find_date_row(target_sheet, date_parsed.date())
    if target_row_idx is None:
        if verbose:
            print(f"Creating new row for date {date_parsed.date()}")
        target_row_idx = create_new_date_row(target_sheet, date_parsed.date(), doc)

    return target_sheet, target_row_idx


def _write_first_bill_item(
    cells: list[table.TableCell], bill_data: dict[str, Any]
) -> None:
    """Write the first bill item to the target row.

    :param cells: List of cells in the target row
    :type cells: list[table.TableCell]
    :param bill_data: Bill data dictionary
    :type bill_data: dict[str, Any]
    """
    first_item: dict[str, Any] = bill_data["items"][0]
    set_cell_value(cells[COL_STORE], bill_data["store"])
    set_cell_value(cells[COL_ITEM], first_item["name"])
    set_cell_value(cells[COL_PRICE], first_item["price"])

    # Clear total column
    for col_idx in range(COL_TOTAL, len(cells)):
        clear_cell_completely(cells[col_idx])

    # Add total if only one item
    if len(bill_data["items"]) == 1 and len(cells) > COL_TOTAL:
        set_cell_value(cells[COL_TOTAL], bill_data["total"])


def _insert_remaining_items(
    target_sheet: table.Table,
    cells: list[table.TableCell],
    row_style: str | None,
    bill_data: dict[str, Any],
    reference_row: table.TableRow | None,
) -> None:
    """Insert remaining bill items as new rows.

    :param target_sheet: Target ODS sheet
    :type target_sheet: table.Table
    :param cells: Template cells for styling
    :type cells: list[table.TableCell]
    :param row_style: Row style to apply
    :type row_style: str | None
    :param bill_data: Bill data dictionary
    :type bill_data: dict[str, Any]
    :param reference_row: Row to insert before (or None to append)
    :type reference_row: table.TableRow | None
    """
    remaining_items: list[dict[str, Any]] = bill_data["items"][1:]
    for idx, item in enumerate(remaining_items):
        is_last_item: bool = idx == len(remaining_items) - 1
        total: float | None = bill_data["total"] if is_last_item else None

        new_row: table.TableRow = create_item_row(
            cells, row_style, item["name"], item["price"], total_price=total
        )

        if reference_row is not None:
            target_sheet.insertBefore(new_row, reference_row)
        else:
            target_sheet.addElement(new_row)


def _restore_old_data(
    target_sheet: table.Table,
    old_row_data: list[tuple[Any, str | None]],
    cells: list[table.TableCell],
    row_style: str | None,
    reference_row: table.TableRow | None,
) -> None:
    """Restore old row data that existed before overwriting.

    :param target_sheet: Target ODS sheet
    :type target_sheet: table.Table
    :param old_row_data: Saved old row data (must not be None)
    :type old_row_data: list[tuple[Any, str | None]]
    :param cells: Template cells for styling
    :type cells: list[table.TableCell]
    :param row_style: Row style to apply
    :type row_style: str | None
    :param reference_row: Row to insert before (or None to append)
    :type reference_row: table.TableRow | None
    """
    blank_row: table.TableRow = create_blank_separator_row(cells, row_style)
    if reference_row is not None:
        target_sheet.insertBefore(blank_row, reference_row)
    else:
        target_sheet.addElement(blank_row)

    old_row: table.TableRow = restore_row_as_new(old_row_data, row_style)
    if reference_row is not None:
        target_sheet.insertBefore(old_row, reference_row)
    else:
        target_sheet.addElement(old_row)


def _insert_single_bill_data(
    doc: Any, bill_data: dict[str, Any], verbose: bool = True
) -> None:
    """Insert a single bill's data into an already-loaded ODS document.

    This is an internal function that performs the actual data insertion
    without handling file I/O (loading/saving/backup).

    NOTE: Duplicate checking should be performed BEFORE calling this function
    (e.g., in main.py). This function assumes the bill_data is already validated.

    :param doc: Loaded ODS document object
    :type doc: Any
    :param bill_data: Dictionary containing 'store', 'date', 'items', 'total'
    :type bill_data: dict[str, Any]
    :param verbose: Whether to print progress messages
    :type verbose: bool
    :raises Exception: If sheet is not found or data insertion fails
    """
    # Find target sheet and row
    target_sheet, target_row_idx = _find_target_sheet_and_row(doc, bill_data, verbose)
    if target_sheet is None or target_row_idx is None:
        return

    # Get row and cells
    rows: list[table.TableRow] = target_sheet.getElementsByType(table.TableRow)
    target_row: table.TableRow = rows[target_row_idx]
    cells: list[table.TableCell] = target_row.getElementsByType(table.TableCell)
    row_style: str | None = target_row.getAttrNS(TABLENS, "style-name")

    # Save existing data (will be empty for newly created rows)
    old_data_exists: bool = has_existing_data(cells)
    old_row_data: list[tuple[Any, str | None]] | None = (
        save_existing_row_data(cells) if old_data_exists else None
    )

    # Write first item to the target row
    _write_first_bill_item(cells, bill_data)

    # Calculate reference row for inserting additional rows
    reference_row: table.TableRow | None = (
        rows[target_row_idx + 1] if target_row_idx + 1 < len(rows) else None
    )

    # Insert remaining items as new rows
    _insert_remaining_items(target_sheet, cells, row_style, bill_data, reference_row)

    # Restore old data if it existed
    if old_data_exists and old_row_data is not None:
        _restore_old_data(target_sheet, old_row_data, cells, row_style, reference_row)

    if verbose:
        print(
            f"✓ Inserted {len(bill_data['items'])} items + total for {bill_data['store']}"
        )


def process_multiple_bills(bills_data: list[dict[str, Any]]) -> None:
    """Process multiple bills and insert them into the ODS file in a single transaction.

    This function:
    1. Creates a backup of the ODS file (once)
    2. Loads the document (once)
    3. Sorts bills by date (chronologically)
    4. Inserts all bills into the document
    5. Saves the document (once)
    6. Removes backup on success

    This is more efficient than calling insert_bill_into_ods() multiple times,
    as it only performs file I/O once instead of for each bill.

    :param bills_data: List of bill dictionaries, each containing 'store', 'date', 'items', 'total'
    :type bills_data: list[dict[str, Any]]
    :raises Exception: If any step fails (backup is automatically restored)
    """
    if not bills_data:
        print("No bills to process.")
        return

    # Sort bills by date (chronologically) to ensure correct insertion order
    sorted_bills = sorted(
        bills_data, key=lambda bill: dparser.parse(bill["date"], dayfirst=True).date()
    )

    # Create backup
    backup_path: str = create_backup(ODS_FILE)

    try:
        # Load document once
        print("Loading ODS file...")
        doc = load(ODS_FILE)

        # Insert all bills in chronological order
        for idx, bill_data in enumerate(sorted_bills, 1):
            print(
                f"\n[{idx}/{len(sorted_bills)}] Processing {bill_data.get('store', 'Unknown')}..."
            )
            _insert_single_bill_data(doc, bill_data, verbose=True)

        # Save document once
        print("\nSaving all changes to document...")
        doc.save(ODS_FILE)
        print(f"✓ Successfully saved {len(bills_data)} bill(s) to {ODS_FILE}")

        # Remove backup
        remove_backup(backup_path)

    except Exception as e:
        print(f"✗ Error: {e}")
        restore_from_backup(backup_path, ODS_FILE)
        raise


def insert_bill_into_ods(bill_data: dict[str, Any]) -> None:
    """Insert a single bill into the ODS file.

    This is a convenience wrapper around process_multiple_bills() for
    processing a single bill. For processing multiple bills, use
    process_multiple_bills() directly for better performance.

    :param bill_data: Dictionary containing 'store', 'date', 'items', 'total'
    :type bill_data: dict[str, Any]
    :raises Exception: If any step fails (backup is automatically restored)
    """
    process_multiple_bills([bill_data])


def check_duplicate_bill(bill_data: dict[str, Any], verbose: bool = False) -> bool:
    """Check if a bill is a duplicate by loading and checking the ODS file.

    This is a convenience function that loads the ODS file, checks for duplicates,
    and returns the result. Use this before uploading to external services.

    :param bill_data: Dictionary containing 'store', 'date', 'items', 'total'
    :type bill_data: dict[str, Any]
    :param verbose: Whether to print debug information
    :type verbose: bool
    :return: True if duplicate exists, False otherwise
    :rtype: bool
    :raises Exception: If ODS file cannot be loaded
    """
    doc = load(ODS_FILE)
    return _check_duplicate_bill(doc, bill_data, verbose=verbose)
