"""
ODS sheet operations
"""

from datetime import date
from typing import Any

import dateutil.parser as dparser
from odf import table, text
from odf.namespaces import TABLENS

from .config import COL_DATE, COL_STORE, COL_TOTAL
from .ods_cells import get_cell_value, set_cell_value


def _find_last_data_row_and_template(
    rows: list[table.TableRow],
) -> tuple[int | None, list[table.TableCell] | None, int]:
    """Find the last row with data and determine cell count.

    :param rows: List of table rows
    :type rows: list[table.TableRow]
    :return: Tuple of (last_data_row_idx, template_cells, num_cells_to_create)
    :rtype: tuple[int | None, list[table.TableCell] | None, int]
    """
    last_data_row_idx: int | None = None
    template_cells: list[table.TableCell] | None = None
    num_cells_to_create: int = COL_TOTAL + 1  # Default: at least 6 cells

    for idx, row in enumerate(rows):
        cells: list[table.TableCell] = row.getElementsByType(table.TableCell)
        if len(cells) <= COL_TOTAL:
            continue

        # Update number of cells to create based on existing rows
        # Limit to 10 to avoid issues with repeated columns
        if len(cells) > num_cells_to_create and len(cells) <= 10:
            num_cells_to_create = len(cells)

        # Check if this row has data in any column
        has_data: bool = False
        for col_idx in range(COL_STORE, COL_TOTAL + 1):
            if col_idx < len(cells):
                cell_value: Any = get_cell_value(cells[col_idx])
                if cell_value and str(cell_value).strip():
                    has_data = True
                    break

        # If this row has data, remember it as last data row and save cell styles
        if has_data:
            last_data_row_idx = idx
            template_cells = cells

    return last_data_row_idx, template_cells, num_cells_to_create


def _create_row_with_cells(
    num_cells: int,
    template_cells: list[table.TableCell] | None,
    date_col_idx: int | None = None,
    new_date: date | None = None,
    doc: Any | None = None,
) -> table.TableRow:
    """Create a row with specified number of cells.

    :param num_cells: Number of cells to create
    :type num_cells: int
    :param template_cells: Template cells for styling (or None)
    :type template_cells: list[table.TableCell] | None
    :param date_col_idx: Column index for date (or None for blank row)
    :type date_col_idx: int | None
    :param new_date: Date to insert (required if date_col_idx is set)
    :type new_date: date | None
    :param doc: Document for date styling (required if date_col_idx is set)
    :type doc: Any | None
    :return: New table row
    :rtype: table.TableRow
    """
    new_row: table.TableRow = table.TableRow()

    for col_idx in range(num_cells):
        new_cell: table.TableCell = table.TableCell()

        # Copy cell style from template if available (but NOT for date column)
        if template_cells and col_idx < len(template_cells) and col_idx != date_col_idx:
            cell_style: str | None = template_cells[col_idx].getAttrNS(
                TABLENS, "style-name"
            )
            if cell_style:
                new_cell.setAttrNS(TABLENS, "style-name", cell_style)

        # Set cell content
        if col_idx == date_col_idx and new_date is not None and doc is not None:
            set_cell_value(new_cell, new_date, doc)
        else:
            new_cell.appendChild(text.P(text=""))

        new_row.appendChild(new_cell)

    return new_row


def _insert_row(
    sheet: table.Table, row: table.TableRow, reference_row: table.TableRow | None
) -> None:
    """Insert a row into the sheet.

    :param sheet: Target sheet
    :type sheet: table.Table
    :param row: Row to insert
    :type row: table.TableRow
    :param reference_row: Row to insert before (or None to append)
    :type reference_row: table.TableRow | None
    """
    if reference_row is not None:
        sheet.insertBefore(row, reference_row)
    else:
        sheet.addElement(row)


def _find_new_row_index(sheet: table.Table, target_date_iso: str) -> int:
    """Find the index of a newly created date row.

    :param sheet: Sheet containing the row
    :type sheet: table.Table
    :param target_date_iso: ISO format date string
    :type target_date_iso: str
    :return: Row index
    :rtype: int
    """
    updated_rows: list[table.TableRow] = sheet.getElementsByType(table.TableRow)

    for idx, row in enumerate(updated_rows):
        cells: list[table.TableCell] = row.getElementsByType(table.TableCell)
        if len(cells) > COL_DATE:
            cell_value: Any = get_cell_value(cells[COL_DATE])
            if isinstance(cell_value, str) and cell_value == target_date_iso:
                return idx

    # Fallback: return the row before last (should be our new row)
    return len(updated_rows) - 1


def find_sheet_by_name(doc: Any, sheet_name: str) -> table.Table | None:
    """Find a sheet in an ODS document by name.

    :param doc: ODS document object
    :type doc: Any
    :param sheet_name: Name of the sheet to find
    :type sheet_name: str
    :return: Sheet object if found, None otherwise
    :rtype: table.Table | None
    """
    sheets = doc.spreadsheet.getElementsByType(table.Table)
    for sheet in sheets:
        if sheet.getAttrNS(TABLENS, "name") == sheet_name:
            return sheet
    return None


def find_date_row(sheet: table.Table, target_date: date) -> int | None:
    """Find the row index containing a specific date.

    Handles both:
    - New format: date-value attribute with ISO format (YYYY-MM-DD)
    - Old format: text content with German format (DD.MM.YY)

    :param sheet: ODS sheet to search
    :type sheet: table.Table
    :param target_date: Date to find
    :type target_date: date
    :return: Row index if found, None otherwise
    :rtype: int | None
    """
    rows = sheet.getElementsByType(table.TableRow)

    for idx, row in enumerate(rows):
        cells = row.getElementsByType(table.TableCell)
        if len(cells) <= COL_DATE:
            continue

        cell_value = get_cell_value(cells[COL_DATE])

        # Try to parse as date (handles both ISO and German formats)
        try:
            if isinstance(cell_value, str) and cell_value.strip():
                cell_date = dparser.parse(cell_value, dayfirst=True).date()
                if cell_date == target_date:
                    return idx
        except (ValueError, TypeError, OverflowError):
            # Skip cells that can't be parsed as dates
            pass

    return None


def create_new_date_row(sheet: table.Table, new_date: date, doc: Any) -> int:
    """Create a new row after the last entry with the given date.

    Inserts a blank separator row before the new date row.

    :param sheet: ODS sheet to add row to
    :type sheet: table.Table
    :param new_date: Date to insert
    :type new_date: date
    :param doc: ODS document object (needed for date style)
    :type doc: Any
    :return: Index of the newly created date row
    :rtype: int
    """
    rows: list[table.TableRow] = sheet.getElementsByType(table.TableRow)

    # Find the last row with actual data and determine number of cells needed
    last_data_row_idx, template_cells, num_cells_to_create = (
        _find_last_data_row_and_template(rows)
    )

    # Determine insertion point (after last data row)
    insert_after_idx: int = (
        last_data_row_idx if last_data_row_idx is not None else len(rows) - 1
    )
    reference_row: table.TableRow | None = (
        rows[insert_after_idx + 1] if insert_after_idx + 1 < len(rows) else None
    )

    # Create and insert blank separator row
    blank_row: table.TableRow = _create_row_with_cells(
        num_cells_to_create, template_cells
    )
    _insert_row(sheet, blank_row, reference_row)

    # Create and insert date row
    date_row: table.TableRow = _create_row_with_cells(
        num_cells_to_create, template_cells, COL_DATE, new_date, doc
    )
    _insert_row(sheet, date_row, reference_row)

    # Find and return the index of the new date row
    target_date_iso: str = new_date.strftime("%Y-%m-%d")
    return _find_new_row_index(sheet, target_date_iso)


def has_existing_data(cells: list[table.TableCell]) -> bool:
    """Check if a row has existing data in store/item columns.

    :param cells: List of cells to check
    :type cells: list[table.TableCell]
    :return: True if data exists, False otherwise
    :rtype: bool
    """
    for col_idx in range(COL_STORE, COL_TOTAL):
        if col_idx >= len(cells):
            continue
        cell_value = get_cell_value(cells[col_idx])
        if cell_value and str(cell_value).strip():
            return True
    return False
