"""
ODS sheet operations
"""

# pyright: reportGeneralTypeIssues=false

from datetime import date

import dateutil.parser as dparser
from odf import table
from odf.namespaces import TABLENS
from odf.opendocument import OpenDocument

from .config import COL_DATE, COL_STORE, COL_TOTAL
from .utils import parse_date


def find_last_row(rows: list[table.TableRow]) -> int:
    """Find last row with data."""
    last_row_idx = 0

    for idx, row in enumerate(rows):
        cells: list[table.TableCell] = row.getElementsByType(table.TableCell)
        if len(cells) <= COL_TOTAL:
            continue

        has_data: bool = False
        for col_idx in range(COL_STORE, COL_TOTAL + 1):
            if col_idx < len(cells):
                cell_value: str = str(cells[col_idx])
                if cell_value and cell_value.strip():
                    has_data = True
                    break

        if has_data:
            last_row_idx = idx

    return last_row_idx


def find_sheet_by_name(doc: OpenDocument, sheet_name: str) -> table.Table | None:
    """Find a sheet in an ODS document by name.

    :param doc: ODS document object
    :type doc: OpenDocument
    :param sheet_name: Name of the sheet to find
    :type sheet_name: str
    :return: Sheet object if found, None otherwise
    :rtype: table.Table | None
    """
    sheets = doc.getElementsByType(table.Table)
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
    rows: list[table.TableRow] = sheet.getElementsByType(table.TableRow)

    for idx, row in enumerate(rows):
        cells: list[table.TableCell] = row.getElementsByType(table.TableCell)
        if len(cells) <= COL_DATE:
            continue

        cell_value: str = str(cells[COL_DATE])
        if not cell_value:
            continue

        row_date: str | None = parse_date(cell_value)
        if row_date is None:
            continue

        date_parsed: date = dparser.parse(row_date).date()
        if date_parsed == target_date:
            return idx

    return None


def find_chronological_insertion_point(
    rows: list[table.TableRow], new_date: date
) -> int | None:
    """Find the correct row index to insert a new date chronologically.

    Returns the index of the first row with a date that is AFTER the new_date.
    Returns None if the new date should be inserted at the end.

    :param rows: List of table rows
    :type rows: list[table.TableRow]
    :param new_date: Date to insert
    :type new_date: date
    :return: Row index to insert before, or None to append at end
    :rtype: int | None
    """
    for idx, row in enumerate(rows):
        cells: list[table.TableCell] = row.getElementsByType(table.TableCell)
        if len(cells) <= COL_DATE:
            continue

        cell_value: str = str(cells[COL_DATE])

        # Try to parse as date
        try:
            if cell_value and cell_value.strip():
                cell_date: date = dparser.parse(cell_value, dayfirst=True).date()
                # If we found a date that's after our new date, insert before it
                if cell_date > new_date:
                    return idx
        except (ValueError, TypeError, OverflowError):
            # Skip cells that can't be parsed as dates
            pass

    # No date found that's after new_date, so insert at end
    return None
