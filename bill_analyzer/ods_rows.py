"""
ODS row operations
"""

# pyright: reportGeneralTypeIssues=false

from odf import table, text
from odf.namespaces import TABLENS

from .config import COL_ITEM, COL_PRICE, COL_STORE, COL_TOTAL
from .ods_cells import set_cell_value


def create_item_row(  # pylint: disable=too-many-arguments
    template_cells: list[table.TableCell],
    template_row_style: str | None,
    item_name: str,
    item_price: float,
    *,
    store_name: str | None = None,
    total_price: float | None = None,
) -> table.TableRow:
    """Create a new row for a bill item with proper formatting.

    :param template_cells: List of cells to copy styles from
    :type template_cells: list[table.TableCell]
    :param template_row_style: Row style to apply
    :type template_row_style: str | None
    :param item_name: Name of the item
    :type item_name: str
    :param item_price: Price of the item
    :type item_price: float
    :param store_name: Store name (only for first row)
    :type store_name: str | None
    :param total_price: Total price (only for last row)
    :type total_price: float | None
    :return: New table row with all cells properly formatted
    :rtype: table.TableRow
    """
    new_row = table.TableRow()

    # Set row style
    if template_row_style:
        new_row.setAttrNS(TABLENS, "style-name", template_row_style)

    # Create cells for each column
    for col_idx, template_cell in enumerate(template_cells):
        new_cell = table.TableCell()

        # Copy cell style
        cell_style = template_cell.getAttrNS(TABLENS, "style-name")
        if cell_style:
            new_cell.setAttrNS(TABLENS, "style-name", cell_style)

        # Set cell content based on column
        if col_idx == COL_STORE and store_name:
            set_cell_value(new_cell, store_name)
        elif col_idx == COL_ITEM:
            set_cell_value(new_cell, item_name)
        elif col_idx == COL_PRICE:
            set_cell_value(new_cell, item_price)
        elif col_idx == COL_TOTAL and total_price:
            set_cell_value(new_cell, total_price)
        else:
            new_cell.appendChild(text.P(text=""))

        new_row.appendChild(new_cell)

    return new_row
