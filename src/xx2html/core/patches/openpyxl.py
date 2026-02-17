# Copyright (c) 2024-2026 gonzalo covarrubias <gocova.dev+xx2html@gmail.com>

# Monkey patch for: WorkSheetParser
from openpyxl.worksheet._reader import WorkSheetParser

# Monkey patch for: WorksheetReader
from openpyxl.worksheet._reader import WorksheetReader

## required imports
from warnings import warn
from openpyxl.utils import coordinate_to_tuple
from openpyxl.utils.datetime import from_excel, from_ISO8601
from openpyxl.worksheet._reader import (
    parse_richtext_string,
    VALUE_TAG, FORMULA_TAG, INLINE_STRING,
    _cast_number, # type: ignore
)
from openpyxl.cell.text import Text

from xx2html.core.types import CovaCell


def cova_parse_cell(self, element):
    """
    Parse a cell element into a dictionary containing the cell's row, column, value, data type, style id, and vm id.

    :param element: The cell element to parse
    :return: A dictionary containing the cell's row, column, value, data type, style id, and vm id
    """
    data_type = element.get("t", "n")
    coordinate = element.get("r")
    style_id = element.get("s", 0)
    vm_id = element.get("vm", None)

    if style_id:
        style_id = int(style_id)

    # if vm_id:
    #     vm_id = int(vm_id)

    if data_type == "inlineStr":
        value = None
    else:
        value = element.findtext(VALUE_TAG, None) or None

    if coordinate:
        row, column = coordinate_to_tuple(coordinate)
        self.col_counter = column
    else:
        self.col_counter += 1
        row, column = self.row_counter, self.col_counter

    if not self.data_only and element.find(FORMULA_TAG) is not None:
        data_type = "f"
        value = self.parse_formula(element)

    elif value is not None:
        if data_type == "n":
            value = _cast_number(value)
            if style_id in self.date_formats:
                data_type = "d"
                try:
                    value = from_excel(
                        value, self.epoch, timedelta=style_id in self.timedelta_formats
                    )
                except (OverflowError, ValueError):
                    msg = f"""Cell {coordinate} is marked as a date but the serial value {value} is outside the limits for dates. The cell will be treated as an error."""
                    warn(msg)
                    data_type = "e"
                    value = "#VALUE!"
        elif data_type == "s":
            value = self.shared_strings[int(value)]
        elif data_type == "b":
            value = bool(int(value))
        elif data_type == "str":
            data_type = "s"
        elif data_type == "d":
            value = from_ISO8601(value)

    elif data_type == "inlineStr":
        child = element.find(INLINE_STRING)
        if child is not None:
            data_type = "s"
            if self.rich_text:
                value = parse_richtext_string(child)
            else:
                text_child = Text.from_tree(child)
                value = text_child.content if text_child else ""

    return {
        "row": row,
        "column": column,
        "value": value,
        "data_type": data_type,
        "style_id": style_id,
        "vm_id": vm_id,
    }


def cova__bind_cells(self):
    """
    Bind CovaCells to the worksheet.

    This method is used to process the parsed data and bind it to the worksheet.
    It creates a CovaCell for each cell in the parsed data and sets the value and data_type.
    The CovaCells are then stored in the worksheet's _cells dictionary.

    """
    for idx, row in self.parser.parse():
        for cell in row:
            style = self.ws.parent._cell_styles[cell["style_id"]]
            c = CovaCell(
                self.ws,
                row=cell["row"],
                column=cell["column"],
                style_array=style,
                vm_id=cell["vm_id"],
            )
            c._value = cell["value"]
            c.data_type = cell["data_type"]

            self.ws._cells[(cell["row"], cell["column"])] = c

    if self.ws._cells:
        self.ws._current_row = self.ws.max_row  # use cells not row dimensions



def apply_patches():
    WorkSheetParser.parse_cell = cova_parse_cell
    WorksheetReader.bind_cells = cova__bind_cells
