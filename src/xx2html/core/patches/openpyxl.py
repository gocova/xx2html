# Copyright (c) 2024-2026 gonzalo covarrubias <gocova.dev+xx2html@gmail.com>

"""Monkey patches required to expose rich-value metadata in openpyxl cells."""

import os
import re
from typing import Any, cast

# Monkey patch for: WorkSheetParser
from openpyxl.worksheet._reader import WorkSheetParser

# Monkey patch for: WorksheetReader
from openpyxl.worksheet._reader import WorksheetReader
from openpyxl import __version__ as OPENPYXL_VERSION

## required imports
from warnings import warn
from openpyxl.utils import coordinate_to_tuple
from openpyxl.utils.datetime import from_excel, from_ISO8601
from openpyxl.worksheet._reader import (
    parse_richtext_string,
    VALUE_TAG, FORMULA_TAG, INLINE_STRING,
)
from openpyxl.cell.text import Text

from xx2html.core.types import CovaCell

SUPPORTED_OPENPYXL_MAJOR_MINORS = {(3, 1)}
ALLOW_UNSUPPORTED_OPENPYXL_ENV = "XX2HTML_ALLOW_UNSUPPORTED_OPENPYXL"


def _get_openpyxl_major_minor(version_value: str) -> tuple[int, int] | None:
    version_match = re.match(r"^(\d+)\.(\d+)", version_value)
    if version_match is None:
        return None
    return int(version_match.group(1)), int(version_match.group(2))


def ensure_openpyxl_compatibility() -> None:
    """Validate that the current openpyxl version is in the tested range."""
    parsed_version = _get_openpyxl_major_minor(OPENPYXL_VERSION)
    if parsed_version is None:
        warn(
            "Unable to parse openpyxl version '%s'; continuing without a strict "
            "compatibility guarantee." % OPENPYXL_VERSION
        )
        return

    if parsed_version in SUPPORTED_OPENPYXL_MAJOR_MINORS:
        return

    compatibility_error = (
        "Unsupported openpyxl version '%s'. Tested major.minor versions: %s. "
        "Set %s=1 to bypass this check at your own risk."
        % (
            OPENPYXL_VERSION,
            ", ".join(
                f"{major}.{minor}"
                for major, minor in sorted(SUPPORTED_OPENPYXL_MAJOR_MINORS)
            ),
            ALLOW_UNSUPPORTED_OPENPYXL_ENV,
        )
    )
    if os.getenv(ALLOW_UNSUPPORTED_OPENPYXL_ENV) == "1":
        warn(compatibility_error)
        return

    raise RuntimeError(compatibility_error)


def _cova_cast_number(value: object) -> int | float:
    if isinstance(value, bool):
        raise TypeError("Boolean is not a numeric cell payload.")
    if isinstance(value, (int, float)):
        return value
    if not isinstance(value, str):
        raise TypeError(f"Unsupported numeric payload type: {type(value)!r}")

    raw_value = value.strip()
    if raw_value == "":
        raise ValueError("Empty numeric payload")

    if "." in raw_value or "e" in raw_value.lower():
        return float(raw_value)
    return int(raw_value)


def cova_parse_cell(self, element) -> dict[str, object]:
    """
    Parse a cell element into a dictionary containing the cell's row, column, value, data type, style id, and vm id.

    :param element: The cell element to parse
    :return: A dictionary containing the cell's row, column, value, data type, style id, and vm id
    """
    data_type = element.get("t", "n")
    coordinate = element.get("r")
    raw_style_id = element.get("s", 0)
    style_id = 0
    vm_id = element.get("vm", None)

    try:
        if raw_style_id not in (None, ""):
            style_id = int(raw_style_id)
    except (TypeError, ValueError):
        warn(
            f"Cell {coordinate} has an invalid style id {raw_style_id!r}. Falling back to style 0."
        )
        style_id = 0

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
            try:
                value = _cova_cast_number(value)
            except (TypeError, ValueError):
                warn(
                    f"Cell {coordinate} has an invalid numeric value {value!r}. The cell will be treated as an error."
                )
                data_type = "e"
                value = "#VALUE!"
            if data_type == "n" and style_id in self.date_formats:
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
            try:
                value = self.shared_strings[int(value)]
            except (TypeError, ValueError, IndexError):
                warn(
                    f"Cell {coordinate} references an invalid shared string index {value!r}. The cell will be treated as an error."
                )
                data_type = "e"
                value = "#VALUE!"
        elif data_type == "b":
            try:
                value = bool(int(value))
            except (TypeError, ValueError):
                warn(
                    f"Cell {coordinate} has an invalid boolean value {value!r}. The cell will be treated as an error."
                )
                data_type = "e"
                value = "#VALUE!"
        elif data_type == "str":
            data_type = "s"
        elif data_type == "d":
            try:
                value = from_ISO8601(value)
            except (TypeError, ValueError):
                warn(
                    f"Cell {coordinate} has an invalid ISO8601 date value {value!r}. The cell will be treated as an error."
                )
                data_type = "e"
                value = "#VALUE!"

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


def cova_bind_cells(self) -> None:
    """
    Bind CovaCells to the worksheet.

    This method is used to process the parsed data and bind it to the worksheet.
    It creates a CovaCell for each cell in the parsed data and sets the value and data_type.
    The CovaCells are then stored in the worksheet's _cells dictionary.

    """
    for _, row in self.parser.parse():
        for cell in row:
            style_id = cell.get("style_id", 0)
            cell_styles = self.ws.parent._cell_styles
            if (
                not isinstance(style_id, int)
                or style_id < 0
                or style_id >= len(cell_styles)
            ):
                warn(
                    f"Cell {cell.get('row')}:{cell.get('column')} has an out-of-range style id {style_id!r}. Falling back to style 0."
                )
                style_id = 0
            style = cell_styles[style_id]
            cova_cell = CovaCell(
                self.ws,
                row=cell["row"],
                column=cell["column"],
                style_array=style,
                vm_id=cell["vm_id"],
            )
            setattr(cova_cell, "_value", cell["value"])
            cova_cell.data_type = cell["data_type"]

            self.ws._cells[(cell["row"], cell["column"])] = cova_cell

    if self.ws._cells:
        self.ws._current_row = self.ws.max_row  # use cells not row dimensions



def apply_patches() -> None:
    """Install parser/reader monkey patches used by `xx2html`."""
    ensure_openpyxl_compatibility()
    setattr(cast(Any, WorkSheetParser), "parse_cell", cova_parse_cell)
    setattr(cast(Any, WorksheetReader), "bind_cells", cova_bind_cells)
