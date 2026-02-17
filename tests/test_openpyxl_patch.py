import unittest
import warnings
from types import SimpleNamespace
from xml.etree.ElementTree import Element, SubElement
from unittest.mock import patch

from openpyxl import Workbook

import xx2html.core.patches.openpyxl as patch_module
from xx2html.core.patches.openpyxl import apply_patches, cova_bind_cells, cova_parse_cell
from xx2html.core.types import CovaCell


def _make_parser(
    *,
    data_only: bool = True,
    rich_text: bool = False,
    date_formats: set[int] | None = None,
):
    return SimpleNamespace(
        data_only=data_only,
        rich_text=rich_text,
        col_counter=0,
        row_counter=1,
        date_formats=date_formats or set(),
        timedelta_formats=set(),
        epoch=None,
        shared_strings=["zero", "one", "two"],
        parse_formula=lambda _element: "FORMULA_RESULT",
    )


def _make_cell_element(
    *,
    data_type: str = "n",
    coordinate: str = "A1",
    style_id: str = "0",
    value: str | None = None,
    include_formula: bool = False,
    include_inline: bool = False,
) -> Element:
    element = Element("c", {"t": data_type, "r": coordinate, "s": style_id})
    if value is not None:
        value_tag = SubElement(element, patch_module.VALUE_TAG)
        value_tag.text = value
    if include_formula:
        SubElement(element, patch_module.FORMULA_TAG)
    if include_inline:
        SubElement(element, patch_module.INLINE_STRING)
    return element


class CovaParseCellTests(unittest.TestCase):
    def test_formula_cells_use_parse_formula_when_data_only_is_false(self):
        parser = _make_parser(data_only=False)
        cell_element = _make_cell_element(
            data_type="n", value="5", include_formula=True
        )

        parsed = cova_parse_cell(parser, cell_element)

        self.assertEqual("f", parsed["data_type"])
        self.assertEqual("FORMULA_RESULT", parsed["value"])

    def test_date_overflow_marks_cell_as_error_and_warns(self):
        parser = _make_parser(date_formats={1})
        cell_element = _make_cell_element(data_type="n", style_id="1", value="42")

        with patch(
            "xx2html.core.patches.openpyxl.from_excel", side_effect=OverflowError
        ):
            with warnings.catch_warnings(record=True) as warning_records:
                warnings.simplefilter("always")
                parsed = cova_parse_cell(parser, cell_element)

        self.assertEqual("e", parsed["data_type"])
        self.assertEqual("#VALUE!", parsed["value"])
        self.assertTrue(
            any("outside the limits for dates" in str(w.message) for w in warning_records)
        )

    def test_inline_rich_text_uses_parse_richtext_string(self):
        parser = _make_parser(rich_text=True)
        cell_element = _make_cell_element(data_type="inlineStr", include_inline=True)

        with patch(
            "xx2html.core.patches.openpyxl.parse_richtext_string",
            return_value="rich-text-value",
        ) as rich_text_mock:
            parsed = cova_parse_cell(parser, cell_element)

        rich_text_mock.assert_called_once()
        self.assertEqual("s", parsed["data_type"])
        self.assertEqual("rich-text-value", parsed["value"])

    def test_inline_plain_text_uses_text_from_tree_content(self):
        parser = _make_parser(rich_text=False)
        cell_element = _make_cell_element(data_type="inlineStr", include_inline=True)
        text_obj = SimpleNamespace(content="plain-text")

        with patch(
            "xx2html.core.patches.openpyxl.Text.from_tree", return_value=text_obj
        ) as text_mock:
            parsed = cova_parse_cell(parser, cell_element)

        text_mock.assert_called_once()
        self.assertEqual("s", parsed["data_type"])
        self.assertEqual("plain-text", parsed["value"])


class CovaBindCellsTests(unittest.TestCase):
    def test_bind_cells_creates_covacell_and_updates_current_row(self):
        workbook = Workbook()
        worksheet = workbook.active
        parser = SimpleNamespace(
            parse=lambda: iter(
                [
                    (
                        1,
                        [
                            {
                                "style_id": 0,
                                "row": 1,
                                "column": 1,
                                "value": "hello",
                                "data_type": "s",
                                "vm_id": "9",
                            }
                        ],
                    )
                ]
            )
        )
        reader = SimpleNamespace(parser=parser, ws=worksheet)

        cova_bind_cells(reader)

        bound_cell = worksheet._cells[(1, 1)]
        self.assertIsInstance(bound_cell, CovaCell)
        self.assertEqual("hello", bound_cell._value)
        self.assertEqual("s", bound_cell.data_type)
        self.assertEqual("9", bound_cell._vm_id)
        self.assertEqual(1, worksheet._current_row)

    def test_apply_patches_registers_cova_hooks(self):
        apply_patches()
        self.assertIs(patch_module.WorkSheetParser.parse_cell, patch_module.cova_parse_cell)
        self.assertIs(
            patch_module.WorksheetReader.bind_cells, patch_module.cova_bind_cells
        )


if __name__ == "__main__":
    unittest.main()
