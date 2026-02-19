import unittest

from condif2css.css import CssBuilder, CssRulesRegistry
from openpyxl import Workbook

from xx2html.core.types import CovaCell
from xx2html.core.utils import (
    CELL_HEIGHT__DEFAULT,
    COL_WIDTH__DEFAULT,
    cova_render_table,
    get_worksheet_contents,
)


class UtilsTests(unittest.TestCase):
    def test_cova_render_table_sorts_cell_classes_deterministically(self):
        html = cova_render_table(
            {
                "rows": [
                    [
                        {
                            "attrs": {"id": "Sheet1!A1"},
                            "column": 1,
                            "row": 1,
                            "value": "X",
                            "formatted_value": "X",
                            "style": {},
                            "classes": {"zeta", "alpha", "beta"},
                            "vm_id": None,
                        }
                    ]
                ],
                "cols": [
                    {
                        "attrs": {},
                        "index": "A",
                        "width": 65,
                        "style": {"visibility": "visible"},
                        "hidden": False,
                        "collapsed": False,
                    }
                ],
                "images": {},
                "vm_ids": set(),
                "vm_ids_dimension_references": {},
                "vm_cell_vm_ids": {},
                "table_width": 65,
            }
        )

        self.assertIn('class="alpha beta zeta"', html)

    def test_cova_render_table_incell_placeholder_has_safe_attributes(self):
        html = cova_render_table(
            {
                "rows": [
                    [
                        {
                            "attrs": {"id": "Sheet1!A1"},
                            "column": 1,
                            "row": 1,
                            "value": "X",
                            "formatted_value": "",
                            "style": {},
                            "classes": {"incell-image"},
                            "vm_id": "1",
                        }
                    ]
                ],
                "cols": [
                    {
                        "attrs": {},
                        "index": "A",
                        "width": 65,
                        "style": {"visibility": "visible"},
                        "hidden": False,
                        "collapsed": False,
                    }
                ],
                "images": {},
                "vm_ids": {"1"},
                "vm_ids_dimension_references": {},
                "vm_cell_vm_ids": {},
                "table_width": 65,
            }
        )

        self.assertIn('<img alt="" loading="lazy" decoding="async" />', html)

    def test_get_worksheet_contents_vm_dimensions_fall_back_when_effective_size_is_zero(self):
        workbook = Workbook()
        worksheet = workbook.active
        worksheet.title = "Data"
        worksheet["A1"] = "Image"
        worksheet.merge_cells("A1:B2")

        worksheet.column_dimensions["A"].hidden = True
        worksheet.column_dimensions["A"].min = 1
        worksheet.column_dimensions["A"].max = 1
        worksheet.column_dimensions["B"].hidden = True
        worksheet.column_dimensions["B"].min = 2
        worksheet.column_dimensions["B"].max = 2
        worksheet.row_dimensions[1].height = 0
        worksheet.row_dimensions[2].hidden = True

        style = worksheet.parent._cell_styles[worksheet["A1"].style_id]
        vm_cell = CovaCell(worksheet, row=1, column=1, style_array=style, vm_id="1")
        vm_cell._value = "Image"
        vm_cell.data_type = "s"
        worksheet._cells[(1, 1)] = vm_cell

        css_builder = CssBuilder(lambda _color: None)
        css_rules_registry = CssRulesRegistry()

        contents = get_worksheet_contents(
            worksheet,
            css_rules_registry=css_rules_registry,
            css_builder=css_builder,
            get_css_from_cell=lambda _cell, _merged: set(),
            locale="en_US",
            ws_index=0,
        )

        self.assertEqual({"1"}, contents["vm_ids"])
        self.assertEqual(
            {"width": COL_WIDTH__DEFAULT, "height": CELL_HEIGHT__DEFAULT},
            contents["vm_ids_dimension_references"]["cell_0_0_0"],
        )


if __name__ == "__main__":
    unittest.main()
