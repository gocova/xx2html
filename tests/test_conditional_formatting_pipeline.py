import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

import xx2html.core as core_module
from xx2html import get_xlsx_transform

FIXTURES_DIR = Path(__file__).resolve().parent / "fixtures"

SHEET_HTML = (
    '<section id="{enc_sheet_name}" data-sheet="{sheet_name}">'
    "{table_generated_html}"
    "</section>"
)
SHEETNAME_HTML = '<a class="sheet-nav" href="#{sheet_name}.A1">{sheet_name}</a>'
INDEX_HTML = (
    "<!doctype html><html><head>"
    "{fonts_html}{core_css_html}{user_css_html}{generated_css_html}"
    "{generated_incell_css_html}{conditional_css_html}"
    "</head><body>{sheets_names_generated_html}{sheets_generated_html}{safari_js}</body></html>"
)


def _run_transform(source_file: Path, apply_cf: bool) -> tuple[bool, str | None]:
    transform = get_xlsx_transform(
        sheet_html=SHEET_HTML,
        sheetname_html=SHEETNAME_HTML,
        index_html=INDEX_HTML,
        fonts_html="",
        core_css="",
        user_css="",
        safari_js="",
        apply_cf=apply_cf,
    )
    with tempfile.TemporaryDirectory() as tmp_dir:
        output_file = Path(tmp_dir) / "output.html"
        return transform(str(source_file), str(output_file), "en_US")


class ConditionalFormattingPipelineTests(unittest.TestCase):
    def test_apply_cf_true_calls_condif2css_processor(self):
        source_file = FIXTURES_DIR / "merged_cells_cf.xlsx"

        with patch(
            "xx2html.core.process_conditional_formatting",
            wraps=core_module.process_conditional_formatting,
        ) as process_mock:
            ok, err = _run_transform(source_file, apply_cf=True)

        self.assertTrue(ok, err)
        self.assertGreaterEqual(process_mock.call_count, 1)
        args, kwargs = process_mock.call_args
        self.assertEqual("Data", args[0].title)
        self.assertTrue(kwargs.get("fail_ok", False))

    def test_apply_cf_false_skips_condif2css_processor(self):
        source_file = FIXTURES_DIR / "merged_cells_cf.xlsx"

        with patch(
            "xx2html.core.process_conditional_formatting",
            wraps=core_module.process_conditional_formatting,
        ) as process_mock:
            ok, err = _run_transform(source_file, apply_cf=False)

        self.assertTrue(ok, err)
        self.assertEqual(0, process_mock.call_count)


if __name__ == "__main__":
    unittest.main()
