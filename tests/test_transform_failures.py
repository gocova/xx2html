import tempfile
import unittest
from pathlib import Path

from xx2html import create_xlsx_transform


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


def _build_transform():
    return create_xlsx_transform(
        sheet_html=SHEET_HTML,
        sheetname_html=SHEETNAME_HTML,
        index_html=INDEX_HTML,
        fonts_html="",
        core_css="",
        user_css="",
        safari_js="",
        apply_cf=False,
    )


class TransformFailureTests(unittest.TestCase):
    def test_missing_source_file_returns_false_and_error(self):
        transform = _build_transform()
        missing_source = "__missing__/missing.xlsx"

        with tempfile.TemporaryDirectory() as tmp_dir:
            output_file = Path(tmp_dir) / "out.html"
            ok, err = transform(missing_source, str(output_file), "en_US")

        self.assertFalse(ok)
        self.assertIsInstance(err, str)
        self.assertIn("FileNotFoundError", err)

    def test_destination_directory_returns_false_and_error(self):
        transform = _build_transform()
        source_file = FIXTURES_DIR / "merged_cells_cf.xlsx"

        with tempfile.TemporaryDirectory() as tmp_dir:
            ok, err = transform(str(source_file), tmp_dir, "en_US")

        self.assertFalse(ok)
        self.assertIsInstance(err, str)
        self.assertIn("IsADirectoryError", err)


if __name__ == "__main__":
    unittest.main()
