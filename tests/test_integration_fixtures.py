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


def _render_fixture(source_file: Path, apply_cf: bool) -> str:
    transform = create_xlsx_transform(
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
        ok, err = transform(str(source_file), str(output_file), "en_US")
        if not ok:
            raise AssertionError(f"Transform failed for {source_file.name}: {err}")
        return output_file.read_text(encoding="utf-8")


class FixturesIntegrationTests(unittest.TestCase):
    def test_merged_cells_and_conditional_formatting_fixture(self):
        html = _render_fixture(FIXTURES_DIR / "merged_cells_cf.xlsx", apply_cf=True)

        self.assertIn('colspan="2"', html)
        self.assertIn('rowspan="2"', html)
        self.assertIn("/*conditional formatting*/", html)
        self.assertIn("Data!C1", html)

    def test_incell_image_fixture(self):
        html = _render_fixture(FIXTURES_DIR / "incell_image.xlsx", apply_cf=False)

        self.assertIn("vm-richvaluerel_rid1", html)
        self.assertIn("incell-image", html)
        self.assertIn("content:url(\"data:image/png;base64", html)
        self.assertIn("object-fit: contain;", html)
        self.assertIn('loading="lazy"', html)
        self.assertIn('decoding="async"', html)


if __name__ == "__main__":
    unittest.main()
