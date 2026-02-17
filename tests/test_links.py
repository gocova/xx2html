import unittest

from xx2html.core.links import update_links


class UpdateLinksTests(unittest.TestCase):
    def test_skips_anchors_without_href_and_rewrites_links(self):
        html = (
            '<div>'
            '<a><span>No href</span></a>'
            '<a href="#Sheet 1.A1"><span>Go local</span></a>'
            '<a href="https://example.com"><strong>Go external</strong></a>'
            '</div>'
        )

        output = update_links(html, {"Sheet 1": "sheet_001"})

        self.assertIn('about:srcdoc#sheet_001', output)
        self.assertIn('href="#sheet_001"', output)
        self.assertIn('class="xlsx_sheet-link sharepoint_visible"', output)
        self.assertIn('class="xlsx_sheet-link js_visible"', output)
        self.assertIn('<span>Go local</span>', output)

        self.assertIn('href="https://example.com"', output)
        self.assertIn('target="_blank"', output)
        self.assertIn('rel="noopener noreferrer"', output)
        self.assertIn('<strong>Go external</strong>', output)

    def test_rewrites_local_links_for_sheet_names_containing_dots(self):
        html = '<a href="#Q1.2026.A1">Go to dotted sheet</a>'

        output = update_links(html, {"Q1.2026": "sheet_abc"})

        self.assertIn('about:srcdoc#sheet_abc', output)
        self.assertIn('href="#sheet_abc"', output)
        self.assertNotIn('about:srcdoc"', output)

    def test_rewrites_local_links_for_dotted_sheet_name_without_cell_suffix(self):
        html = '<a href="#Q1.2026">Go to dotted sheet root</a>'

        output = update_links(html, {"Q1.2026": "sheet_abc"})

        self.assertIn('about:srcdoc#sheet_abc', output)
        self.assertIn('href="#sheet_abc"', output)

    def test_rewrites_excel_style_local_reference(self):
        html = '<a href="#\'Q1.2026 Sheet\'!A1">Excel style reference</a>'

        output = update_links(html, {"Q1.2026 Sheet": "sheet_q1"})

        self.assertIn('about:srcdoc#sheet_q1', output)
        self.assertIn('href="#sheet_q1"', output)

    def test_rewrites_local_links_without_cell_reference_suffix(self):
        html = '<a href="#Dashboard">Open dashboard</a>'

        output = update_links(html, {"Dashboard": "sheet_dash"})

        self.assertIn('about:srcdoc#sheet_dash', output)
        self.assertIn('href="#sheet_dash"', output)

    def test_keeps_original_local_link_when_sheet_mapping_is_missing(self):
        html = '<a href="#MissingSheet.A1">Broken ref</a>'

        output = update_links(html, {"Sheet 1": "sheet_001"})

        self.assertIn('href="#MissingSheet.A1"', output)
        self.assertNotIn("about:srcdoc", output)


if __name__ == "__main__":
    unittest.main()
