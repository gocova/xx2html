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


if __name__ == "__main__":
    unittest.main()
