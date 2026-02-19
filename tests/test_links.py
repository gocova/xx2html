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

    def test_resolves_excel_style_reference_with_escaped_apostrophes(self):
        html = '<a href="#\'Bob\'\'s Sheet\'!A1">Excel escaped apostrophe</a>'

        output = update_links(html, {"Bob's Sheet": "sheet_bob"})

        self.assertIn('about:srcdoc#sheet_bob', output)
        self.assertIn('href="#sheet_bob"', output)

    def test_flags_can_disable_local_and_external_rewrites(self):
        html = (
            '<div>'
            '<a href="#Sheet 1.A1">Local</a>'
            '<a href="https://example.com">External</a>'
            "</div>"
        )

        output = update_links(
            html,
            {"Sheet 1": "sheet_001"},
            update_local_links=False,
            update_ext_links=False,
        )

        self.assertIn('href="#Sheet 1.A1"', output)
        self.assertIn('href="https://example.com"', output)
        self.assertNotIn("about:srcdoc", output)
        self.assertNotIn("xlsx_sheet-link", output)

    def test_external_rewrite_ignores_non_http_protocols_and_relative_paths(self):
        html = (
            '<div>'
            '<a href="mailto:test@example.com">mail</a>'
            '<a href="tel:+15551234567">phone</a>'
            '<a href="/docs/index.html">relative</a>'
            '<a href="https://example.com">http</a>'
            "</div>"
        )

        output = update_links(html, {"Sheet 1": "sheet_001"})

        self.assertIn('href="mailto:test@example.com"', output)
        self.assertIn('href="tel:+15551234567"', output)
        self.assertIn('href="/docs/index.html"', output)
        self.assertIn('href="https://example.com"', output)
        self.assertIn('target="_blank"', output)
        self.assertIn('rel="noopener noreferrer"', output)
        self.assertEqual(1, output.count('class="xlsx_sheet-link js_visible"'))

    def test_rewrites_preserve_original_non_navigation_attributes(self):
        html = (
            '<a id="lnk" data-x="1" aria-label="go" class="legacy" '
            'href="#Sheet 1.A1"><span>Go</span></a>'
        )

        output = update_links(html, {"Sheet 1": "sheet_001"})

        self.assertIn('id="lnk"', output)
        self.assertIn('data-x="1"', output)
        self.assertIn('aria-label="go"', output)
        self.assertIn("legacy xlsx_sheet-link sharepoint_visible", output)
        self.assertIn("legacy xlsx_sheet-link js_visible", output)

    def test_external_rewrite_preserves_attributes_and_merges_rel_tokens(self):
        html = (
            '<a id="ext" data-x="1" class="legacy" rel="nofollow" target="_self" '
            'href="https://example.com"><span>Open</span></a>'
        )

        output = update_links(html, {"Sheet 1": "sheet_001"})

        self.assertIn('id="ext"', output)
        self.assertIn('data-x="1"', output)
        self.assertIn('class="legacy xlsx_sheet-link js_visible"', output)
        self.assertIn('target="_self"', output)
        self.assertIn('rel="nofollow noopener noreferrer"', output)


if __name__ == "__main__":
    unittest.main()
