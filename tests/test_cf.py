import unittest

from bs4 import BeautifulSoup

from xx2html.core.cf import apply_cf_styles


class ApplyCfStylesTests(unittest.TestCase):
    def test_applies_new_classes_to_matching_cells_and_preserves_existing(self):
        html = (
            "<table><tbody>"
            '<tr><td id="Sheet1!A1" class="base">X</td></tr>'
            '<tr><td id="Sheet1!B1">Y</td></tr>'
            '<tr><td id="Sheet2!A1" class="other">Z</td></tr>'
            "</tbody></table>"
        )
        rels = [
            ("Sheet1", "A1", {"cf-red", "cf-bold"}),
            ("Sheet1", "B1", {"cf-green"}),
            ("Sheet2", "A1", {"cf-blue"}),
        ]

        output = apply_cf_styles(html, rels)

        self.assertIn('id="Sheet1!A1"', output)
        self.assertIn('class="base ', output)
        self.assertIn("cf-red", output)
        self.assertIn("cf-bold", output)

        self.assertIn('id="Sheet1!B1"', output)
        self.assertIn("cf-green", output)

        self.assertIn('id="Sheet2!A1"', output)
        self.assertIn('class="other ', output)
        self.assertIn("cf-blue", output)

    def test_ignores_non_matching_references(self):
        html = (
            "<table><tbody>"
            '<tr><td id="Sheet1!A1" class="base">X</td></tr>'
            "</tbody></table>"
        )
        rels = [
            ("Sheet1", "B2", {"cf-red"}),
            ("Unknown", "A1", {"cf-blue"}),
        ]

        output = apply_cf_styles(html, rels)

        self.assertIn('id="Sheet1!A1"', output)
        self.assertIn('class="base"', output)
        self.assertNotIn("cf-red", output)
        self.assertNotIn("cf-blue", output)

    def test_deduplicates_and_stabilizes_class_names(self):
        html = (
            "<table><tbody>"
            '<tr><td id="Sheet1!A1" class="base cf-red">X</td></tr>'
            "</tbody></table>"
        )
        rels = [("Sheet1", "A1", {"cf-red", "cf-bold"})]

        output = apply_cf_styles(html, rels)
        soup = BeautifulSoup(output, "lxml")
        classes = soup.find("td", {"id": "Sheet1!A1"}).get("class")

        self.assertEqual(["base", "cf-red", "cf-bold"], classes)


if __name__ == "__main__":
    unittest.main()
