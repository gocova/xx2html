import io
import unittest
from zipfile import ZIP_DEFLATED, ZipFile

from PIL import Image

from xx2html.core.incell import get_incell_css


def _make_png(width: int, height: int) -> bytes:
    image = Image.new("RGBA", (width, height), (255, 0, 0, 255))
    buffer = io.BytesIO()
    image.save(buffer, format="PNG")
    return buffer.getvalue()


class InCellCssTests(unittest.TestCase):
    def test_wide_image_uses_full_width_auto_height(self):
        zip_buffer = io.BytesIO()
        with ZipFile(zip_buffer, mode="w", compression=ZIP_DEFLATED) as zf:
            zf.writestr("xl/media/image1.png", _make_png(400, 100))

        zip_buffer.seek(0)
        with ZipFile(zip_buffer, mode="r") as zf:
            css = get_incell_css(
                vm_ids={"1"},
                vm_ids_dimension_references={"cell_1": {"width": 100, "height": 100}},
                vm_cell_vm_ids={"cell_1": "1"},
                incell_images_refs={"1": "xl/media/image1.png"},
                archive=zf,
            )

        self.assertIn(".vm-richvaluerel_rid1 img", css)
        self.assertIn(".cell_1 img", css)
        self.assertIn("width: 100%;", css)
        self.assertIn("height: auto;", css)
        self.assertIn("object-fit: contain;", css)

    def test_tall_image_uses_auto_width_full_height(self):
        zip_buffer = io.BytesIO()
        with ZipFile(zip_buffer, mode="w", compression=ZIP_DEFLATED) as zf:
            zf.writestr("xl/media/image2.png", _make_png(100, 400))

        zip_buffer.seek(0)
        with ZipFile(zip_buffer, mode="r") as zf:
            css = get_incell_css(
                vm_ids={"2"},
                vm_ids_dimension_references={"cell_2": {"width": 100, "height": 100}},
                vm_cell_vm_ids={"cell_2": "2"},
                incell_images_refs={"2": "xl/media/image2.png"},
                archive=zf,
            )

        self.assertIn(".vm-richvaluerel_rid2 img", css)
        self.assertIn(".cell_2 img", css)
        self.assertIn("width: auto;", css)
        self.assertIn("height: 100%;", css)
        self.assertIn("object-fit: contain;", css)

    def test_missing_vm_reference_logs_warning_and_skips_vm_content_rule(self):
        zip_buffer = io.BytesIO()
        with ZipFile(zip_buffer, mode="w", compression=ZIP_DEFLATED) as zf:
            zf.writestr("xl/media/image1.png", _make_png(50, 50))

        zip_buffer.seek(0)
        with ZipFile(zip_buffer, mode="r") as zf:
            with self.assertLogs(level="WARNING") as log_ctx:
                css = get_incell_css(
                    vm_ids={"1"},
                    vm_ids_dimension_references={"cell_1": {"width": 120, "height": 80}},
                    vm_cell_vm_ids={"cell_1": "1"},
                    incell_images_refs={},
                    archive=zf,
                )

        self.assertTrue(
            any("not found in incell images references" in msg for msg in log_ctx.output)
        )
        self.assertNotIn(".vm-richvaluerel_rid1 img", css)
        self.assertIn(".cell_1 img", css)
        self.assertIn("width: 100%;", css)
        self.assertIn("height: 100%;", css)

    def test_missing_image_file_logs_warning_and_skips_vm_content_rule(self):
        zip_buffer = io.BytesIO()
        with ZipFile(zip_buffer, mode="w", compression=ZIP_DEFLATED):
            pass

        zip_buffer.seek(0)
        with ZipFile(zip_buffer, mode="r") as zf:
            with self.assertLogs(level="WARNING") as log_ctx:
                css = get_incell_css(
                    vm_ids={"2"},
                    vm_ids_dimension_references={"cell_2": {"width": 90, "height": 90}},
                    vm_cell_vm_ids={"cell_2": "2"},
                    incell_images_refs={"2": "xl/media/missing.png"},
                    archive=zf,
                )

        self.assertTrue(any("not found in excel file" in msg for msg in log_ctx.output))
        self.assertNotIn(".vm-richvaluerel_rid2 img", css)
        self.assertIn(".cell_2 img", css)

    def test_unreadable_image_falls_back_to_default_fit_rules(self):
        zip_buffer = io.BytesIO()
        with ZipFile(zip_buffer, mode="w", compression=ZIP_DEFLATED) as zf:
            zf.writestr("xl/media/image3.png", b"not-a-valid-image")

        zip_buffer.seek(0)
        with ZipFile(zip_buffer, mode="r") as zf:
            with self.assertLogs(level="WARNING") as log_ctx:
                css = get_incell_css(
                    vm_ids={"3"},
                    vm_ids_dimension_references={"cell_3": {"width": 100, "height": 100}},
                    vm_cell_vm_ids={"cell_3": "3"},
                    incell_images_refs={"3": "xl/media/image3.png"},
                    archive=zf,
                )

        self.assertTrue(any("Unable to read image size" in msg for msg in log_ctx.output))
        self.assertIn(".vm-richvaluerel_rid3 img", css)
        self.assertIn(".cell_3 img", css)
        self.assertIn("width: 100%;", css)
        self.assertIn("height: 100%;", css)

    def test_missing_cell_vm_mapping_uses_default_cell_fit_rules(self):
        zip_buffer = io.BytesIO()
        with ZipFile(zip_buffer, mode="w", compression=ZIP_DEFLATED) as zf:
            zf.writestr("xl/media/image1.png", _make_png(64, 64))

        zip_buffer.seek(0)
        with ZipFile(zip_buffer, mode="r") as zf:
            css = get_incell_css(
                vm_ids={"1"},
                vm_ids_dimension_references={"cell_missing_vm": {"width": 80, "height": 120}},
                vm_cell_vm_ids={},
                incell_images_refs={"1": "xl/media/image1.png"},
                archive=zf,
            )

        self.assertIn(".vm-richvaluerel_rid1 img", css)
        self.assertIn(".cell_missing_vm img", css)
        self.assertIn("width: 100%;", css)
        self.assertIn("height: 100%;", css)


if __name__ == "__main__":
    unittest.main()
