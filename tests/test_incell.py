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


if __name__ == "__main__":
    unittest.main()
