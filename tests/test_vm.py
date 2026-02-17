import io
import unittest
from unittest.mock import patch
from zipfile import ZIP_DEFLATED, ZipFile

from lxml import etree

import xx2html.core.vm as vm_module
from xx2html.core.vm import _get_xml_from_archive, get_incell_images_refs


class VmTests(unittest.TestCase):
    @staticmethod
    def _valid_rels_xml() -> str:
        return """
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://example" Target="../media/image1.png"/>
</Relationships>
""".strip()

    @staticmethod
    def _valid_structure_xml() -> str:
        return """
<rd:richValueStructures xmlns:rd="http://schemas.microsoft.com/office/spreadsheetml/2017/richdata">
  <rd:s t="_localImage" />
</rd:richValueStructures>
""".strip()

    @staticmethod
    def _valid_richvalue_xml() -> str:
        return """
<rd:richValueData xmlns:rd="http://schemas.microsoft.com/office/spreadsheetml/2017/richdata">
  <rd:rv s="0"><rd:v>0</rd:v><rd:v>ignored</rd:v></rd:rv>
</rd:richValueData>
""".strip()

    @staticmethod
    def _valid_metadata_xml() -> str:
        return """
<metadata xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <valueMetadata>
    <bk>
      <rc v="0" />
    </bk>
  </valueMetadata>
</metadata>
""".strip()

    def _build_archive(
        self,
        *,
        rels_xml: str | None = None,
        structure_xml: str | None = None,
        richvalue_xml: str | None = None,
        metadata_xml: str | None = None,
    ) -> io.BytesIO:
        zip_buffer = io.BytesIO()
        with ZipFile(zip_buffer, mode="w", compression=ZIP_DEFLATED) as zf:
            if rels_xml is not None:
                zf.writestr("xl/richData/_rels/richValueRel.xml.rels", rels_xml)
            if structure_xml is not None:
                zf.writestr("xl/richData/rdrichvaluestructure.xml", structure_xml)
            if richvalue_xml is not None:
                zf.writestr("xl/richData/rdrichvalue.xml", richvalue_xml)
            if metadata_xml is not None:
                zf.writestr("xl/metadata.xml", metadata_xml)
        zip_buffer.seek(0)
        return zip_buffer

    def test_get_incell_images_refs_extracts_vm_to_target_map(self):
        zip_buffer = self._build_archive(
            rels_xml=self._valid_rels_xml(),
            structure_xml=self._valid_structure_xml(),
            richvalue_xml=self._valid_richvalue_xml(),
            metadata_xml=self._valid_metadata_xml(),
        )
        with ZipFile(zip_buffer, mode="r") as zf:
            refs, err = get_incell_images_refs(zf)

        self.assertIsNone(err)
        self.assertIn("1", refs)
        self.assertTrue(refs["1"].endswith("/image1.png"))

    def test_try_parse_int_handles_none_and_invalid_values(self):
        self.assertIsNone(vm_module._try_parse_int(None))
        self.assertIsNone(vm_module._try_parse_int("x"))
        self.assertEqual(7, vm_module._try_parse_int("7"))

    def test_get_xml_from_archive_handles_missing_file(self):
        zip_buffer = self._build_archive()
        with ZipFile(zip_buffer, mode="r") as zf:
            root, err = _get_xml_from_archive(zf, "does/not/exist.xml")
        self.assertIsNone(root)
        self.assertIsInstance(err, KeyError)

    def test_get_xml_from_archive_handles_invalid_xml(self):
        zip_buffer = io.BytesIO()
        with ZipFile(zip_buffer, mode="w", compression=ZIP_DEFLATED) as zf:
            zf.writestr("broken.xml", "<x>")
        zip_buffer.seek(0)
        with ZipFile(zip_buffer, mode="r") as zf:
            root, err = _get_xml_from_archive(zf, "broken.xml")
        self.assertIsNone(root)
        self.assertIsInstance(err, etree.XMLSyntaxError)

    def test_get_xml_from_archive_handles_unexpected_exception(self):
        class BrokenArchive:
            def open(self, *_args, **_kwargs):  # noqa: ANN002, ANN003
                raise RuntimeError("boom")

        root, err = _get_xml_from_archive(BrokenArchive(), "x.xml")  # type: ignore[arg-type]
        self.assertIsNone(root)
        self.assertIsInstance(err, RuntimeError)

    def test_get_local_image_type_indexes_filters_expected_types(self):
        structure_tree = etree.fromstring(
            """
<rd:richValueStructures xmlns:rd="http://schemas.microsoft.com/office/spreadsheetml/2017/richdata">
  <rd:s t="_localImage" />
  <rd:s t="otherType" />
  <rd:s t="_localImage" />
</rd:richValueStructures>
""".strip().encode("utf-8")
        )
        self.assertEqual(
            {"0", "2"},
            vm_module._get_local_image_type_indexes(structure_tree),
        )

    def test_get_rich_data_value_targets_skips_invalid_and_keeps_valid_records(self):
        richvalue_tree = etree.fromstring(
            """
<rd:richValueData xmlns:rd="http://schemas.microsoft.com/office/spreadsheetml/2017/richdata">
  <rd:rv s="9"><rd:v>0</rd:v><rd:v>a</rd:v></rd:rv>
  <rd:rv s="0"><rd:v>0</rd:v></rd:rv>
  <rd:rv s="0"><rd:v>x</rd:v><rd:v>a</rd:v></rd:rv>
  <rd:rv s="0"><rd:v>5</rd:v><rd:v>a</rd:v></rd:rv>
  <rd:rv s="0"><rd:v>0</rd:v><rd:v>a</rd:v></rd:rv>
</rd:richValueData>
""".strip().encode("utf-8")
        )
        targets = vm_module._get_rich_data_value_targets(
            richvalue_tree=richvalue_tree,
            local_image_type_indexes={"0"},
            relationship_targets={"rId1": "xl/media/image1.png"},
        )
        self.assertEqual({"4": "xl/media/image1.png"}, targets)

    def test_map_vm_ids_to_targets_skips_invalid_metadata_rows(self):
        metadata_tree = etree.fromstring(
            """
<metadata xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <valueMetadata>
    <bk>
      <rc />
      <rc v="9" />
      <rc v="x" />
      <rc v="0" />
    </bk>
  </valueMetadata>
</metadata>
""".strip().encode("utf-8")
        )
        with self.assertLogs(level="WARNING") as log_ctx:
            mapping = vm_module._map_vm_ids_to_targets(
                metadata_tree=metadata_tree,
                rich_data_value_targets={"x": "skip-me", "0": "xl/media/image1.png"},
            )
        self.assertEqual({"1": "xl/media/image1.png"}, mapping)
        self.assertTrue(
            any("unable to parse value metadata index 'x'" in msg for msg in log_ctx.output)
        )

    def test_get_incell_images_refs_returns_file_not_found_when_parts_are_missing(self):
        zip_buffer = self._build_archive()
        with ZipFile(zip_buffer, mode="r") as zf:
            refs, err = get_incell_images_refs(zf)
        self.assertEqual({}, refs)
        self.assertIsInstance(err, FileNotFoundError)

    def test_get_incell_images_refs_returns_xml_error_for_invalid_structure_part(self):
        zip_buffer = self._build_archive(
            rels_xml=self._valid_rels_xml(),
            structure_xml="<rd:richValueStructures",
            richvalue_xml=self._valid_richvalue_xml(),
            metadata_xml=self._valid_metadata_xml(),
        )
        with ZipFile(zip_buffer, mode="r") as zf:
            refs, err = get_incell_images_refs(zf)
        self.assertEqual({}, refs)
        self.assertIsInstance(err, etree.XMLSyntaxError)

    def test_get_incell_images_refs_uses_runtime_error_when_xml_tree_missing_without_error(self):
        zip_buffer = self._build_archive(
            rels_xml=self._valid_rels_xml(),
            structure_xml=self._valid_structure_xml(),
            richvalue_xml=self._valid_richvalue_xml(),
            metadata_xml=self._valid_metadata_xml(),
        )
        with ZipFile(zip_buffer, mode="r") as zf:
            with patch("xx2html.core.vm._get_xml_from_archive", return_value=(None, None)):
                refs, err = get_incell_images_refs(zf)
        self.assertEqual({}, refs)
        self.assertIsInstance(err, RuntimeError)
        self.assertIn("rich value structure XML", str(err))

    def test_get_incell_images_refs_uses_runtime_error_for_missing_richvalue_tree(self):
        structure_tree = etree.fromstring(self._valid_structure_xml().encode("utf-8"))
        zip_buffer = self._build_archive(
            rels_xml=self._valid_rels_xml(),
            structure_xml=self._valid_structure_xml(),
            richvalue_xml=self._valid_richvalue_xml(),
            metadata_xml=self._valid_metadata_xml(),
        )
        with ZipFile(zip_buffer, mode="r") as zf:
            with patch(
                "xx2html.core.vm._get_xml_from_archive",
                side_effect=[(structure_tree, None), (None, None)],
            ):
                refs, err = get_incell_images_refs(zf)
        self.assertEqual({}, refs)
        self.assertIsInstance(err, RuntimeError)
        self.assertIn("rich value XML", str(err))

    def test_get_incell_images_refs_uses_runtime_error_for_missing_metadata_tree(self):
        structure_tree = etree.fromstring(self._valid_structure_xml().encode("utf-8"))
        richvalue_tree = etree.fromstring(self._valid_richvalue_xml().encode("utf-8"))
        zip_buffer = self._build_archive(
            rels_xml=self._valid_rels_xml(),
            structure_xml=self._valid_structure_xml(),
            richvalue_xml=self._valid_richvalue_xml(),
            metadata_xml=self._valid_metadata_xml(),
        )
        with ZipFile(zip_buffer, mode="r") as zf:
            with patch(
                "xx2html.core.vm._get_xml_from_archive",
                side_effect=[(structure_tree, None), (richvalue_tree, None), (None, None)],
            ):
                refs, err = get_incell_images_refs(zf)
        self.assertEqual({}, refs)
        self.assertIsInstance(err, RuntimeError)
        self.assertIn("metadata XML", str(err))

    def test_get_incell_images_refs_catches_unexpected_exceptions(self):
        zip_buffer = self._build_archive(
            rels_xml=self._valid_rels_xml(),
            structure_xml=self._valid_structure_xml(),
            richvalue_xml=self._valid_richvalue_xml(),
            metadata_xml=self._valid_metadata_xml(),
        )
        with ZipFile(zip_buffer, mode="r") as zf:
            with patch(
                "xx2html.core.vm._get_relationship_targets",
                side_effect=RuntimeError("boom"),
            ):
                refs, err = get_incell_images_refs(zf)
        self.assertEqual({}, refs)
        self.assertIsInstance(err, RuntimeError)
        self.assertEqual("boom", str(err))


if __name__ == "__main__":
    unittest.main()
