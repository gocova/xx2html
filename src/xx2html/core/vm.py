"""Extract vm-id to media target mappings from XLSX rich-data parts."""

import logging
from typing import TypeAlias
from zipfile import ZipFile

from lxml import etree
from openpyxl.packaging.relationship import get_dependents

_METADATA_XML = "xl/metadata.xml"
_RICHVALUE_REL_XML_RELS = "xl/richData/_rels/richValueRel.xml.rels"
_RICHVALUE_XML = "xl/richData/rdrichvalue.xml"
_RICHVALUES_STRUCTURE_XML = "xl/richData/rdrichvaluestructure.xml"

SHEET_MAIN_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
RICHDATA_NS = "http://schemas.microsoft.com/office/spreadsheetml/2017/richdata"

EXPECTED_RICHDATA_IMAGE_TYPES = {"_localImage"}

NAMESPACES = {
    "x": SHEET_MAIN_NS,
    "rd": RICHDATA_NS,
}

VmId: TypeAlias = str
TargetPath: TypeAlias = str
RichDataValueIndex: TypeAlias = str

RelationshipTargets: TypeAlias = dict[str, TargetPath]
RichDataValueTargets: TypeAlias = dict[RichDataValueIndex, TargetPath]
VmIdToTargetMap: TypeAlias = dict[VmId, TargetPath]


def _try_parse_int(value: str | None) -> int | None:
    """Parse `value` as int, returning `None` when parsing fails."""
    if value is None:
        return None
    try:
        return int(value)
    except ValueError:
        return None


def _get_xml_from_archive(
    archive: ZipFile, file_path: str
) -> tuple[etree._Element | None, None | Exception]:
    """
    Reads an XML file from the given archive and returns it as an lxml Element.
    If the file does not exist in the archive, returns None.
    """
    try:
        with archive.open(file_path) as xml_file:
            return etree.parse(xml_file).getroot(), None
    except KeyError as key_error:
        logging.warning("File %s not found in the archive.", file_path)
        return None, key_error
    except etree.XMLSyntaxError as xml_exc:
        logging.error("XML syntax error in %s: %s", file_path, xml_exc)
        return None, xml_exc
    except Exception as exc:
        logging.error("Error reading %s from archive: %s", file_path, exc)
        return None, exc


def _get_relationship_targets(archive: ZipFile) -> RelationshipTargets:
    """Read relationship ids (`rId*`) and target paths for rich data."""
    image_rels = get_dependents(archive, _RICHVALUE_REL_XML_RELS)
    return {rel.Id: rel.Target for rel in image_rels}


def _get_local_image_type_indexes(richvalues_structure_tree: etree._Element) -> set[str]:
    """Collect rich-data structure indexes that represent local images."""
    local_image_type_indexes = set()
    for index, structure_node in enumerate(
        richvalues_structure_tree.xpath("//rd:s", namespaces=NAMESPACES)
    ):
        if structure_node.get("t") in EXPECTED_RICHDATA_IMAGE_TYPES:
            local_image_type_indexes.add(str(index))
    return local_image_type_indexes


def _get_rich_data_value_targets(
    richvalue_tree: etree._Element,
    local_image_type_indexes: set[str],
    relationship_targets: RelationshipTargets,
) -> RichDataValueTargets:
    """
    Returns a map from rich-data value index (metadata rc@v) to the target image path.
    """
    rich_data_value_targets: RichDataValueTargets = {}
    for value_index, rich_value_node in enumerate(
        richvalue_tree.xpath("//rd:rv", namespaces=NAMESPACES)
    ):
        rich_value_type_index = rich_value_node.get("s")
        if rich_value_type_index not in local_image_type_indexes:
            continue

        # Keep compatibility with the expected structure the old parser required.
        if len(rich_value_node) != 2:
            continue

        rel_index = _try_parse_int(rich_value_node[0].text)
        if rel_index is None:
            continue

        target_path = relationship_targets.get(f"rId{rel_index + 1}")
        if target_path is None:
            continue
        rich_data_value_targets[str(value_index)] = target_path

    return rich_data_value_targets


def _map_vm_ids_to_targets(
    metadata_tree: etree._Element, rich_data_value_targets: RichDataValueTargets
) -> VmIdToTargetMap:
    """
    Returns a map from vm_id (worksheet cell vm attr) to target image path.
    """
    vm_id_to_target: VmIdToTargetMap = {}
    for record in metadata_tree.xpath("//x:valueMetadata/x:bk/x:rc", namespaces=NAMESPACES):
        rich_data_value_index = record.get("v")
        if rich_data_value_index is None:
            continue

        target_path = rich_data_value_targets.get(rich_data_value_index)
        if target_path is None:
            continue

        vm_id_as_int = _try_parse_int(rich_data_value_index)
        if vm_id_as_int is None:
            logging.warning(
                "get_incell_images_refs: unable to parse value metadata index '%s'.",
                rich_data_value_index,
            )
            continue

        vm_id_to_target[str(vm_id_as_int + 1)] = target_path

    return vm_id_to_target


def get_incell_images_refs(
    archive: ZipFile,
) -> tuple[VmIdToTargetMap, Exception | None]:
    """
    Extracts in-cell image references from the given archive and returns:
    - vm_id -> target image path
    """
    archive_namelist = archive.namelist()
    required_files = (
        _RICHVALUE_REL_XML_RELS,
        _METADATA_XML,
        _RICHVALUE_XML,
        _RICHVALUES_STRUCTURE_XML,
    )
    if not all(required_file in archive_namelist for required_file in required_files):
        return {}, FileNotFoundError(
            "Missing required files in the archive for in-cell images extraction."
        )

    try:
        relationship_targets = _get_relationship_targets(archive)
        logging.info(
            "get_incell_images_refs: relationship_targets -> %s", relationship_targets
        )

        richvalues_structure_tree, xml_error = _get_xml_from_archive(
            archive, _RICHVALUES_STRUCTURE_XML
        )
        if xml_error is not None or richvalues_structure_tree is None:
            return {}, xml_error or RuntimeError("Failed to read rich value structure XML.")
        local_image_type_indexes = _get_local_image_type_indexes(richvalues_structure_tree)

        richvalue_tree, xml_error = _get_xml_from_archive(archive, _RICHVALUE_XML)
        if xml_error is not None or richvalue_tree is None:
            return {}, xml_error or RuntimeError("Failed to read rich value XML.")
        rich_data_value_targets = _get_rich_data_value_targets(
            richvalue_tree, local_image_type_indexes, relationship_targets
        )
        logging.info(
            "get_incell_images_refs: rich_data_value_targets -> %s",
            rich_data_value_targets,
        )

        metadata_tree, xml_error = _get_xml_from_archive(archive, _METADATA_XML)
        if xml_error is not None or metadata_tree is None:
            return {}, xml_error or RuntimeError("Failed to read metadata XML.")
        vm_id_to_target = _map_vm_ids_to_targets(metadata_tree, rich_data_value_targets)
        logging.info("get_incell_images_refs: vm_id_to_target -> %s", vm_id_to_target)

        return vm_id_to_target, None
    except Exception as incell_exc:
        logging.warning(
            "get_incell_images_refs: unable to read in-cell images due to: %r",
            incell_exc,
        )
        return {}, incell_exc
