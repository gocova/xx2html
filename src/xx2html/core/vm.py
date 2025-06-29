import logging
from zipfile import ZipFile
# from openpyxl.xml.functions import fromstring
from lxml import etree
from openpyxl.packaging.relationship import get_dependents
from typing import Tuple, Dict

_METADATA_XML = "xl/metadata.xml"
_RICHVALUE_REL_XML_RELS = "xl/richData/_rels/richValueRel.xml.rels"
_RICHVALUE_XML = "xl/richData/rdrichvalue.xml"
_RICHVALUES_STRUCTURE_XML = "xl/richData/rdrichvaluestructure.xml"

SHEET_MAIN_NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
RICHDATA_NS = "http://schemas.microsoft.com/office/spreadsheetml/2017/richdata"

EXPECTED_RICHDATA_IMAGE_TYPES= ["_localImage"]

def get_xml_from_archive(archive: ZipFile, file_path: str) -> Tuple[etree._Element | None, None | RuntimeError]:
    """
    Reads an XML file from the given archive and returns it as an lxml Element.
    If the file does not exist in the archive, returns None.
    """
    try:
        # content = archive.read(file_path)
        # return etree.fromstring(content, parser=etree.XMLParser(ns_clean=True))
        with archive.open(file_path) as xml_file:
            return etree.parse(
                xml_file,
                # parser=etree.XMLParser(ns_clean=True)
            ).getroot(), None
    except KeyError as ke:
        logging.warning(f"File {file_path} not found in the archive.")
        return None, ke
    except etree.XMLSyntaxError as e:
        logging.error(f"XML syntax error in {file_path}: {e}")
        return None, e
    except Exception as e:
        logging.error(f"Error reading {file_path} from archive: {e}")
        return None, e

namespaces = {
    'x': SHEET_MAIN_NS,
    'rd': RICHDATA_NS,
    'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
    'mc': 'http://schemas.openxmlformats.org/markup-compatibility/2006',
    'xr': 'http://schemas.microsoft.com/office/spreadsheetml/2014/revision',
    'xm': 'http://schemas.microsoft.com/office/spreadsheetml/2016/metadata', # For metadata.xml
    'xlrd': 'http://schemas.microsoft.com/office/spreadsheetml/2019/richdata', # For richData parts
    'xlam': 'http://schemas.microsoft.com/office/excel/autofill/mac',
    'wlm': 'http://schemas.microsoft.com/office/web/excel/longnames',
    'xpt': 'http://schemas.microsoft.com/office/spreadsheetml/2015/pivotTable',
    'xda': 'http://schemas.microsoft.com/office/spreadsheetml/2016/dataaccess',
    'xpr': 'http://schemas.microsoft.com/office/spreadsheetml/2015/pivotRichData',
    'xne': 'http://schemas.microsoft.com/office/excel/paneNew',
    'xnm': 'http://schemas.microsoft.com/office/excel/nameExt',
    'xsv': 'http://schemas.microsoft.com/office/spreadsheetml/2017/sheetview',
    'xcc': 'http://schemas.microsoft.com/office/spreadsheetml/2017/conditionalformat',
    'xalc': 'http://schemas.microsoft.com/office/spreadsheetml/2019/calcfeatures',
}

def get_incell_images_refs(archive: ZipFile) -> Tuple[Dict[str, str], RuntimeError | None]:
    """
    Extracts incell images from the given archive and returns a dictionary
    mapping image references to their content.
    """
    incell_images_refs = {}

    archive_namelist = archive.namelist()
    if not (
        _RICHVALUE_REL_XML_RELS in archive_namelist and \
        _METADATA_XML in archive_namelist and \
        _RICHVALUE_XML in archive_namelist and\
        _RICHVALUES_STRUCTURE_XML in archive_namelist \
    ):
        return incell_images_refs, FileExistsError("Missing required files in the archive for incell images extraction.")

    def try_parse_int(value: str) -> int | None:
        """
        Attempts to parse a string as an integer, returning None if it fails.
        """
        try:
            return int(value)
        except ValueError:
            return None
    try:
        # # 1. Get workbook.xml and its relationships to find other parts
        # workbook_rels_tree = get_xml_from_archive(archive, 'xl/_rels/workbook.xml.rels')
        # if workbook_rels_tree is None:
        #     print("Could not find workbook relationships.")
        #     return
        
        # # Find the relationship ID for metadata.xml and richData parts
        # metadata_rel_id = None
        # metadata_target = None
        # rich_data_target = None
        # rich_data_rel_id = None
        # for r in workbook_rels_tree.xpath('//r:Relationship', namespaces=namespaces):
        #     if r.get('Type') == 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/workbookMetadata':
        #         metadata_rel_id = r.get('Id')
        #         metadata_target = r.get('Target') # e.g., 'metadata.xml'
        #     if r.get('Type') == 'http://schemas.microsoft.com/office/2019/Excel/RichData/relationships/richData':
        #         rich_data_rel_id = r.get('Id')
        #         rich_data_target = r.get('Target') # e.g., 'richData/rdrichvalue.xml'

        # print(f"Metadata Rel ID: {metadata_rel_id}, Target: {metadata_target}")
        # print(f"Rich Data Rel ID: {rich_data_rel_id}, Target: {rich_data_target}")
        incell_image_rels = get_dependents(
            archive, _RICHVALUE_REL_XML_RELS
        )
        incell_images = dict([(x.Id, x.Target) for x in incell_image_rels])
        # incell_images = dict([
        #     (
        #         str(i), y
        #     )
        #     for i, y in enumerate([
        #         y[1] for y in
        #         sorted([ (x.Id, x.Target) for x in incell_image_rels])
        #     ])
        # ])
        logging.info(f"get_incell_images_refs: incell_images -> {incell_images}")

        rdrichvalue_structure_tree, err = get_xml_from_archive(
            archive, _RICHVALUES_STRUCTURE_XML
        )
        if err is not None:
            return incell_images_refs, err
        rv_types = dict([
            (lambda index:
                (
                    index,
                    (lambda x : incell_images.get(
                        f"rId{int(x[0].text)+1}", None) \
                        # f"rId{x[0].text}", None) \
                        if len(x) == 2  and x.get("s", None) == index and try_parse_int(x[0].text) is not None \
                            else None
                    )
                    # lambda x: f"'{index}' == '{x.get("s", None)}' --> {index == x.get("s", None)}"
                    # lambda x: (x[0].text)
                ) 
            )(str(i))
            for i, x in enumerate(
                rdrichvalue_structure_tree.xpath(
                    "//rd:s", namespaces=namespaces
                )
            )
            if x.get('t') in EXPECTED_RICHDATA_IMAGE_TYPES
        ])
        # print(f"RichData Types: {rv_types}")

        rdrichvalue_tree, err = get_xml_from_archive(
            archive, _RICHVALUE_XML
        )
        if err is not None:
            return incell_images_refs, err
        richdata_values = []
        for rv in rdrichvalue_tree.xpath(
            "//rd:rv", namespaces=namespaces
        ):
            rv_type_index = rv.get("s", None)
            # print(rv_type_index)
            if rv_type_index is None:
                richdata_values.append(None)
                continue

            if rv_type_index not in rv_types:
                richdata_values.append(None)
                continue

            # print(rv_types[rv_type_index](rv))
            richdata_values.append(
                rv_types[rv_type_index](rv)
            )
        richdata_values = dict([
            (str(i), v) for i, v in enumerate(richdata_values)
            if v is not None
        ])

        logging.info(f"get_incell_images_refs: richData_values -> {richdata_values}")

        

        metadata_tree, err = get_xml_from_archive(archive, 'xl/metadata.xml')
        if err is not None:
            return incell_images_refs, err
        # print(f"Metadata Tree: {metadata_tree.getchildren(f"{SHEET_MAIN_NS}metadata")}")

        for record in metadata_tree.xpath('//x:valueMetadata/x:bk/x:rc', namespaces=namespaces):
            rd_value_index = record.get("v", None)
            if rd_value_index is None:
                continue
            if rd_value_index not in richdata_values:
                continue
            final_index = try_parse_int(rd_value_index)
            if final_index is None:
                logging.warning(
                    f"get_incell_images_refs: Unable to parse rd_value_index '{rd_value_index}' as an integer."
                )
                continue
            incell_images_refs[
                str(
                    # rd_value_index
                    final_index + 1
                )] = richdata_values[rd_value_index]
        
        logging.info(f"get_incell_images_refs: incell_images_refs -> {incell_images_refs}")


    except Exception as incell_exc:
        logging.warning(
            f"get_incell_images_refs: Unable to read incell images due to: {repr(incell_exc)}"
        )
        return incell_images_refs, incell_exc
    return incell_images_refs, None
