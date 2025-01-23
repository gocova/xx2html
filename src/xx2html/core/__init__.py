from typing import Tuple, Callable

import logging
from zipfile import ZipFile

from openpyxl import load_workbook
from openpyxl.packaging.relationship import get_dependents

# Monkey patch!
from xx2html.core.patches.openpyxl import apply_patches

from .incell import get_incell_css
from .links import update_links
from .utils import cova__render_table, get_worksheet_contents
from .css import CssRegistry, create_get_css_components_from_cell
from condif2css.themes import get_theme_colors
from condif2css.core import create_themed_get_css_color


apply_patches()


def get_xlsx_transform(
    sheet_html: str,
    sheetname_html: str,
    index_html: str,
    fonts_html: str,
    core_css: str,
    user_css: str,
    update_local_links: bool = True,
    # prepare_iframe_noscript: bool = True,
) -> Callable[[str, str, str], Tuple[bool, None | str]]:
    def xlsx_transform(
        source: str, dest: str, locale: str
    ) -> Tuple[bool, None | str]:  # -> (bool, None|str):
        try:
            logging.info(f"Transform (out): Opening '{dest}' for writing...")
            output = open(dest, "w", encoding="utf-8")

            links = []
            html_tables = []
            classes = {}

            logging.info(f"Transform (wb): Reading '{source}' as xlsx file...")
            wb = load_workbook(source, data_only=True, rich_text=True)

            logging.debug("Transform (wb|css): Reading theme colors...")
            theme_aRGBs_list = get_theme_colors(wb)
            get_css_color = create_themed_get_css_color(theme_aRGBs_list)
            css_registry = CssRegistry(get_css_color, classes)
            get_css_components_from_cell = create_get_css_components_from_cell(
                css_registry
            )

            logging.debug("Transform (wb|incell): Reading incell images...")
            incell_images = None
            try:
                archive = ZipFile(source, "r")
                incell_image_rels = get_dependents(
                    archive, "xl/richData/_rels/richValueRel.xml.rels"
                )
                incell_images = dict([(x.Id, x.Target) for x in incell_image_rels])
                logging.info("Transform (wb|incell): Reading complete!")
            except Exception as incell_exc:
                logging.warning(
                    f"Transform (wb|incell): Unable to read incell images due to: {repr(incell_exc)}"
                )
                # archive = None
            finally:
                if incell_images is None:
                    incell_images = dict()

            vm_ids = set()
            vm_ids_dimension_references = dict()

            enc_names = dict()

            sheets_names = wb.sheetnames

            for sheet_name in sheets_names:
                ws = wb[sheet_name]

                ws_index = wb.index(ws)
                enc_sheet_name = f"sheet_{hex(ws_index)[2:].zfill(3)}"

                enc_names[sheet_name] = enc_sheet_name

                if ws.sheet_state == "visible":
                    logging.info(
                        f"Application (ws): Sheet[{ws_index}]:'{sheet_name}' (enc_sheet_name: {enc_sheet_name}) -> is visible"
                    )

                    contents = get_worksheet_contents(
                        ws,
                        css_registry,
                        get_css_components_from_cell,
                        locale,
                        ws_index=ws_index,
                    )

                    vm_ids.update(contents["vm_ids"])
                    vm_ids_dimension_references.update(
                        contents["vm_ids_dimension_references"]
                    )

                    html_tables.append(
                        sheet_html.format(
                            enc_sheet_name=enc_sheet_name,
                            sheet_name=sheet_name,
                            table_generated_html=cova__render_table(contents),
                        )
                    )

                    links.append(
                        sheetname_html.format(
                            enc_sheet_name=enc_sheet_name, sheet_name=sheet_name
                        )
                    )

            generated_css = "\n".join([f".{k} {{ {v} }}" for k, v in classes.items()])

            if archive:
                logging.debug(
                    "Transform (wb|incell): Preparing incell images output..."
                )
                generated_incell_css = get_incell_css(
                    vm_ids, vm_ids_dimension_references, incell_images, archive
                )
            else:
                generated_incell_css = ""

            logging.debug("Transform (html|1): Preparing html (pass 1)")
            html = (
                index_html.format(
                    sheets_generated_html="\n".join(html_tables),
                    sheets_names_generated_html="\n".join(links),
                    source_filename=source,
                    fonts_html=fonts_html,
                    core_css_html=f"<style>{core_css}</style>",
                    user_css_html=f"<style>{user_css}</style>",
                    generated_css_html=f"<style>{generated_css}</style>",
                    generated_incell_css_html=f"<style>{generated_incell_css}</style>",
                    conditional_css_html="<style>/*conditional formatting*/</style>",
                )
                .replace('"$"', "$")
                .replace('"-"', "-")
            )

            logging.debug("Transform (html|2): Updating links...")
            html_2 = update_links(
                html, enc_names, update_local_links=update_local_links
            )

            logging.info(f"Transform (out): Writing output: {dest}")
            output.write(html_2)
            output.flush()
            logging.info(f"Transform (out): Closing output: {dest}")
            output.close()
            logging.info(f"Transform (wb): Closing wb: {source}")
            wb.close()
            logging.info("Transform (wb): Done!")

            return (True, None)
        except Exception as e:
            logging.error(e)
            print(e)
            print(repr(e))
            return (False, repr(e))

    return xlsx_transform
