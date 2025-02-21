from typing import Set, Tuple, Callable, List

import logging
from zipfile import ZipFile

from openpyxl import load_workbook
from openpyxl.packaging.relationship import get_dependents
from openpyxl.styles.differential import DifferentialStyleList

# Monkey patch!
from xx2html.core.cf import apply_cf_styles
from xx2html.core.patches.openpyxl import apply_patches

from .incell import get_incell_css
from .links import update_links
from .utils import cova__render_table, get_worksheet_contents

# from .css import CssRegistry, create_get_css_components_from_cell
from condif2css.processor import process
from condif2css.themes import get_theme_colors
from condif2css.core import create_themed_get_css_color
from condif2css.color import aRGB_to_css
from condif2css.css import CssBuilder, CssRulesRegistry, create_get_css_from_cell

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
    apply_cf: bool = False
) -> Callable[[str, str, str], Tuple[bool, None | str]]:
    def xlsx_transform(
        source: str, dest: str, locale: str
    ) -> Tuple[bool, None | str]:  # -> (bool, None|str):
        try:
            logging.info(f"Transform (out): Opening '{dest}' for writing...")
            output = open(dest, "w", encoding="utf-8")

            links = []
            html_tables = []
            # classes = {}

            logging.info(f"Transform (wb): Reading '{source}' as xlsx file...")
            wb = load_workbook(source, data_only=True, rich_text=True)

            logging.debug("Transform (wb|css): Reading theme colors...")
            theme_aRGBs_list = get_theme_colors(wb)
            # get_css_color = create_themed_get_css_color(theme_aRGBs_list)
            # css_registry = CssRegistry(get_css_color, classes)
            # get_css_components_from_cell = create_get_css_components_from_cell(
            #     css_registry
            # )
            pre_get_css_color = create_themed_get_css_color(theme_aRGBs_list)

            def get_css_color(color):
                argb_color = pre_get_css_color(color)
                if argb_color is None:
                    return None
                return aRGB_to_css(argb_color)

            css_builder = CssBuilder(get_css_color)
            css_registry = CssRulesRegistry()
            get_css_from_cell = create_get_css_from_cell(
                css_registry, css_builder=css_builder
            )

            # conditional formatting vars
            css_cf_registry = CssRulesRegistry(prefix="xx2h_cf")
            get_cf_css_from_diff = create_get_css_from_cell(
                css_registry=css_cf_registry, css_builder=css_builder
            )

            logging.debug("Transform (wb|incell): Reading incell images...")
            incell_images = None
            archive = None
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

            effective_cf_rules_details = {}  # all conditional formatting rules

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
                        # css_registry,
                        # get_css_components_from_cell,
                        css_rules_registry=css_registry,
                        css_builder=css_builder,
                        get_css_from_cell=get_css_from_cell,
                        locale=locale,
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

                    if apply_cf:
                        logging.info(
                            f"Application (wb|cf): Processing conditional formatting for '{sheet_name}'"
                        )
                        effective_cf_rules_details.update(process(ws))

            # generated_css = "\n".join([f".{k} {{ {v} }}" for k, v in classes.items()])
            generated_css = "\n".join(css_registry.get_rules())

            if archive:
                logging.debug(
                    "Transform (wb|incell): Preparing incell images output..."
                )
                generated_incell_css = get_incell_css(
                    vm_ids, vm_ids_dimension_references, incell_images, archive
                )
            else:
                generated_incell_css = ""

            logging.info(
                f"Transform (html|1): Pass 1 --> Preparing {len(effective_cf_rules_details)} conditional formatting styles..."
            )
            cf_styles_rels: List[Tuple[str, str, Set[str]]] = []
            if hasattr(wb, "_differential_styles") and isinstance(
                wb._differential_styles,  # type: ignore
                DifferentialStyleList,
            ):
                # print(effective_cf_rules_details)
                for full_ref, details in effective_cf_rules_details.items():
                    sheet_name, cell_ref, _, dxf_id, _ = details
                    class_names = get_cf_css_from_diff(
                        wb._differential_styles[dxf_id],  # type: ignore
                        is_important=True,
                    )
                    cf_styles_rels.append((sheet_name, cell_ref, class_names))
                # print(css_cf_registry.get_rules())
            print(cf_styles_rels)

            logging.debug("Transform (html|2): Pass 2 --> Preparing html")

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
                    conditional_css_html=f"<style>/*conditional formatting*/\n{'\n'.join(css_cf_registry.get_rules())}</style>",
                )
                .replace('"$"', "$")
                .replace('"-"', "-")
            )

            logging.debug("Transform (html|3): Pass 3 --> Updating links...")
            html_2 = update_links(
                html, enc_names, update_local_links=update_local_links
            )

            logging.debug(
                "Transform (html|4): Pass 4 --> Applying conditional formatting styles..."
            )
            html_3 = apply_cf_styles(html_2, cf_styles_rels)

            logging.info(f"Transform (out): Writing output: {dest}")
            output.write(html_3)
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
