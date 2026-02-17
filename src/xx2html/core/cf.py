import logging
from bs4 import BeautifulSoup

from xx2html.core.types import ConditionalFormattingRelation


def apply_cf_styles(
    html: str, cf_style_relations: list[ConditionalFormattingRelation]
) -> str:
    soup = BeautifulSoup(html, "lxml")
    for sheet_name, cell_ref, class_names in cf_style_relations:
        class_names_str = " ".join(class_names)
        logging.debug(
            f"apply_cf_styles: '{sheet_name}!{cell_ref}' with additional class_names: {class_names_str}"
        )
        for cell_tag in soup.find_all("td", id=f"{sheet_name}!{cell_ref}"):
            # print(" ".join(class_names))
            previous_classes = cell_tag.get("class")
            previous_classes = (
                previous_classes if previous_classes is not None else []
            )
            # print(prev_classes)
            # new_td_tag["class"] = " ".join([prev_classes, class_names_str])
            cell_tag["class"] = previous_classes + [x for x in class_names]
            # print(new_td_tag)
    return str(soup)
