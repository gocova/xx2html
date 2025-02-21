import logging
from typing import List, Set, Tuple
from bs4 import BeautifulSoup


def apply_cf_styles(html: str, cf_styles_rels: List[Tuple[str, str, Set[str]]]) -> str:
    soup = BeautifulSoup(html, "lxml")
    for sheet_name, cell_ref, class_names in cf_styles_rels:
        class_names_str = " ".join(class_names)
        logging.debug(
            f"apply_cf_styles: '{sheet_name}!{cell_ref}' with additional class_names: {class_names_str}"
        )
        for td_tag in soup.findAll("td", id=f"{sheet_name}!{cell_ref}"):
            # print(" ".join(class_names))
            prev_classes = td_tag.get("class")
            prev_classes: List[str] = prev_classes if prev_classes is not None else []
            # print(prev_classes)
            # new_td_tag["class"] = " ".join([prev_classes, class_names_str])
            td_tag["class"] = prev_classes + [x for x in class_names]
            # print(new_td_tag)
    return str(soup)
