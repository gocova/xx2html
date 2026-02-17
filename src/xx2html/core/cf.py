"""Conditional-formatting HTML post-processing."""

import logging
from bs4 import BeautifulSoup

from xx2html.core.types import ConditionalFormattingRelation


def apply_cf_styles(
    html: str, cf_style_relations: list[ConditionalFormattingRelation]
) -> str:
    """Attach generated conditional-formatting class names to target cells."""
    soup = BeautifulSoup(html, "lxml")
    for sheet_name, cell_ref, class_names in cf_style_relations:
        class_names_str = " ".join(class_names)
        logging.debug(
            f"apply_cf_styles: '{sheet_name}!{cell_ref}' with additional class_names: {class_names_str}"
        )
        for cell_tag in soup.find_all("td", id=f"{sheet_name}!{cell_ref}"):
            previous_classes = cell_tag.get("class")
            previous_classes = (
                previous_classes if previous_classes is not None else []
            )
            for class_name in sorted(class_names):
                if class_name not in previous_classes:
                    previous_classes.append(class_name)
            cell_tag["class"] = previous_classes
    return str(soup)
