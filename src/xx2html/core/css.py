from typing import Callable, Literal, Tuple
from openpyxl.cell import Cell
from openpyxl.styles.colors import Color

from xx2html.core.types import CovaCell
from condif2css.color import aRGB_to_CSS

from kivy.logger import Logger

DEFAULT_BORDER_STYLE = (
    "border-{direction}-style: solid; border-{direction}-width: 1px; "
)
BORDER_STYLES = {
    "dashDot": None,
    "dashDotDot": None,
    "dashed": "border-{direction}-style: dashed; ",
    "dotted": "border-{direction}-style: dotted; ",
    "double": "border-{direction}-style: double; ",
    "hair": None,
    "medium": "border-{direction}-style: solid; border-{direction}-width: 2px; ",
    "mediumDashDot": "border-{direction}-style: solid; border-{direction}-width: 2px; ",
    "mediumDashDotDot": "border-{direction}-style: solid; border-{direction}-width: 2px; ",
    "mediumDashed": "border-{direction}-style: dashed; border-{direction}-width: 2px; ",
    "slantDashDot": None,
    "thick": "border-{direction}-style: solid; border-{direction}-width: 1px;",
    "thin": "border-{direction}-style: solid; border-{direction}-width: 1px; ",
}


class CssRegistry:
    def __init__(
        self,
        get_css_color: Callable[[Color], str | None],
        classes: dict[str, object] = dict(),
    ) -> None:
        self.get_css_color = get_css_color
        self.classes = classes

    def register_font_size(self, size: int) -> str:
        class_name = f"xlsx_cell_font_size_{size}"
        if class_name not in self.classes:
            self.classes[class_name] = f"font-size: {size}px;"
        return class_name

    def register_height(self, size: int) -> str:
        class_name = f"xlsx_element_height_{size}"
        if class_name not in self.classes:
            self.classes[class_name] = f"height: {size}px;"
        return class_name

    def register_font_color(self, color: Color) -> str | None:
        aRGB = self.get_css_color(color)
        if aRGB is not None:
            class_name = f"xlsx_cell_font_color_{aRGB}"

            if class_name not in self.classes:
                self.classes[class_name] = f"color: {aRGB_to_CSS(aRGB)};"
            return class_name
        else:
            return None

    def register_background_color(self, color: Color) -> str | None:
        aRGB = self.get_css_color(color)
        if aRGB is not None:
            class_name = f"xlsx_cell_background_color_{aRGB}"

            if class_name not in self.classes:
                self.classes[class_name] = f"background-color : {aRGB_to_CSS(aRGB)}"

            return class_name
        else:
            return None

    def register_font_underline(self) -> str:
        class_name = "xlsx_cell_font_underline"
        if class_name not in self.classes:
            self.classes[class_name] = "font-decoration: underline;"
        return class_name

    def register_font_bold(self) -> str:
        class_name = "xlsx_cell_font_bold"
        if class_name not in self.classes:
            self.classes[class_name] = "font-weight: bold;"
        return class_name

    def register_font_italic(self) -> str:
        class_name = "xlsx_cell_font_italic"
        if class_name not in self.classes:
            self.classes[class_name] = "font-style: italic;"
        return class_name

    def register_border(
        self,
        style: str | None,
        direction: Literal["right", "left", "top", "bottom"],
        color: Color,
    ) -> str | None:
        if style is None:
            return None

        border_style = BORDER_STYLES.get(style)
        border_style = DEFAULT_BORDER_STYLE if border_style is None else border_style

        class_name = f"border_{style}_{direction[0:1]}"

        class_value = border_style.format(direction=direction)

        aRGB = self.get_css_color(color)

        if aRGB is not None:
            class_name = f"{class_name}_{aRGB}"
            class_value = (
                f"{class_value} border-{direction}-color: {aRGB_to_CSS(aRGB)};"
            )

        if class_name not in self.classes:
            self.classes[class_name] = class_value

        return class_name


def get_border_classes_from_cell(cell: Cell | CovaCell, css_registry: CssRegistry):
    classes = set()

    for b_dir in ["right", "left", "top", "bottom"]:
        b_s = getattr(cell.border, b_dir)
        if not b_s:
            continue

        border_class = css_registry.register_border(
            b_s.style,
            direction=b_dir,  # type:ignore
            color=b_s.color,
        )
        if border_class is not None:
            classes.add(border_class)

    return classes


def create_get_css_components_from_cell(css_registry: CssRegistry):
    def get_css_components_from_cell(
        cell: Cell | CovaCell,
        merged_cell_map=None,  # odefault_cell_border="none"
    ) -> Tuple[dict, set]:
        merged_cell_map = merged_cell_map or {}

        classes = set()

        # h_styles = {"border-collapse": "collapse"}
        h_styles = {}
        classes.update(get_border_classes_from_cell(cell, css_registry))
        if merged_cell_map:
            # TODO edged_cells
            for m_cell in merged_cell_map["cells"]:
                classes.update(get_border_classes_from_cell(m_cell, css_registry))

        # for b_dir in ["border-right", "border-left", "border-top", "border-bottom"]:
        # style_tag = b_dir + "-style"
        # if (b_dir not in b_styles) and (style_tag not in b_styles):
        # b_styles[b_dir] = default_cell_border
        # h_styles.update(b_styles)

        if cell.alignment.horizontal:
            h_styles["text-align"] = cell.alignment.horizontal
        if cell.alignment.vertical:
            h_styles["vertical-align"] = cell.alignment.vertical

        # with contextlib.suppress(AttributeError):
        if hasattr(cell, "fill") and hasattr(cell.fill, "patternType"):
            patternType = cell.fill.patternType
            if  patternType == "solid":
                background_color_class = css_registry.register_background_color(
                    cell.fill.fgColor
                )
                if background_color_class is not None:
                    classes.add((background_color_class))
                # h_styles["background-color"] = get_css_color(cell.fill.fgColor)
            elif patternType is not None:
                # TODO patternType != 'solid'
                Logger.warning(
                    f"css (components): Pattern type is not supported: {cell.fill.patternType}"
                )

        if cell.font:
            classes.add(css_registry.register_font_size(int(cell.font.sz)))
            # h_styles["font-size"] = "%spx" % cell.font.sz
            if cell.font.color:
                font_color_class = css_registry.register_font_color(cell.font.color)
                if font_color_class is not None:
                    classes.add(font_color_class)
                # h_styles["color"] = get_css_color(cell.font.color)
            if cell.font.b:
                classes.add(css_registry.register_font_bold())
                # h_styles["font-weight"] = "bold"
            if cell.font.i:
                classes.add(css_registry.register_font_italic())
                # h_styles["font-style"] = "italic"
            if cell.font.u:
                classes.add(css_registry.register_font_underline())
                # h_styles["font-decoration"] = "underline"

        return h_styles, classes

    return get_css_components_from_cell
