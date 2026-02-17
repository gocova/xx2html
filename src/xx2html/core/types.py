from collections.abc import Callable
from typing import TypeAlias, TypedDict

from openpyxl.cell import Cell


CellCoordinate: TypeAlias = tuple[int | str, int]
TransformResult: TypeAlias = tuple[bool, str | None]
XlsxTransformCallable: TypeAlias = Callable[[str, str, str], TransformResult]
ConditionalFormattingRelation: TypeAlias = tuple[str, str, set[str]]


class ImageRenderData(TypedDict):
    width: int
    height: int
    style: dict[str, str]
    src: str


class CellDimensions(TypedDict):
    width: int
    height: int


class CellRenderData(TypedDict):
    attrs: dict[str, object]
    column: int | str
    row: int
    value: object
    formatted_value: str
    style: dict[str, str]
    classes: set[str]
    vm_id: str | None


class ColumnRenderData(TypedDict):
    attrs: dict[str, object]
    index: str
    width: int
    style: dict[str, str]
    hidden: bool
    collapsed: bool


class WorksheetContents(TypedDict):
    rows: list[list[CellRenderData]]
    cols: list[ColumnRenderData]
    images: dict[CellCoordinate, list[ImageRenderData]]
    vm_ids: set[str]
    vm_ids_dimension_references: dict[str, CellDimensions]
    vm_cell_vm_ids: dict[str, str]
    table_width: int


class CovaCell(Cell):
    _vm_id = None

    def __init__(
        self, worksheet, row, column, value=None, style_array=None, vm_id=None
    ):
        super().__init__(worksheet, row, column, value, style_array)
        # if vm_id is not None:
        if isinstance(vm_id, str):
            self._vm_id = vm_id

    def __repr__(self):
        if hasattr(self, "_vm_id"):
            # if self._vm_id > 0:
            if isinstance(self._vm_id, str):
                return "<CovaCell {0!r}.{1} vm_id:{2}>".format(
                    self.parent.title, self.coordinate, self._vm_id
                )
        return "<CovaCell {0!r}.{1}>".format(self.parent.title, self.coordinate)
