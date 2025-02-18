from openpyxl.cell import Cell


class CovaCell(Cell):
    _vm_id = 0

    def __init__(
        self, worksheet, row, column, value=None, style_array=None, vm_id=None
    ):
        super().__init__(worksheet, row, column, value, style_array)
        if vm_id is not None:
            self._vm_id = vm_id

    def __repr__(self):
        if hasattr(self, "_vm_id"):
            if self._vm_id > 0:
                return "<CovaCell {0!r}.{1} vm_id:{2}>".format(
                    self.parent.title, self.coordinate, self._vm_id
                )
        return "<CovaCell {0!r}.{1}>".format(self.parent.title, self.coordinate)
