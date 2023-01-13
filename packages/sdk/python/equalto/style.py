from __future__ import annotations

import json
from functools import cached_property
from typing import TYPE_CHECKING

if TYPE_CHECKING:
    from equalto._pycalc import PyCalcModel
    from equalto.cell import Cell


class Style:
    def __init__(self, cell: Cell) -> None:
        self.cell = cell
        self._data = json.loads(self._model.get_style_for_cell(*self.cell.cell_ref))

    @property
    def format(self) -> str:
        return self._data["num_fmt"]

    @format.setter
    def format(self, number_format: str) -> None:
        self._data["num_fmt"] = number_format
        self._model.set_cell_style(*self.cell.cell_ref, json.dumps(self._data))
        self._model.evaluate()

    @cached_property
    def _model(self) -> PyCalcModel:
        return self.cell._model  # noqa: WPS437
