from __future__ import annotations

import json
from enum import Enum
from functools import cached_property
from typing import TYPE_CHECKING

from equalto.exceptions import WorkbookError

if TYPE_CHECKING:
    from equalto._pycalc import PyCalcModel
    from equalto.sheet import Sheet


class CellType(Enum):
    """Cell type enum matching Excel TYPE() function values."""

    number = 1
    text = 2
    logical_value = 4
    error_value = 16
    array = 64
    compound_data = 128


class Cell:
    """Represents a single cell."""

    def __init__(self, sheet: Sheet, row: int, column: int) -> None:
        self.sheet = sheet
        self.row = row
        self.column = column

    def __str__(self) -> str:
        """Get formatted cell value."""
        return self._model.get_text_at(*self._cell_ref)

    @property
    def type(self) -> CellType:
        return CellType(self._model.get_cell_type(*self._cell_ref))

    @property
    def value(self) -> float | bool | str | None:
        value = json.loads(self._model.get_cell_value_by_index(*self._cell_ref))
        assert value is None or isinstance(value, (float, bool, str))
        return value

    @value.setter
    def value(self, value: float | bool | str | None) -> None:
        if value is None:
            value = ""

        if isinstance(value, str):
            # NOTE: At the moment we can't manually set an error value, i.e. "#VALUE!" will be
            #       treated as string.
            self._model.update_cell_with_text(*self._cell_ref, value)
        elif isinstance(value, bool):
            self._model.update_cell_with_bool(*self._cell_ref, value)
        elif isinstance(value, (float, int)):
            self._model.update_cell_with_number(*self._cell_ref, float(value))
        else:  # pragma: no cover
            raise ValueError(f"unrecognized value type ({value=})")

        self._model.evaluate()

    @property
    def formula(self) -> str | None:
        if not self._model.has_formula(*self._cell_ref):
            return None
        formula = self._model.get_formula_or_value(*self._cell_ref)
        assert isinstance(formula, str)
        return formula

    @formula.setter
    def formula(self, formula: str) -> None:
        if not formula.startswith("="):
            raise WorkbookError(f'"{formula}" is not a valid formula')
        self._model.set_input(*self._cell_ref, formula)
        self._model.evaluate()

    def delete(self) -> None:
        """Delete the cell content and style."""
        self._model.delete_cell(*self._cell_ref)

    @cached_property
    def _model(self) -> PyCalcModel:
        return self.sheet.workbook_sheets.workbook._model  # noqa: WPS437

    @property
    def _cell_ref(self) -> tuple[int, int, int]:
        return self.sheet.index, self.row, self.column
