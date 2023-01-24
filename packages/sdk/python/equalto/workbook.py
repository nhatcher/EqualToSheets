from __future__ import annotations

import os
from functools import cached_property
from typing import TYPE_CHECKING
from zoneinfo import ZoneInfo

from equalto.exceptions import CellReferenceError
from equalto.reference import parse_cell_reference
from equalto.sheet import WorkbookSheets

if TYPE_CHECKING:
    from equalto._equalto import PyCalcModel
    from equalto.cell import Cell


class Workbook:
    def __init__(self, model: PyCalcModel):
        self._model = model

    def __getitem__(self, key: str) -> Cell:
        """Get cell by the reference (i.e. "Sheet1!A1")."""
        sheet_name, row, column = parse_cell_reference(key)
        if sheet_name is None:
            raise CellReferenceError(f'"{key}" reference is missing the sheet name')
        return self.sheets[sheet_name].cell(row, column)

    def __delitem__(self, key: str) -> None:
        """Delete the cell content and style."""
        self[key].delete()

    @cached_property
    def timezone(self) -> ZoneInfo:
        return ZoneInfo(self._model.get_timezone())

    def cell(self, sheet_index: int, row: int, column: int) -> Cell:
        return self.sheets[sheet_index].cell(row, column)

    @cached_property
    def sheets(self) -> WorkbookSheets:
        """Get container with workbook sheets."""
        return WorkbookSheets(self)

    def save(self, file: str) -> None:
        _, ext = os.path.splitext(file)
        if ext == ".xlsx":
            self._model.save_to_xlsx(file)
        else:
            raise NotImplementedError(f"Exporting to {ext} files is not supported yet.")
