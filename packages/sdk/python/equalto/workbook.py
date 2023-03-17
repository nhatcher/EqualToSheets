from __future__ import annotations

import os
from functools import cached_property
from typing import TYPE_CHECKING
from zoneinfo import ZoneInfo

from equalto.exceptions import CellReferenceError, SuppressEvaluationErrors, WorkbookError, WorkbookEvaluationError
from equalto.reference import parse_cell_reference
from equalto.sheet import WorkbookSheets

if TYPE_CHECKING:
    from equalto._equalto import PyCalcModel
    from equalto.cell import Cell


class Workbook:
    def __init__(self, model: PyCalcModel):
        self._model = model
        self._check_model_support()

    def __repr__(self) -> str:
        return f"<Workbook: {self.name}>"

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
    def name(self) -> str:
        return self._model.get_name()

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

    @property
    def json(self) -> str:
        return self._model.to_json()

    def evaluate(self) -> None:
        errors = self._model.evaluate_with_error_check()
        if not errors:
            return

        if SuppressEvaluationErrors.in_context():
            SuppressEvaluationErrors.log_errors(self, errors)
        else:
            raise WorkbookEvaluationError(errors[0])

    def _check_model_support(self) -> None:
        try:
            self._model.check_model_support()
        except WorkbookError as err:
            error_message = str(err)
            if SuppressEvaluationErrors.in_context():
                SuppressEvaluationErrors.log_errors(self, [error_message])
            else:
                raise WorkbookEvaluationError(error_message)
