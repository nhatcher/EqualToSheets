from __future__ import annotations

import json
from datetime import date, datetime, timedelta
from enum import Enum
from functools import cached_property
from typing import TYPE_CHECKING

from equalto._equalto import number_to_column
from equalto.exceptions import WorkbookValueError
from equalto.style import Style

if TYPE_CHECKING:
    from equalto._equalto import PyCalcModel
    from equalto.sheet import Sheet
    from equalto.workbook import Workbook


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
        return self._model.get_formatted_cell_value(*self.cell_ref)

    def __repr__(self) -> str:
        return f"<Cell: {self.text_ref}>"

    @property
    def type(self) -> CellType:
        return CellType(self._model.get_cell_type(*self.cell_ref))

    @property
    def value(self) -> float | bool | str | date | datetime | None:
        """Get raw value from the represented cell."""
        value = json.loads(self._model.get_cell_value_by_index(*self.cell_ref))
        assert value is None or isinstance(value, (float, bool, str))
        return value

    @value.setter
    def value(self, value: float | bool | str | date | datetime | None) -> None:
        if value is None:
            value = ""

        if isinstance(value, str):
            # NOTE: At the moment we can't manually set an error value, i.e. "#VALUE!" will be
            #       treated as string.
            self._model.update_cell_with_text(*self.cell_ref, value)
        elif isinstance(value, bool):
            self._model.update_cell_with_bool(*self.cell_ref, value)
        elif isinstance(value, (float, int)):
            self._model.update_cell_with_number(*self.cell_ref, float(value))
        elif isinstance(value, (date, datetime)):
            self._model.update_cell_with_number(*self.cell_ref, self._get_excel_date(value))
        else:  # pragma: no cover
            raise ValueError(f"unrecognized value type ({value=})")

        self._model.evaluate()

    @property
    def str_value(self) -> str:
        value = self.value
        if not isinstance(value, str):
            raise WorkbookValueError(f"{repr(value)} is not a string value")
        return value

    @property
    def float_value(self) -> float:
        value = self.value
        if not isinstance(value, float):
            raise WorkbookValueError(f"{repr(value)} is not a number")
        return value

    @property
    def int_value(self) -> int:
        float_value = self.float_value
        if not float_value.is_integer():
            raise WorkbookValueError(f"{float_value} is not an integer")
        return int(float_value)

    @property
    def bool_value(self) -> bool:
        value = self.value
        if not isinstance(value, bool):
            raise WorkbookValueError(f"{repr(value)} is not a logical value")
        return value

    @property
    def date_value(self) -> date:
        int_value = self.int_value
        if int_value < 0:
            raise WorkbookValueError(f"{int_value} does not represent a valid date")
        return self._excel_base_dt.date() + timedelta(days=int_value)

    @property
    def datetime_value(self) -> datetime:
        float_value = self.float_value
        if float_value < 0:
            raise WorkbookValueError(f"{float_value} does not represent a valid datetime")
        naive_dt = self._excel_base_dt + timedelta(days=float_value)
        return naive_dt.replace(tzinfo=self.workbook.timezone)

    @property
    def formula(self) -> str | None:
        formula = self._model.get_cell_formula(*self.cell_ref)
        assert formula is None or isinstance(formula, str)
        return formula

    @formula.setter
    def formula(self, formula: str) -> None:
        self._model.update_cell_with_formula(*self.cell_ref, formula)
        self._model.evaluate()

    def set_user_input(self, value: str) -> None:
        """
        Update the cell emulating a user typing something in Excel.

        Receives a string representation of the value and attempts to guess the correct
        type and style/formatting.
        """
        self._model.set_user_input(*self.cell_ref, value)
        self._model.evaluate()

    @property
    def style(self) -> Style:
        return Style(self)

    def delete(self) -> None:
        """Delete the cell content and style."""
        self._model.delete_cell(*self.cell_ref)

    @cached_property
    def workbook(self) -> Workbook:
        return self.sheet.workbook_sheets.workbook

    @property
    def cell_ref(self) -> tuple[int, int, int]:
        return self.sheet.index, self.row, self.column

    @property
    def text_ref(self) -> str:
        column_repr = number_to_column(self.column)
        return f"{self.sheet.name}!{column_repr}{self.row}"

    @cached_property
    def _model(self) -> PyCalcModel:
        return self.workbook._model  # noqa: WPS437

    _excel_base_dt = datetime(1899, 12, 30)

    def _get_excel_date(self, dt: date | datetime) -> float:
        if isinstance(dt, datetime):
            if dt.utcoffset() is None:
                raise WorkbookValueError(f"Naive datetime encountered: {dt}")
            naive_dt = dt.astimezone(self.workbook.timezone).replace(tzinfo=None)
            time_delta = naive_dt - self._excel_base_dt
            return float(time_delta.days) + (float(time_delta.seconds) / 86400)
        return float((dt - self._excel_base_dt.date()).days)
