from __future__ import annotations

from functools import cached_property
from typing import TYPE_CHECKING

from equalto.cell import Cell
from equalto.exceptions import CellReferenceError, WorkbookError
from equalto.reference import parse_cell_reference

if TYPE_CHECKING:
    from equalto._pycalc import PyCalcModel
    from equalto.workbook import Workbook


class Sheet:
    """
    Represents a single sheet in the workbook.

    Instances of this class reference the underlying sheet by `sheet_id` and use dynamic properties
    to access the name and the index. This way the objects don't need to be recreated when
    the sheets are renamed/moved.
    """

    def __init__(self, workbook_sheets: WorkbookSheets, sheet_id: int) -> None:
        self.workbook_sheets = workbook_sheets
        self.sheet_id = sheet_id

    def __getitem__(self, key: str) -> Cell:
        """Get cell by Excel reference."""
        sheet_name, row, column = parse_cell_reference(key)
        if sheet_name is not None:
            raise CellReferenceError("sheet name cannot be specified in this context")
        return self.cell(row, column)

    def __delitem__(self, key: str) -> None:
        """Delete the cell content and style."""
        self[key].delete()

    def cell(self, row: int, column: int) -> Cell:
        return Cell(sheet=self, row=row, column=column)

    @property
    def name(self) -> str:
        return self.workbook_sheets.get_sheet_name(self.sheet_id)

    @property
    def index(self) -> int:
        return self.workbook_sheets.get_sheet_index(self.sheet_id)


class WorkbookSheets:
    """Represents all sheets in the workbook."""

    def __init__(self, workbook: Workbook) -> None:
        self.workbook = workbook
        self._load_sheets_metadata()

    def __getitem__(self, key: str | int) -> Sheet:
        """Get sheet by either name or index."""
        if isinstance(key, str):
            return Sheet(self, sheet_id=self._get_sheet_id_from_name(name=key))
        elif isinstance(key, int):
            if key < 0:
                key = len(self) + key
            return Sheet(self, sheet_id=self._get_sheet_id_from_index(index=key))
        raise ValueError("invalid sheet lookup key type")  # pragma: no cover

    def __len__(self) -> int:
        return len(self._sheet_index_to_sheet_id)

    def add(self, name: str) -> Sheet:
        result = self._model.add_sheet(name)
        if not result:
            # TODO: Shouldn't rust raise a proper exception with error message instead of
            #       returning a boolean?
            raise WorkbookError(f'could not create "{name}" sheet')
        self._load_sheets_metadata()
        return self[-1]

    def get_sheet_name(self, sheet_id: int) -> str:
        return self._sheet_id_to_sheet_name[sheet_id]

    def get_sheet_index(self, sheet_id: int) -> int:
        return self._sheet_id_to_sheet_index[sheet_id]

    _sheet_name_to_sheet_id: dict[str, int]
    _sheet_id_to_sheet_name: dict[int, str]
    _sheet_index_to_sheet_id: dict[int, int]
    _sheet_id_to_sheet_index: dict[int, int]

    def _load_sheets_metadata(self) -> None:
        """
        Load sheets metadata from the workbook.

        This method should be called whenever we create, delete, rename or move a sheet.
        """
        names = self._model.get_worksheet_names()
        ids = self._model.get_worksheet_ids()

        assert len(names) == len(set(names))
        assert len(ids) == len(set(ids))

        self._sheet_name_to_sheet_id = {name: sheet_id for name, sheet_id in zip(names, ids)}
        self._sheet_id_to_sheet_name = {sheet_id: name for name, sheet_id in zip(names, ids)}
        self._sheet_index_to_sheet_id = {index: sheet_id for index, sheet_id in enumerate(ids)}
        self._sheet_id_to_sheet_index = {sheet_id: index for index, sheet_id in enumerate(ids)}

    def _get_sheet_id_from_name(self, name: str) -> int:
        try:
            return self._sheet_name_to_sheet_id[name]
        except KeyError:
            raise WorkbookError(f'"{name}" sheet does not exist')

    def _get_sheet_id_from_index(self, index: int) -> int:
        try:
            return self._sheet_index_to_sheet_id[index]
        except KeyError:
            raise WorkbookError("index out of bounds")

    @cached_property
    def _model(self) -> PyCalcModel:
        return self.workbook._model  # noqa: WPS437
