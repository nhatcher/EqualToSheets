import os
from tempfile import TemporaryDirectory
from zoneinfo import ZoneInfo

import pytest

import equalto
from equalto.cell import Cell
from equalto.exceptions import CellReferenceError, SuppressEvaluationErrors, WorkbookEvaluationError
from equalto.workbook import Workbook


def test_workbook_repr(example_workbook: Workbook) -> None:
    assert repr(example_workbook) == "<Workbook: example>"


def test_workbook_name(example_workbook: Workbook) -> None:
    assert example_workbook.name == "example"


def test_get_cell_by_reference(empty_workbook: Workbook) -> None:
    cell = empty_workbook["Sheet1!C2"]

    assert cell.row == 2
    assert cell.column == 3

    cell.formula = "=1+2"
    assert cell.value == 3


def test_get_cell_by_index(empty_workbook: Workbook) -> None:
    cell = empty_workbook.cell(sheet_index=0, row=4, column=2)

    cell.formula = "=1+2"
    assert cell.value == 3
    assert empty_workbook["Sheet1!B4"].value == 3


@pytest.mark.parametrize(
    "reference, error",
    [
        ("foobar", '"foobar" reference cannot be parsed'),
        ("C2", '"C2" reference is missing the sheet name'),
    ],
)
def test_get_cell_by_invalid_reference(
    empty_workbook: Workbook,
    reference: str,
    error: str,
) -> None:
    with pytest.raises(CellReferenceError, match=error):
        _ = empty_workbook[reference]  # noqa: WPS122


def test_delete_cell(empty_workbook: Workbook) -> None:
    reference = "Sheet1!A1"

    empty_workbook[reference].value = 42
    del empty_workbook[reference]
    assert not empty_workbook[reference].value


@pytest.mark.parametrize("tz", [ZoneInfo("UTC"), ZoneInfo("Europe/Berlin")])
def test_timezone_property(tz: ZoneInfo) -> None:
    assert equalto.new(timezone=tz).timezone == tz


def test_save_xlsx(empty_workbook: Workbook) -> None:
    empty_workbook["Sheet1!A1"].value = 42

    with TemporaryDirectory() as dir_path:
        file_path = os.path.join(dir_path, "output.xlsx")
        empty_workbook.save(file_path)

        # simple assert confirming that the exported file is valid
        assert equalto.load(file_path)["Sheet1!A1"].value == 42


def test_suppress_evaluation_errors(cell: Cell) -> None:
    # the errors are normally raised before entering the context
    with pytest.raises(WorkbookEvaluationError):
        cell.formula = "=INVALID()"

    # the errors are suppressed when in the context
    with SuppressEvaluationErrors() as context:
        cell.formula = "=INVALID()"
        errors = context.suppressed_errors(cell.workbook)

    assert errors == ["Sheet1!A1 ('=INVALID()'): Invalid function: INVALID"]
    assert cell.formula == "=INVALID()"
    assert cell.value == "#ERROR!"

    # the errors are normally raised after leaving the context
    with pytest.raises(WorkbookEvaluationError):
        cell.formula = "=INVALID()"
