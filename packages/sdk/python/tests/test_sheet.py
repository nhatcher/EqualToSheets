import pytest

from equalto.exceptions import CellReferenceError, WorkbookError
from equalto.sheet import Sheet
from equalto.workbook import Workbook


@pytest.fixture(name="sheet")
def fixture_sheet(empty_workbook: Workbook) -> Sheet:
    return empty_workbook.sheets[0]


def test_get_sheet_by_name(example_workbook: Workbook) -> None:
    sheet = example_workbook.sheets["Second"]
    assert sheet.name == "Second"
    assert sheet.index == 1


def test_get_sheet_by_index(example_workbook: Workbook) -> None:
    assert example_workbook.sheets[0].name == "Sheet1"
    assert example_workbook.sheets[1].name == "Second"
    assert example_workbook.sheets[2].name == "Sheet4"


def test_get_sheet_by_negative_index(example_workbook: Workbook) -> None:
    assert example_workbook.sheets[-8].name == "Sheet1"
    assert example_workbook.sheets[-7].name == "Second"
    assert example_workbook.sheets[-1].name == "Hidden"


def test_sheets_len(empty_workbook: Workbook, example_workbook: Workbook) -> None:
    assert len(empty_workbook.sheets) == 1
    assert len(example_workbook.sheets) == 8


def test_non_existent_sheet_name(empty_workbook: Workbook) -> None:
    with pytest.raises(WorkbookError, match='"NonExistentSheet" sheet does not exist'):
        _ = empty_workbook.sheets["NonExistentSheet"]  # noqa: WPS122


def test_sheet_index_out_of_bounds(empty_workbook: Workbook) -> None:
    with pytest.raises(WorkbookError, match="index out of bounds"):
        _ = empty_workbook.sheets[42]  # noqa: WPS122


def test_get_sheet_cell_by_reference(sheet: Sheet) -> None:
    cell = sheet["C2"]

    assert cell.row == 2
    assert cell.column == 3

    cell.formula = "=1+2"
    assert cell.value == 3


@pytest.mark.parametrize(
    "reference, error",
    [
        ("foobar", '"foobar" reference cannot be parsed'),
        ("Sheet1!A2", "sheet name cannot be specified in this context"),
    ],
)
def test_get_sheet_cell_by_invalid_reference(sheet: Sheet, reference: str, error: str) -> None:
    with pytest.raises(CellReferenceError, match=error):
        _ = sheet[reference]  # noqa: WPS122


def test_get_sheet_cell_by_index(empty_workbook: Workbook, sheet: Sheet) -> None:
    cell = sheet.cell(row=4, column=2)
    reference = "Sheet1!B4"

    assert empty_workbook[reference].value != 3

    cell.formula = "=1+2"
    assert cell.value == 3
    assert empty_workbook[reference].value == 3


def test_add_sheet(empty_workbook: Workbook) -> None:
    sheet = empty_workbook.sheets.add("New Sheet")
    assert sheet.name == "New Sheet"
    assert sheet.index == 1


def test_add_sheet_name_in_use(empty_workbook: Workbook) -> None:
    with pytest.raises(WorkbookError, match="A worksheet already exists with that name"):
        empty_workbook.sheets.add(empty_workbook.sheets[0].name)


def test_delete_cell(sheet: Sheet) -> None:
    reference = "A1"

    sheet[reference].value = 42
    del sheet[reference]
    assert not sheet[reference].value
