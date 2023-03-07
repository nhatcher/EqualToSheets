from __future__ import annotations

import pytest

from equalto.exceptions import CellReferenceError, WorkbookError
from equalto.sheet import Sheet
from equalto.workbook import Workbook


@pytest.fixture(name="sheet")
def fixture_sheet(empty_workbook: Workbook) -> Sheet:
    return empty_workbook.sheets[0]


def test_sheet_repr(sheet: Sheet) -> None:
    assert repr(sheet) == "<Sheet: Sheet1>"


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


def test_sheets_iter(example_workbook: Workbook) -> None:
    assert [(sheet.index, sheet.name) for sheet in example_workbook.sheets] == [
        (0, "Sheet1"),
        (1, "Second"),
        (2, "Sheet4"),
        (3, "shared"),
        (4, "Table"),
        (5, "Sheet2"),
        (6, "Created fourth"),
        (7, "Hidden"),
    ]


def test_sheet_delete(example_workbook: Workbook) -> None:
    sheet = example_workbook.sheets["Sheet4"]
    sheet_table = example_workbook.sheets["Table"]
    assert sheet.index < sheet_table.index

    sheet.delete()

    assert _get_sheets(example_workbook) == [
        (0, "Sheet1"),
        (1, "Second"),
        (2, "shared"),
        (3, "Table"),
        (4, "Sheet2"),
        (5, "Created fourth"),
        (6, "Hidden"),
    ]

    # confirm that references to other sheets are not broken
    assert sheet_table.name == "Table"
    assert sheet_table.index == 3
    assert sheet_table["A1"].value == "Cars"


def test_delete_sheet_by_name(example_workbook: Workbook) -> None:
    del example_workbook.sheets["Table"]

    assert _get_sheets(example_workbook) == [
        (0, "Sheet1"),
        (1, "Second"),
        (2, "Sheet4"),
        (3, "shared"),
        (4, "Sheet2"),
        (5, "Created fourth"),
        (6, "Hidden"),
    ]


def test_delete_sheet_by_index(example_workbook: Workbook) -> None:
    del example_workbook.sheets[6]

    assert _get_sheets(example_workbook) == [
        (0, "Sheet1"),
        (1, "Second"),
        (2, "Sheet4"),
        (3, "shared"),
        (4, "Table"),
        (5, "Sheet2"),
        (6, "Hidden"),
    ]


def test_delete_only_sheet(empty_workbook: Workbook) -> None:
    assert len(empty_workbook.sheets) == 1
    with pytest.raises(WorkbookError, match="Cannot delete only sheet"):
        del empty_workbook.sheets[0]


def test_delete_non_existent_sheet(example_workbook: Workbook) -> None:
    sheet = example_workbook.sheets[0]
    sheet.delete()

    with pytest.raises(WorkbookError, match="Sheet not found"):
        sheet.delete()


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


@pytest.mark.parametrize(
    "row, column",
    [
        (1, 0),
        (0, 1),
        (-1, 1),
    ],
)
def test_get_sheet_cell_by_invalid_index(sheet: Sheet, row: int, column: int) -> None:
    with pytest.raises(WorkbookError, match="invalid cell index"):
        sheet.cell(row, column)


def test_add_sheet(empty_workbook: Workbook) -> None:
    sheet = empty_workbook.sheets.add("New Sheet")
    assert sheet.name == "New Sheet"
    assert sheet.index == 1


def test_add_sheet_default_name(empty_workbook: Workbook) -> None:
    sheet = empty_workbook.sheets.add()
    assert sheet.name == "Sheet2"
    assert sheet.index == 1

    sheet = empty_workbook.sheets.add()
    assert sheet.name == "Sheet3"
    assert sheet.index == 2


def test_add_sheet_name_in_use(empty_workbook: Workbook) -> None:
    with pytest.raises(WorkbookError, match="A worksheet already exists with that name"):
        empty_workbook.sheets.add(empty_workbook.sheets[0].name)


def test_add_sheet_invalid_name(empty_workbook: Workbook) -> None:
    with pytest.raises(WorkbookError, match="Invalid name for a sheet: '.*?'"):
        empty_workbook.sheets.add(".*?")


def test_rename_sheet(example_workbook: Workbook) -> None:
    assert _get_sheets(example_workbook) == [
        (0, "Sheet1"),
        (1, "Second"),
        (2, "Sheet4"),
        (3, "shared"),
        (4, "Table"),
        (5, "Sheet2"),
        (6, "Created fourth"),
        (7, "Hidden"),
    ]

    sheet = example_workbook.sheets["Sheet4"]
    sheet.name = "New name of Sheet4"

    assert sheet.name == "New name of Sheet4"
    assert _get_sheets(example_workbook) == [
        (0, "Sheet1"),
        (1, "Second"),
        (2, "New name of Sheet4"),
        (3, "shared"),
        (4, "Table"),
        (5, "Sheet2"),
        (6, "Created fourth"),
        (7, "Hidden"),
    ]


def test_rename_sheet_noop(sheet: Sheet) -> None:
    sheet.name = sheet.name


def test_rename_sheet_name_in_use(example_workbook: Workbook) -> None:
    with pytest.raises(WorkbookError, match="Sheet already exists: 'Second'"):
        example_workbook.sheets[0].name = "Second"


def test_rename_sheet_invalid_name(sheet: Sheet) -> None:
    with pytest.raises(WorkbookError, match="Invalid name for a sheet: '.*?'"):
        sheet.name = ".*?"


def test_delete_cell(sheet: Sheet) -> None:
    reference = "A1"

    sheet[reference].value = 42
    del sheet[reference]
    assert not sheet[reference].value


def test_sheet_instance_reused(empty_workbook: Workbook) -> None:
    assert id(empty_workbook.sheets[0]) == id(empty_workbook.sheets["Sheet1"])


def test_cell_instance_reused(empty_workbook: Workbook) -> None:
    cell = empty_workbook.sheets[0].cell(1, 1)
    assert id(cell) == id(empty_workbook["Sheet1!A1"])


def _get_sheets(workbook: Workbook) -> list[tuple[int, str]]:
    return [(sheet.index, sheet.name) for sheet in workbook.sheets]
