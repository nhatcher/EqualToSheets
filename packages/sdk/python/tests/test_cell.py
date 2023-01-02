from typing import Any

import pytest

from equalto.cell import Cell, CellType
from equalto.exceptions import WorkbookError, WorkbookValueError
from equalto.workbook import Workbook


@pytest.fixture(name="cell")
def fixture_cell(empty_workbook: Workbook) -> Cell:
    return empty_workbook["Sheet1!A1"]


@pytest.mark.parametrize(
    "value, text",
    [
        ("foo", "foo"),
        (42, "42"),
        (1.23, "1.23"),
        (False, "FALSE"),
        (None, ""),
    ],
)
def test_cell_str(cell: Cell, value: Any, text: str) -> None:
    cell.value = value
    assert str(cell) == text


def test_set_formula(empty_workbook: Workbook) -> None:
    sheet = empty_workbook.sheets[0]

    sheet["A1"].value = 2
    sheet["A2"].value = 7
    sheet["B1"].formula = "=A1*A2*3"

    assert sheet["B1"].value == 42
    assert sheet["B1"].formula == "=A1*A2*3"


def test_set_invalid_formula(cell: Cell) -> None:
    with pytest.raises(WorkbookError, match='"foo" is not a valid formula'):
        cell.formula = "foo"


@pytest.mark.parametrize(
    "value",
    [
        "foobar",
        "42",  # a string with number assigned to `value` should not be automatically converted
        "",
        "=1+2",  # formula assigned to `value` should not be evaluated
        "#VALUE!",
    ],
)
def test_set_text_value(cell: Cell, value: str) -> None:
    cell.value = value
    assert cell.value == value
    assert str(cell) == value
    assert cell.formula is None


@pytest.mark.parametrize(
    "value, cell_type",
    [
        ("foobar", CellType.text),
        (True, CellType.logical_value),
        (42, CellType.number),
        (12.34, CellType.number),
        # TODO: An empty string should likely be considered a number (like an empty cell).
        ("", CellType.text),
    ],
)
def test_cell_type(cell: Cell, value: Any, cell_type: CellType) -> None:
    cell.value = value
    assert cell.type == cell_type


@pytest.mark.parametrize(
    "formula, cell_type",
    [
        ("=1/0", CellType.error_value),
        ("=TRUE", CellType.logical_value),
        ("=40+2", CellType.number),
        ("=12.34", CellType.number),
        ('="foo"', CellType.text),
        ("=A100", CellType.number),  # an empty cell
    ],
)
def test_formula_cell_type(cell: Cell, formula: str, cell_type: CellType) -> None:
    cell.formula = formula
    assert cell.type == cell_type


def test_delete_cell(cell: Cell) -> None:
    cell.value = 42
    cell.delete()
    assert not cell.value
    # TODO: Once styles are introduced, confirm that the cell style is deleted as well.


def test_int_value(cell: Cell) -> None:
    cell.value = 42.0
    assert cell.int_value == 42


@pytest.mark.parametrize(
    "value, error",
    [
        ("foobar", "'foobar' is not a number"),
        (True, "True is not a number"),
        (None, "'' is not a number"),
        ("4", "'4' is not a number"),  # the value is not automatically converted
        (4.2, "4.2 is not an integer"),
    ],
)
def test_int_value_error(cell: Cell, value: Any, error: str) -> None:
    cell.value = value
    with pytest.raises(WorkbookValueError, match=error):
        _ = cell.int_value  # noqa: WPS122


@pytest.mark.parametrize("value", [4.2, 7.0, -1.2, -100.0])
def test_float_value(cell: Cell, value: float) -> None:
    cell.value = value
    assert cell.float_value == value


@pytest.mark.parametrize(
    "value, error",
    [
        ("foobar", "'foobar' is not a number"),
        (True, "True is not a number"),
        (None, "'' is not a number"),
        ("4.2", "'4.2' is not a number"),  # the value is not automatically converted
    ],
)
def test_float_value_error(cell: Cell, value: Any, error: str) -> None:
    cell.value = value
    with pytest.raises(WorkbookValueError, match=error):
        _ = cell.float_value  # noqa: WPS122


@pytest.mark.parametrize("value", ["foobar", "=A1", "#VALUE!", "4.2", "1", "True", ""])
def test_str_value(cell: Cell, value: str) -> None:
    cell.value = value
    assert cell.str_value == value


@pytest.mark.parametrize(
    "value, error",
    [
        (4.2, "4.2 is not a string value"),
        (7, "7.0 is not a string value"),
        (True, "True is not a string value"),
    ],
)
def test_str_value_error(cell: Cell, value: Any, error: str) -> None:
    cell.value = value
    with pytest.raises(WorkbookValueError, match=error):
        _ = cell.str_value  # noqa: WPS122


@pytest.mark.parametrize("value", [True, False])
def test_bool_value(cell: Cell, value: bool) -> None:
    cell.value = value
    assert cell.bool_value is value


@pytest.mark.parametrize(
    "value, error",
    [
        ("foobar", "'foobar' is not a logical value"),
        (4.2, "4.2 is not a logical value"),
        (7, "7.0 is not a logical value"),
        ("True", "'True' is not a logical value"),  # the value is not automatically converted
    ],
)
def test_bool_value_error(cell: Cell, value: Any, error: str) -> None:
    cell.value = value
    with pytest.raises(WorkbookValueError, match=error):
        _ = cell.bool_value  # noqa: WPS122
