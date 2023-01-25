from datetime import date, datetime, timezone
from typing import Any
from zoneinfo import ZoneInfo

import pytest

import equalto
from equalto.cell import Cell, CellType
from equalto.exceptions import WorkbookError, WorkbookValueError
from equalto.workbook import Workbook


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
    cell.style.format = "#.###"
    cell.value = 9.87654
    assert str(cell) == "9.877"
    cell.delete()
    assert not cell.value
    assert cell.style.format == "general"


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


@pytest.mark.parametrize(
    "cell_reference, formula, error",
    [
        ("Sheet1!A1", "=INVALID()", "Sheet1!A1 ('=INVALID()'): Invalid function: INVALID"),
        ("Sheet1!A1", "=SIN(1, 2, 3)", "Sheet1!A1 ('=SIN(1,2,3)'): Wrong number of arguments"),
        ("Sheet1!C1", "={{1}}", "Sheet1!C1 ('={{1}}'): Arrays not implemented"),
        ("Sheet1!ABC42", "=2*ABC42", "Sheet1!ABC42 ('=2*ABC42'): Circular reference detected"),
        ("Sheet1!H2", "=[1]", "Sheet1!H2 ('=[1]'): Error parsing [1]: Unexpected token: '['"),
    ],
)
def test_formula_error_propagation(
    empty_workbook: Workbook,
    cell_reference: str,
    formula: str,
    error: str,
) -> None:
    cell = empty_workbook[cell_reference]

    with pytest.raises(WorkbookError) as err:
        cell.formula = formula

    assert err.value.args[0] == error  # noqa: WPS441


@pytest.mark.parametrize(
    "workbook_timezone, date_value, raw_value",
    [
        ("US/Central", date(1969, 7, 20), 25404),
        ("Europe/Berlin", date(1969, 7, 20), 25404),
        ("UTC", date(2023, 1, 9), 44935),
        ("Asia/Dhaka", date(2030, 1, 1), 47484),
    ],
)
def test_date_value(workbook_timezone: str, date_value: date, raw_value: int) -> None:
    cell = _get_tz_cell(workbook_timezone)

    cell.value = date_value
    assert cell.value == raw_value  # type: ignore

    # confirm that `date_value` property properly converts the raw value back to a date
    assert cell.date_value == date_value


@pytest.mark.parametrize(
    "value, error",
    [
        ("foobar", "'foobar' is not a number"),
        (44903.8, "44903.8 is not an integer"),
        (-2, "-2 does not represent a valid date"),
    ],
)
def test_date_value_error(cell: Cell, value: Any, error: str) -> None:
    cell.value = value
    with pytest.raises(WorkbookValueError, match=error):
        _ = cell.date_value  # noqa: WPS122


@pytest.mark.parametrize(
    "workbook_timezone, datetime_value, raw_value",
    [
        # compare a few random dates with values computed by Excel from MS Office 2019
        ("UTC", datetime(1993, 7, 21, 17, 15, 43, tzinfo=timezone.utc), 34171.71925),
        ("UTC", datetime(2008, 12, 24, 19, 10, tzinfo=timezone.utc), 39806.79861),
        # confirm that DST offset is properly handled
        ("UTC", datetime(2022, 12, 8, 19, 12, tzinfo=timezone.utc), 44903.8),
        ("US/Central", datetime(2022, 12, 8, 19, 12, tzinfo=timezone.utc), 44903.55),  # no DST
        ("UTC", datetime(2022, 7, 8, 12, 30, tzinfo=timezone.utc), 44750.52083),
        ("US/Central", datetime(2022, 7, 8, 12, 30, tzinfo=timezone.utc), 44750.3125),  # DST offset
        ("UTC", datetime(2023, 1, 9, tzinfo=ZoneInfo("Europe/Berlin")), 44934.95833),
        ("Europe/Berlin", datetime(2023, 1, 9, tzinfo=ZoneInfo("Europe/Berlin")), 44935),
        # using "Asia/Dhaka" should add 6 hours (0.25 of a day)
        ("UTC", datetime(2030, 1, 1, tzinfo=timezone.utc), 47484),
        ("Asia/Dhaka", datetime(2030, 1, 1, tzinfo=timezone.utc), 47484.25),
    ],
)
def test_datetime_value(
    workbook_timezone: str,
    datetime_value: datetime,
    raw_value: float,
) -> None:
    cell = _get_tz_cell(workbook_timezone)

    cell.value = datetime_value
    assert pytest.approx(cell.value, abs=0.00001) == raw_value

    # confirm that `datetime_value` property properly converts the raw value back to a datetime
    assert cell.datetime_value == datetime_value


@pytest.mark.parametrize(
    "value, error",
    [
        ("foobar", "'foobar' is not a number"),
        (-1.5, "-1.5 does not represent a valid datetime"),
    ],
)
def test_datetime_value_error(cell: Cell, value: Any, error: str) -> None:
    cell.value = value
    with pytest.raises(WorkbookValueError, match=error):
        _ = cell.datetime_value  # noqa: WPS122


def test_set_naive_datetime(cell: Cell) -> None:
    with pytest.raises(WorkbookValueError, match="Naive datetime encountered: 2023-01-09 12:30:15"):
        cell.value = datetime(2023, 1, 9, 12, 30, 15)


def _get_tz_cell(tz: str) -> Cell:
    return equalto.new(timezone=ZoneInfo(tz)).sheets[0]["A1"]
