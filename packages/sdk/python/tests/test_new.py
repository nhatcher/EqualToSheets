from zoneinfo import ZoneInfo

import pytest
from pytest_mock import MockerFixture

import equalto
from equalto.exceptions import WorkbookError


def test_create_new_workbook() -> None:
    workbook = equalto.new()

    # perform a few basic operations in order to confirm that the workbook works as expected
    workbook["Sheet1!A1"].value = 1
    workbook["Sheet1!A2"].value = 2
    workbook["Sheet1!B1"].formula = "=A1+A2"
    assert workbook["Sheet1!B1"].value == 3


def test_create_new_workbook_timezone_parameter(mocker: MockerFixture) -> None:
    mocker.patch("equalto._get_local_tz", return_value=ZoneInfo("Europe/Berlin"))

    # the local time zone is used by default
    assert equalto.new().timezone == ZoneInfo("Europe/Berlin")

    # `timezone` parameter can be used to override the time zone
    assert equalto.new(timezone=ZoneInfo("US/Central")).timezone == ZoneInfo("US/Central")


def test_create_new_workbook_error_handling() -> None:
    """Errors coming from pycalc should be wrapped in WorkbookError."""
    with pytest.raises(WorkbookError, match="Invalid timezone: not a timezone"):
        equalto.new(timezone="not a timezone")  # type: ignore
