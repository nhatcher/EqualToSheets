import os
from unittest.mock import patch

import pytest

import equalto
from equalto.exceptions import WorkbookError


def test_load_workbook() -> None:
    filename = os.path.join(os.path.dirname(__file__), "example.xlsx")
    workbook = equalto.load(filename)

    assert workbook["Sheet1!A1"].value == "A string"
    assert workbook["Sheet1!A2"].value == 222
    assert workbook["Second!A1"].value == "Tres"


def test_load_workbook_error_handling() -> None:
    """Errors coming from pycalc should be wrapped in WorkbookError."""
    # PanicException should be caught and wrapped in WorkbookError
    with pytest.raises(WorkbookError, match="No such file or directory"):
        equalto.load(os.path.join(os.path.dirname(__file__), "non-existent.xlsx"))

    # confirm that non-panic base exceptions are re-raised
    with patch("equalto.load_excel", side_effect=BaseException):
        with pytest.raises(BaseException):
            equalto.load("filename.xlsx")
