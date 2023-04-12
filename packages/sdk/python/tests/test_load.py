import os

import pytest

import equalto
from equalto.exceptions import SuppressEvaluationErrors, WorkbookError


def test_load_workbook() -> None:
    filename = os.path.join(os.path.dirname(__file__), "xlsx", "example.xlsx")
    workbook = equalto.load(filename)

    assert workbook["Sheet1!A1"].value == "A string"
    assert workbook["Sheet1!A2"].value == 222
    assert workbook["Second!A1"].value == "Tres"


@pytest.mark.parametrize(
    "file_name, error",
    [
        ("non_existent.xlsx", "I/O Error: No such file or directory"),
        ("not_zip_file.xlsx", "EqualTo can only open workbooks created by Microsoft Excel"),
        ("corrupt.xlsx", "EqualTo can only open workbooks created by Microsoft Excel"),
        ("circ.xlsx", r"Calculation!A1 \('=A1'\): Circular reference detected"),
        (
            "array_formula.xlsx",
            "EqualTo cannot open this workbook due to the following unsupported features: array formulas. ",
        ),
        ("XLOOKUP_with_errors.xlsx", "Models are different"),
        ("UNSUPPORTED_FNS_DAYS_NETWORKDAYS.xlsx", "Invalid function: _xlfn.DAYS"),
    ],
)
def test_load_workbook_error_handling(file_name: str, error: str) -> None:
    with pytest.raises(WorkbookError, match=error):
        path = os.path.join(os.path.dirname(__file__), "xlsx", file_name)
        equalto.load(path)


def test_load_suppress_errors() -> None:
    with SuppressEvaluationErrors() as context:
        path = os.path.join(os.path.dirname(__file__), "xlsx", "circ.xlsx")
        workbook = equalto.load(path)

        assert workbook["Calculation!A1"].value == "#CIRC!"
        assert context.suppressed_errors(workbook) == [
            "EqualTo could not evaluate this workbook without errors. "
            + "This may indicate a bug or missing feature in the EqualTo spreadsheet calculation engine. "
            + "Please contact support@equalto.com, share the entirety of this error message and the relevant workbook, "
            + "and we will work with you to resolve the issue. Detailed error message:\n"
            + "Calculation!A1 ('=A1'): Circular reference detected",
        ]

    # non-evaluation errors should not be suppressed
    with SuppressEvaluationErrors():
        path = os.path.join(os.path.dirname(__file__), "xlsx", "corrupt.xlsx")
        with pytest.raises(WorkbookError, match="EqualTo can only open workbooks created by Microsoft Excel"):
            equalto.load(path)
