import os

import pytest

import equalto
from equalto.exceptions import WorkbookError


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
        ("not_zip_file.xlsx", "Zip Error: invalid Zip archive"),
        ("corrupt.xlsx", 'XML Error: Missing "name" XML attribute'),
    ],
)
def test_load_workbook_error_handling(file_name: str, error: str) -> None:
    with pytest.raises(WorkbookError, match=error):
        path = os.path.join(os.path.dirname(__file__), "xlsx", file_name)
        equalto.load(path)
