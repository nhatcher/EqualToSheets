import pytest

import equalto
from equalto.exceptions import WorkbookError
from equalto.workbook import Workbook


def test_json(empty_workbook: Workbook) -> None:
    empty_workbook["Sheet1!A1"].value = 42
    workbook_json = empty_workbook.json

    new_workbook = equalto.loads(workbook_json)
    assert new_workbook["Sheet1!A1"].value == 42


def test_loads_error() -> None:
    with pytest.raises(WorkbookError, match="Error parsing workbook"):
        equalto.loads("not a workbook json")
