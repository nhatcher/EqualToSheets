import os
from datetime import timezone

import pytest

import equalto
from equalto.cell import Cell
from equalto.workbook import Workbook


@pytest.fixture(name="empty_workbook")
def fixture_empty_workbook() -> Workbook:
    return equalto.new(timezone=timezone.utc)


@pytest.fixture
def cell(empty_workbook: Workbook) -> Cell:
    return empty_workbook["Sheet1!A1"]


@pytest.fixture
def example_workbook() -> Workbook:
    filename = os.path.join(os.path.dirname(__file__), "xlsx", "example.xlsx")
    return equalto.load(filename)
