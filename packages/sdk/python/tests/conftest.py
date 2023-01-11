import os
from datetime import timezone

import pytest

import equalto
from equalto.workbook import Workbook


@pytest.fixture
def empty_workbook() -> Workbook:
    return equalto.new(timezone=timezone.utc)


@pytest.fixture
def example_workbook() -> Workbook:
    filename = os.path.join(os.path.dirname(__file__), "xlsx", "example.xlsx")
    return equalto.load(filename)
