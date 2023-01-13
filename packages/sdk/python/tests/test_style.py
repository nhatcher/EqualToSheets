from typing import Any

import pytest

from equalto.cell import Cell


@pytest.mark.parametrize(
    "value, number_format, text",
    [
        ("foo", "general", "foo"),
        (42, "general", "42"),
        (1.23, "general", "1.23"),
        (False, "general", "FALSE"),
        (None, "general", ""),
        (123.45, "$#,##0", "$123"),
        (41343.125, "#,##0.00", "41,343.13"),
        (41343.125, "yyyy-mm-dd", "2013-03-10"),
        (41343.125, "yyyy, mmmm dd (dddd)", "2013, March 10 (Sunday)"),
        (125, "invalid", "#VALUE!"),
    ],
)
def test_cell_str_format(cell: Cell, value: Any, number_format: str, text: str) -> None:
    """Confirm that Cell.__str__ respects the cell formatting."""
    cell.value = value
    cell.style.format = number_format
    assert str(cell) == text
    assert cell.style.format == number_format
