from __future__ import annotations

import pytest

from equalto.exceptions import CellReferenceError
from equalto.reference import parse_cell_range_reference, parse_cell_reference


@pytest.mark.parametrize(
    "reference, result",
    [
        ("C5", (None, 5, 3)),
        ("Z2", (None, 2, 26)),
        ("AB53", (None, 53, 28)),
        ("Calculation!C42", ("Calculation", 42, 3)),
        ("DataSheet!AA7", ("DataSheet", 7, 27)),
    ],
)
def test_parse_cell_reference(reference: str, result: tuple[str | None, int, int]) -> None:
    assert parse_cell_reference(reference) == result


@pytest.mark.parametrize(
    "reference, result",
    [
        ("A1:C5", ((1, 1), (5, 3))),
        ("G20:B20", ((20, 2), (20, 7))),
        ("ZZ2:ZZ1", ((1, 702), (2, 702))),
        ("Z5:ZZ6", ((5, 26), (6, 702))),
    ],
)
def test_parse_cell_range_reference(reference: str, result: tuple[tuple[int, int], tuple[int, int]]) -> None:
    assert parse_cell_range_reference(reference) == result


@pytest.mark.parametrize("reference", ["Sheet!!A1", "Sheet!AA", "foobar"])
def test_invalid_cell_reference(reference: str) -> None:
    with pytest.raises(CellReferenceError, match=f'"{reference}" reference cannot be parsed'):
        parse_cell_reference(reference)


@pytest.mark.parametrize("reference", ["Sheet!A1:A3", "A1:B", "A:B1", "1:2"])
def test_invalid_cell_range_reference(reference: str) -> None:
    with pytest.raises(CellReferenceError, match=f'"{reference}" reference cannot be parsed'):
        parse_cell_range_reference(reference)
