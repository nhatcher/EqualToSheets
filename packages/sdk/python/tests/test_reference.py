from __future__ import annotations

import pytest

from equalto.exceptions import CellReferenceError
from equalto.reference import parse_cell_reference


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
    "reference, error",
    [
        ("Sheet!!A1", '"Sheet!!A1" reference cannot be parsed'),
        ("Sheet!AA", '"Sheet!AA" reference cannot be parsed'),
        ("foobar", '"foobar" reference cannot be parsed'),
    ],
)
def test_invalid_reference(reference: str, error: str) -> None:
    with pytest.raises(CellReferenceError, match=error):
        parse_cell_reference(reference)
