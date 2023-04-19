from __future__ import annotations

import re

from equalto.exceptions import CellReferenceError

_CELL_REFERENCE_REGEX = re.compile(r"^(?:(?P<sheet>[^!]+)!)?(?P<column>[A-Z]+)(?P<row>\d+)$")
_RANGE_REGEX = re.compile(r"^(?P<column1>[A-Z]+)(?P<row1>\d+):(?P<column2>[A-Z]+)(?P<row2>\d+)$")


def parse_cell_reference(reference: str) -> tuple[str | None, int, int]:
    """Parse cell reference and return (sheet name, row, column) tuple."""
    match = _CELL_REFERENCE_REGEX.match(reference)
    if not match:
        raise CellReferenceError(f'"{reference}" reference cannot be parsed')

    sheet_name = match.group("sheet")
    row = int(match.group("row"))
    column = _column_to_number(match.group("column"))

    return sheet_name, row, column


def parse_cell_range_reference(reference: str) -> tuple[tuple[int, int], tuple[int, int]]:
    """Parse cell range reference and return ((min_row, min_column), (max_row, max_column)) tuple."""
    match = _RANGE_REGEX.match(reference)
    if not match:
        raise CellReferenceError(f'"{reference}" reference cannot be parsed')

    row1 = int(match.group("row1"))
    column1 = _column_to_number(match.group("column1"))
    row2 = int(match.group("row2"))
    column2 = _column_to_number(match.group("column2"))

    return (
        (min(row1, row2), min(column1, column2)),
        (max(row1, row2), max(column1, column2)),
    )


def _column_to_number(column: str) -> int:
    column_number = 0
    factor = 1
    for letter in reversed(column):
        column_number += (ord(letter) - ord("A") + 1) * factor
        factor *= 26
    return column_number
