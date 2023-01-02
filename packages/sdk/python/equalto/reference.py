from __future__ import annotations

import re

from equalto.exceptions import CellReferenceError

_CELL_REFERENCE_REGEX = re.compile(r"^(?:(?P<sheet>[^!]+)!)?(?P<column>[A-Z]+)(?P<row>\d+)$")


def parse_cell_reference(reference: str) -> tuple[str | None, int, int]:
    """Parse cell reference and return (sheet name, row, column) tuple."""
    match = _CELL_REFERENCE_REGEX.match(reference)
    if not match:
        raise CellReferenceError(f'"{reference}" reference cannot be parsed')

    sheet_name = match.group("sheet")
    row = int(match.group("row"))
    column = _column_to_number(match.group("column"))

    return sheet_name, row, column


def _column_to_number(column: str) -> int:
    column_number = 0
    factor = 1
    for letter in reversed(column):
        column_number += (ord(letter) - ord("A") + 1) * factor
        factor *= 26
    return column_number
