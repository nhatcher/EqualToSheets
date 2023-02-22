from __future__ import annotations

import re
from typing import Iterable, TypedDict

import equalto
from equalto.exceptions import WorkbookError

from sheet_ai.completion import create_completion
from sheet_ai.exceptions import WorkbookProcessingError

DEFAULT_RETRIES = 5

ERROR_TEMPLATE = """
Could not process completion

Completion:
{completion}

Error:
{error}
""".strip()


class Style(TypedDict, total=False):
    bold: bool
    num_fmt: str


CellInput = TypedDict(
    "CellInput",
    {
        "input": "str | float | bool | None",
        "style": Style,
    },
    total=False,
)


WorkbookData = list[list[CellInput]]


def generate_workbook_data(prompt: list[str], retries: int = DEFAULT_RETRIES) -> WorkbookData:
    completion = create_completion(prompt)
    try:
        return _process_completion(completion)
    except WorkbookProcessingError as err:
        if retries > 0:
            # TODO: We could try rerunning the same prompt, but we could also consider:
            #         - preparing a few alternative preambles for subsequent calls;
            #         - use different AI model temperature values.
            return generate_workbook_data(prompt, retries - 1)
        raise WorkbookProcessingError(ERROR_TEMPLATE.format(completion=completion, error=str(err)))


def _process_completion(completion: str) -> WorkbookData:
    workbook_data = _get_initial_workbook_data(completion)

    sheet = equalto.new().sheets[0]
    for row, col, cell_input in _get_workbook_data_iter(workbook_data):
        try:
            sheet.cell(row, col).set_user_input(str(cell_input["input"]))
        except WorkbookError as err:
            raise WorkbookProcessingError(f"Generated workbook contains an error\n{err}")

    for row, col, cell_input in _get_workbook_data_iter(workbook_data):
        cell = sheet.cell(row, col)
        if cell.formula and cell.value in {"#VALUE!", "#REF!", "#NAME?"}:
            raise WorkbookProcessingError(
                f"Encountered formula resulting in {cell.value}\n{cell.text_ref}{cell.formula}",
            )

        new_cell_input = cell.formula or cell.value
        assert new_cell_input is None or isinstance(new_cell_input, (float, bool, str))
        cell_input["input"] = new_cell_input

        num_fmt = cell.style.format
        if num_fmt != "general":
            cell_input.setdefault("style", {})
            cell_input["style"]["num_fmt"] = num_fmt

    return workbook_data


def _get_workbook_data_iter(workbook_data: WorkbookData) -> Iterable[tuple[int, int, CellInput]]:
    for row_index, row in enumerate(workbook_data, 1):
        for col_index, cell_input in enumerate(row, 1):
            yield row_index, col_index, cell_input


def _get_initial_workbook_data(completion: str) -> WorkbookData:
    try:
        workbook_data = [
            [_get_initial_cell_input(cell.strip()) for cell in row.split("|")[1:-1]]
            for row in map(str.strip, completion.split("\n"))
            if len(row) > 1 and row[0] == row[-1] == "|"
        ]
    except Exception as err:
        raise WorkbookProcessingError(str(err))

    if not workbook_data:
        raise WorkbookProcessingError("Could not find a workbook")

    # add empty cells so that all rows are of the same length
    max_row_len = max(map(len, workbook_data))
    for row in workbook_data:
        row += [_get_empty_cell() for _ in range(max_row_len - len(row))]

    # if the last row starts with "Total" cell, use style.bold = True in all cells in the row
    last_row = workbook_data[-1]
    first_cell = last_row[0]
    if isinstance(first_cell["input"], str) and first_cell["input"].lower() == "total":
        for cell in last_row:
            cell.setdefault("style", {})
            cell["style"]["bold"] = True

    return workbook_data


def _get_empty_cell() -> CellInput:
    return {"input": ""}


def _get_initial_cell_input(cell_value: str) -> CellInput:
    if match := re.match(r"\*\*(?P<value>.+)\*\*", cell_value):
        return {"input": match.group("value"), "style": {"bold": True}}
    return {"input": cell_value}
