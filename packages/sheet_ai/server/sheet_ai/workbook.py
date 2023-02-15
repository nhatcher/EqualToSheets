from __future__ import annotations

import re
from typing import TypedDict

from sheet_ai.completion import create_completion
from sheet_ai.exceptions import WorkbookProcessingError

DEFAULT_RETRIES = 5


class Style(TypedDict, total=False):
    bold: bool
    num_fmt: str


Cell = TypedDict(
    "Cell",
    {
        "input": "str | float | bool | None",
        "style": Style,
    },
    total=False,
)


WorkbookData = list[list[Cell]]


def generate_workbook_data(prompt: list[str], retries: int = DEFAULT_RETRIES) -> WorkbookData:
    completion = create_completion(prompt)
    try:
        return _process_completion(completion)
    except WorkbookProcessingError:
        if retries > 0:
            # TODO: We could try rerunning the same prompt, but we could also consider:
            #         - preparing a few alternative preambles for subsequent calls;
            #         - use different AI model temperature values.
            return generate_workbook_data(prompt, retries - 1)
        raise


def _process_completion(completion: str) -> WorkbookData:
    # TODO: This obviously needs to be more sophisticated:
    #         - we need to load the Workbook using EqualTo Calc and confirm that seems to be valid;
    #         - we need to parse values and process styles.
    try:
        return [
            [_get_cell(cell) for cell in map(str.strip, row.split("|")[1:-1])]
            for row in completion.split("\n")
            if row.strip() != ""
        ]
    except Exception:
        raise WorkbookProcessingError("Could not process workbook")


def _get_cell(cell_value: str) -> Cell:
    if match := re.match(r"\*\*(?P<value>.+)\*\*", cell_value):
        return {"input": match.group("value"), "style": {"bold": True}}
    return {"input": cell_value}
