from __future__ import annotations

from typing import TypedDict

from sheet_ai.completion import create_completion


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


def generate_workbook_data(prompt: list[str]) -> WorkbookData:
    completion = create_completion(prompt)
    # TODO: Retry logic when we cannot process given response or it contains invalid formulas
    return [
        [{"input": cell} for cell in map(str.strip, row.split("|")[1:-1])]
        for row in completion.split("\n")
        if row.strip() != ""
    ]
