import json
import os
import re
from pathlib import Path
from typing import Any

import pytest

from sheet_ai.exceptions import WorkbookProcessingError
from sheet_ai.workbook import generate_workbook_data


@pytest.mark.parametrize("markdown_path", (Path(os.path.dirname(__file__)) / "workbooks" / "correct").glob("*.md"))
def test_generate_workbook_data(markdown_path: Path, monkeypatch: Any) -> None:
    markdown = markdown_path.read_text().strip()
    expected_workbook = json.loads(markdown_path.with_suffix(".json").read_text())
    with monkeypatch.context() as mock:
        mock.setattr("sheet_ai.workbook.create_completion", lambda _: markdown)
        assert generate_workbook_data([], retries=0) == expected_workbook


@pytest.mark.parametrize("markdown_path", (Path(os.path.dirname(__file__)) / "workbooks" / "error").glob("*.md"))
def test_generate_workbook_data_error(markdown_path: Path, monkeypatch: Any) -> None:
    markdown = markdown_path.read_text().strip()
    error = f"^{re.escape(markdown_path.with_suffix('.error').read_text())}$"
    with monkeypatch.context() as mock:
        mock.setattr("sheet_ai.workbook.create_completion", lambda _: markdown)
        with pytest.raises(WorkbookProcessingError, match=error):
            assert generate_workbook_data([], retries=0)


def test_retry_logic_with_eventual_success(monkeypatch: Any) -> None:
    markdowns = iter(["invalid", "invalid", "invalid", "|valid|"])
    monkeypatch.setattr("sheet_ai.workbook.create_completion", lambda _: next(markdowns))

    assert generate_workbook_data([]) == [[{"input": "valid"}]]


def test_retry_logic_with_failure(monkeypatch: Any) -> None:
    markdowns = iter(["invalid", "invalid", "invalid", "|valid|"])
    monkeypatch.setattr("sheet_ai.workbook.create_completion", lambda _: next(markdowns))

    # with 2 retries, the "valid" response won't be reached
    with pytest.raises(WorkbookProcessingError, match="Could not find a workbook"):
        assert generate_workbook_data([], retries=2)
