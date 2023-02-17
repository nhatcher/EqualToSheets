import json
import os
from pathlib import Path
from typing import Any

import pytest

from sheet_ai.workbook import generate_workbook_data


@pytest.mark.parametrize("markdown_path", (Path(os.path.dirname(__file__)) / "workbooks").glob("*.md"))
def test_generate_workbook_data(markdown_path: Path, monkeypatch: Any) -> None:
    markdown = markdown_path.read_text().strip()
    expected_workbook = json.loads(markdown_path.with_suffix(".json").read_text())
    with monkeypatch.context() as mock:
        mock.setattr("sheet_ai.workbook.create_completion", lambda _: markdown)
        assert generate_workbook_data([]) == expected_workbook
