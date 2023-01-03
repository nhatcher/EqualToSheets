import traceback
from datetime import timezone

import equalto


def test_workbook_error_traceback() -> None:
    try:
        equalto.new("en-US", timezone.utc).sheets.add("Sheet1")
    except Exception as err:
        assert err.__module__ == "equalto.exceptions"
        last_traceback_line = traceback.format_exc().strip().split("\n")[-1]
        assert last_traceback_line == f"equalto.exceptions.WorkbookError: {str(err)}"
