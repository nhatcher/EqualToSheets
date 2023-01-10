from __future__ import annotations

from datetime import tzinfo

from _pycalc import create, load_excel

from equalto.workbook import Workbook


def load(workbook_path: str) -> Workbook:
    """Load a workbook from the file."""
    # TODO: Shouldn't rust recognize the locale and time zone?
    return Workbook(load_excel(workbook_path, "en-US", "UTC"))


def new(locale: str, tz: tzinfo) -> Workbook:
    """Create a new workbook."""
    return Workbook(create("workbook", locale, str(tz)))
