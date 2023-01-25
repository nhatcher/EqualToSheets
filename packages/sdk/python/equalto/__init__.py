from __future__ import annotations

from datetime import datetime, tzinfo
from importlib.metadata import version

from equalto._equalto import create, load_excel
from equalto.workbook import Workbook

__version__ = version("equalto")


def load(workbook_path: str) -> Workbook:
    """Load a workbook from the file."""
    # TODO: Shouldn't rust recognize the locale and time zone?
    # TODO: If rust can't recognize the time zone, should we use local time zone or UTC by default?
    model = load_excel(workbook_path, "en", "UTC")
    model.evaluate()
    return Workbook(model)


def new(*, timezone: tzinfo | None = None) -> Workbook:
    """Create a new workbook."""
    return Workbook(create("workbook", "en", str(timezone or _get_local_tz())))


def _get_local_tz() -> tzinfo:
    tz = datetime.now().astimezone().tzinfo
    assert tz is not None
    return tz
