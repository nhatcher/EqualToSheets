from __future__ import annotations

from datetime import tzinfo

from _pycalc import create, load_excel

from equalto.exceptions import WorkbookError
from equalto.workbook import Workbook


def load(workbook_path: str) -> Workbook:
    """Load a workbook from the file."""
    try:
        # TODO: Shouldn't rust recognize the locale and time zone?
        return Workbook(load_excel(workbook_path, "en-US", "UTC"))
    except BaseException as err:  # noqa: WPS424
        raise _process_pycalc_error(err)


def new(locale: str, tz: tzinfo) -> Workbook:
    """Create a new workbook."""
    try:
        return Workbook(create("workbook", locale, str(tz)))
    except BaseException as err:  # noqa: WPS424
        raise _process_pycalc_error(err)


def _process_pycalc_error(err: BaseException) -> BaseException | WorkbookError:
    # TODO: PanicException is a BaseException (which shouldn't really be caught) and this comparison
    #       by the exception name is a bit flaky. Perhaps we should fix the pycalc code so that it
    #       doesn't panic and raises non-system exiting exceptions instead.
    if isinstance(err, Exception) or type(err).__name__ == "PanicException":
        return WorkbookError(str(err))
    return err
