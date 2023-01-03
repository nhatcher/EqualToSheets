from _pycalc import WorkbookError

# Pretend this exception was defined here, thanks to that it's displayed as
# "equalto.exceptions.WorkbookError" instead of "_pycalc.WorkbookError" in the stack trace.
WorkbookError.__module__ = __name__

# Make it possible to import _pycalc.WorkbookError directly from equalto.exceptions.
__all__ = [
    "WorkbookError",
    "CellReferenceError",
    "WorkbookValueError",
]


class CellReferenceError(WorkbookError):
    """Invalid cell reference error."""


class WorkbookValueError(WorkbookError, ValueError):
    """Workbook value error."""
