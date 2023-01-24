from equalto._equalto import WorkbookError

# Pretend this exception was defined here, thanks to that it's displayed as
# "equalto.exceptions.WorkbookError" instead of "_equalto.WorkbookError" in the stack trace.
WorkbookError.__module__ = __name__

# Make it possible to import _equalto.WorkbookError directly from equalto.exceptions.
__all__ = [
    "WorkbookError",
    "CellReferenceError",
    "WorkbookValueError",
]


class CellReferenceError(WorkbookError):
    """Invalid cell reference error."""


class WorkbookValueError(WorkbookError, ValueError):
    """Workbook value error."""
