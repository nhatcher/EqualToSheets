class WorkbookError(Exception):
    """A generic workbook error."""


class CellReferenceError(WorkbookError):
    """Invalid cell reference error."""
