class SheetAIError(Exception):
    """Generic Sheet AI error."""


class CompletionError(SheetAIError):
    """Could not generate a completion."""


class WorkbookProcessingError(SheetAIError):
    """Could not generate a workbook."""
