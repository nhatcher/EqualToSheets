from __future__ import annotations

import threading
from collections import defaultdict
from contextlib import ContextDecorator
from typing import TYPE_CHECKING, Any

from equalto._equalto import WorkbookError

if TYPE_CHECKING:
    from equalto.workbook import Workbook

# Pretend this exception was defined here, thanks to that it's displayed as
# "equalto.exceptions.WorkbookError" instead of "_equalto.WorkbookError" in the stack trace.
WorkbookError.__module__ = __name__

# Make it possible to import _equalto.WorkbookError directly from equalto.exceptions.
__all__ = [
    "WorkbookError",
    "CellReferenceError",
    "WorkbookValueError",
    "WorkbookEvaluationError",
    "SuppressEvaluationErrors",
]


class CellReferenceError(WorkbookError):
    """Invalid cell reference error."""


class WorkbookValueError(WorkbookError, ValueError):
    """Workbook value error."""


class WorkbookEvaluationError(WorkbookError):
    """Could not evaluate some of the cells (i.e. circular dependency, unsupported function)."""


class _LocalContext(threading.local):
    def __init__(self) -> None:
        self.in_context = False
        self.suppressed_errors: dict[Workbook, list[str]] = defaultdict(list)


class SuppressEvaluationErrors(ContextDecorator):
    context = _LocalContext()

    def __enter__(self) -> SuppressEvaluationErrors:
        assert not self.context.in_context, "already in SuppressEvaluationErrors context"
        self.context.in_context = True
        return self

    def __exit__(self, *exc: Any) -> None:
        self.context.in_context = False
        self.context.suppressed_errors = defaultdict(list)

    @classmethod
    def in_context(cls: type[SuppressEvaluationErrors]) -> bool:
        return cls.context.in_context > 0

    @classmethod
    def log_errors(cls: type[SuppressEvaluationErrors], workbook: Workbook, error: list[str]) -> None:
        cls.context.suppressed_errors[workbook].extend(error)

    def suppressed_errors(self, workbook: Workbook) -> list[str]:
        return self.context.suppressed_errors[workbook]
