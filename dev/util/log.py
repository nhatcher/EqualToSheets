from __future__ import annotations

import sys
from contextlib import contextmanager
from copy import deepcopy
from typing import Any, Callable, Iterator

verbose = False

# the current log stack
log_stack: list[str] = []

# The 'type' (info / error) and log stack of the previously printed message
prev_log: dict[str, Any] = {"type": None, "log_stack": None}


class LogError(Exception):
    """Exception raised by the logging system."""


def set_verbose(ver: bool) -> None:
    """Turn on / off verbose mode for the logging system."""
    global verbose  # noqa: WPS420
    verbose = ver  # noqa: WPS442


def is_verbose() -> bool:
    """Returns true iff the logging system has verbose mode enabled."""
    return verbose


def info(text: str) -> None:
    """Log an info message."""
    global prev_log  # noqa: WPS420
    new_log = {"type": "info", "log_stack": log_stack}
    if is_verbose():
        if prev_log != new_log:
            # This is a new log type / stack, so we print the 'stack trace' line
            print("INFO {0}:".format(get_stack()))
            prev_log = deepcopy(new_log)  # noqa: WPS442
        print("    {0}".format(text))


def info_list(text_list: list[str]) -> None:
    """Log a list of info messages."""
    for text in text_list:
        info(text)


def error(text: str) -> None:
    """Log an error message."""
    global prev_log  # noqa: WPS420
    new_log = {"type": "error", "log_stack": log_stack}
    if prev_log != new_log:
        # This is a new log type / stack, so we print the 'stack trace' line
        print("ERROR {0}:".format(get_stack()), file=sys.stderr)
        prev_log = deepcopy(new_log)  # noqa: WPS442
    print("    {0}".format(text))


def error_list(text_list: list[str]) -> None:
    """Log a list of error messages."""
    for text in text_list:
        error(text)


def push(stack_entry: str) -> None:
    """Add an entry to the log context."""
    log_stack.append(stack_entry)


def push_info(stack_entry: str, text: str) -> None:
    """Add an entry to the log context and log a message."""
    push(stack_entry)
    info(text)


def get_stack() -> str:
    """Return a textual representation of the log context."""
    return " ".join(log_stack)


def pop() -> str:
    """Remove and return the most recent addition to the log context."""
    if not log_stack:
        raise LogError("Empty log stack - cannot pop()")
    return log_stack.pop()


@contextmanager
def context(stack_entry: str) -> Iterator[None]:
    """
    Create a log context for use in a `with` statement.

    For example:
        with context("<context text>"):
            <code>
    """
    push(stack_entry)
    try:
        yield
    finally:
        pop()


def func_context(func: Callable[..., Any]) -> Callable[..., Any]:
    """Decorator to add a function to the log context."""

    def wrapper(*args: list[Any], **kwargs: dict[str, Any]) -> Callable[..., Any]:  # noqa: WPS430
        with context("{0}()".format(func.__name__)):
            return func(*args, **kwargs)

    return wrapper
