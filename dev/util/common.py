from __future__ import annotations

import re
import subprocess  # noqa: S404
from typing import Any

from util import log

# regexp to remove ANSI control chars
# Taken from https://stackoverflow.com/questions/14693701/
ansi_control_re = re.compile(
    r"""
    \x1B  # ESC
    (?:   # 7-bit C1 Fe (except CSI)
        [@-Z\\-_]
    |     # or [ for CSI, followed by a control sequence
        \[
        [0-?]*  # Parameter bytes
        [ -/]*  # Intermediate bytes
        [@-~]   # Final byte
    )
""",
    re.VERBOSE,
)


class RunError(Exception):
    """Indicates that the process exited with some sort of error."""

    def __init__(self, message: str = "", exit_code: int | None = None) -> None:
        """
        Constructor for RunError object.

        Args:
            message: description of the error.
            exit_code: the exit code of the process (if available).
        """
        super().__init__(message)
        self.exit_code = exit_code


def _log_message(line_no_ansi: str, log_policy: str) -> None:
    if log_policy == "all-error":
        log.error(line_no_ansi)
    elif log_policy in {"all-info-unless-error", "stderr-is-info-unless-error"}:
        log.info(line_no_ansi)
    else:
        raise RunError("Invalid log_policy: {0}.".format(log_policy))


@log.func_context
def run(  # noqa: WPS231, C901
    args: list[str],
    cwd: Any = None,
    env: Any = None,
    log_policy: str = "stderr-is-info-unless-error",
) -> None:
    """
    Runs `args` using `subprocess.Popen()`.

    The output is captured, and ANSI control characters are removed before being
    logged.

    Args:
        args: entire command (with parameters) to execute
        cwd: current working directory
        env: environment data structure
        log_policy: how do we log stdout / stderr messages?

    Raises:
        RunError: an error occurred running the process.

    Note on `log_policy`:
        The `log_policy` can be one of these values:
            * "stderr-is-info-unless-error": (default) both stdout and stderr messages
                are logged as INFO messages, unless the process exits with an error,
                in which case the stderr messages are re-logged as ERROR
                messages
            * "all-info-unless-error": both stdout and stderr messages are
                logged as INFO messages, unless the process exits with an error,
                in which case both the stdout and stderr messages are re-logged
                as ERROR messages
            * "all-error": stdout and stderr messages area always logged as
                ERROR messages.
    """
    log_policies = {
        "all-info-unless-error",
        "stderr-is-info-unless-error",
        "all-error",
    }
    if log_policy not in log_policies:
        raise RunError("Invalid log_policy: {0}.".format(log_policy))

    stderr_messages = []
    stdout_messages = []

    save_stderr = log_policy in {
        "stderr-is-info-unless-error",
        "all-info-unless-error",
    }
    save_stdout = log_policy == "all-info-unless-error"

    log_context = '"{0}"'.format(" ".join(args))
    if cwd is not None:
        log_context += ', cwd="{0}"'.format(cwd)

    with log.context(log_context):
        log.info("calling subprocess.run()")
        proc = subprocess.Popen(args, cwd=cwd, env=env)  # noqa: S603
        while True:
            rc = proc.poll()
            (stdout_data, stderr_data) = proc.communicate()

            stdout_lines = stdout_data.decode("utf-8").split(r"\n") if stdout_data else []
            for line_stdout in stdout_lines:
                line_no_ansi = ansi_control_re.sub("", line_stdout.rstrip())
                if line_no_ansi.strip() != "":
                    _log_message(line_no_ansi, log_policy)
                    if save_stdout:
                        stdout_messages.append(line_no_ansi)  # noqa: WPS220

            stderr_lines = stderr_data.decode("utf-8").split(r"\n") if stderr_data else []
            for line_stderr in stderr_lines:
                _log_message(line_stderr, log_policy)
                line_no_ansi = ansi_control_re.sub("", line_stderr.rstrip())
                if line_no_ansi.strip() != "":
                    _log_message(line_no_ansi, log_policy)
                    if save_stderr:
                        stderr_messages.append(line_no_ansi)  # noqa: WPS220

            proc_exit = rc is not None
            proc_exit_success = proc_exit and proc.returncode == 0
            proc_exit_failure = proc_exit and proc.returncode != 0

            if proc_exit_failure:
                if log_policy == "all-info-unless-error":  # noqa: WPS220
                    log.error_list(stdout_messages)
                    log.error_list(stderr_messages)
                elif log_policy == "stderr-is-info-unless-error":  # noqa: WPS220
                    log.error_list(stderr_messages)
                elif log_policy == "all-error":  # noqa: WPS220
                    pass  # noqa: WPS420
                else:
                    raise RunError("Invalid log_policy: {0}.".format(log_policy))

                log.error("Process exited with a non-zero value ({0})".format(proc.returncode))
                raise RunError("Process exited with a non-zero value", proc.returncode)

            if proc_exit_success:
                break


@log.func_context
def run_str(
    cmd: str,
    cwd: Any = None,
    env: Any = None,
    log_policy: str = "all-info-unless-error",
) -> None:
    """
    Runs `cmd` using subprocess.run().

    WARNING: ensure you properly sanitize `cmd`.

    This functions uses subprocess.run() to execute cmd.split(" "). This is only robust
    for certain values of cmd (ie: those that are careful to only use a single space to
    separate the parameters to the executable).

    For an explanation of the various parameters, see `run()`.
    """
    run(cmd.split(" "), cwd=cwd, env=env, log_policy=log_policy)
