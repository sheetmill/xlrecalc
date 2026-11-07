"""Steps of the tests that drive Excel, each with a time limit (``CAPTURING.md`` §2).

A step past its limit is stopped and the evidence saved (``capture/excel_watch.py``). If Excel then
does not answer, the run ends here with a message naming the test, because every later test would
wait out its own limit in turn. Excel is left running: it may hold what the person using it needs,
so ending it is theirs to decide.
"""

from __future__ import annotations

import os
import subprocess

import pytest
from capture.excel_watch import ExcelTimeoutError, answers, osascript

__all__ = ["answers", "step", "stop_the_run"]

STEP_SECONDS = 180.0
"""What one step may take: a large workbook opens, calculates and saves in well under this."""


def stop_the_run(what: str) -> None:
    """End the whole test run, saying which test was running."""
    test = os.environ.get("PYTEST_CURRENT_TEST", "no test").split(" ")[0]
    pytest.exit(
        f"Excel stopped answering during {test} ({what}). The run ends here. Excel is still "
        "running, possibly with alerts off: look at it, and quit it yourself if need be.",
        returncode=3,
    )


def step(*arguments: str, seconds: float = STEP_SECONDS) -> subprocess.CompletedProcess[str]:
    """Run ``osascript`` with a time limit; the run ends if Excel then does not answer."""
    try:
        return osascript(*arguments, seconds=seconds)
    except ExcelTimeoutError as timeout:
        if not timeout.answered:
            stop_the_run(f"{timeout.what}; evidence in {timeout.record}")
        raise AssertionError(f"Excel took too long: {timeout}") from None
