"""Steps of the tests that drive Excel, each with a time limit (``CAPTURING.md`` §2).

A step past its limit is stopped. If Excel then does not answer, the run ends here with a message
naming the test, because every later test would wait out its own limit in turn. Excel is left
running: it may hold what the person using it needs, so ending it is theirs to decide.
"""

from __future__ import annotations

import os
import subprocess
from pathlib import Path

import pytest

ROOT = Path(__file__).resolve().parents[1]
STEP_SECONDS = 180.0
"""What one step may take: a large workbook opens, calculates and saves in well under this."""
ANSWER_SECONDS = 15.0
"""What Excel may take to say how many workbooks it has open."""


def answers(seconds: float = ANSWER_SECONDS) -> bool:
    """Whether Excel answers a question about its workbooks (``version`` alone says too little)."""
    question = ["osascript", str(ROOT / "capture" / "array_size.applescript"), "count"]
    try:
        return subprocess.run(question, capture_output=True, timeout=seconds).returncode == 0
    except subprocess.TimeoutExpired:
        return False


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
    command = ["osascript", *arguments]
    try:
        return subprocess.run(command, capture_output=True, text=True, timeout=seconds)
    except subprocess.TimeoutExpired:
        what = f"{Path(arguments[0]).name} past {seconds:.0f} s"
        if not answers():
            stop_the_run(what)
        raise AssertionError(f"Excel took too long: {what}") from None
