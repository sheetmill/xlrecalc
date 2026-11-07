"""A time limit on each request to Excel, and evidence when one passes (``CAPTURING.md`` §2).

    uv run python -m capture.excel_watch     # asks Excel again; the answer goes in the last record

Past its limit a request is stopped, and what can be learned is saved under ``capture/build``: a
``sample`` of Excel, how long it has been up, its CPU and memory, the display's last change, the
request and its file, and whether Excel answers a question about its workbooks now and after 1 and
5 minutes. Excel is never ended: it may hold what the person using it needs. When it answers again,
the workbooks the request named in Excel's container are closed without saving and alerts are
turned back on, so a recovered hang leaves nothing behind. Imports only the standard library, so
that the default test run can import it.

The display is kept awake while a process drives Excel (``caffeinate -d``, from the first request
until the process ends): every hang so far happened with the display off. The record says whether
it was.
"""

from __future__ import annotations

import datetime as dt
import os
import shutil
import subprocess
import sys
import time
from dataclasses import dataclass
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
EVIDENCE = ROOT / "capture" / "build" / "excel_hangs"
ANSWER_SECONDS = 15.0
"""What Excel may take to say how many workbooks it has open."""
ASK_AGAIN = (60.0, 240.0)
"""Seconds to wait before each further question: 1 minute after the first, then 5."""
SAMPLE_SECONDS = 10
_VERSION = 'tell application "Microsoft Excel" to version'
CONTAINER_MARK = "/Library/Containers/com.microsoft.Excel/"
"""Only workbooks in Excel's own container are ever closed: the probes create theirs there."""
_TIDY = (
    "on run argv",
    'tell application "Microsoft Excel"',
    "set closed to {}",
    "repeat with xlName in argv",
    "if exists workbook (xlName as text) then",
    "close workbook (xlName as text) saving no",
    "set end of closed to (xlName as text)",
    "end if",
    "end repeat",
    "set display alerts to true",
    "end tell",
    'set AppleScript\'s text item delimiters to ", "',
    "return closed as text",
    "end run",
)
_awake: subprocess.Popen[bytes] | None = None


def keep_awake() -> None:
    """Keep the display awake until this process ends (``caffeinate -d -w``), on a Mac."""
    global _awake
    if _awake is not None and _awake.poll() is None:
        return
    tool = shutil.which("caffeinate")
    if tool is None:
        return
    _awake = subprocess.Popen(
        [tool, "-d", "-w", str(os.getpid())],
        stdout=subprocess.DEVNULL,
        stderr=subprocess.DEVNULL,
    )


def awake() -> str:
    """Whether the display is being kept awake, for the record."""
    if _awake is not None and _awake.poll() is None:
        return f"display kept awake: yes (caffeinate -d, pid {_awake.pid})"
    return "display kept awake: no"


def tidy(arguments: tuple[str, ...]) -> str:
    """Close the container workbooks a request named, without saving, and turn alerts back on."""
    names = [Path(a).name for a in arguments if a.endswith(".xlsx") and CONTAINER_MARK in a]
    script = [part for line in _TIDY for part in ("-e", line)]
    closed = _output("osascript", *script, *names, seconds=ANSWER_SECONDS)
    return f"after the answer: closed without saving [{closed}]; display alerts set to true"


def answers(seconds: float = ANSWER_SECONDS) -> bool:
    """Whether Excel answers a question about its workbooks (``version`` alone says too little)."""
    question = ["osascript", str(ROOT / "capture" / "array_size.applescript"), "count"]
    try:
        return subprocess.run(question, capture_output=True, timeout=seconds).returncode == 0
    except subprocess.TimeoutExpired:
        return False


def _output(*command: str, seconds: float = 30.0) -> str:
    """What a command printed; a command that fails must not lose the rest of the record."""
    try:
        done = subprocess.run(command, capture_output=True, text=True, timeout=seconds)
    except (OSError, subprocess.SubprocessError) as problem:
        return f"({command[0]} failed: {problem})"
    return (done.stdout or done.stderr or "").strip()


def _state() -> str:
    """Excel's process figures and the display's last change, as text."""
    pid = _output("pgrep", "-x", "Microsoft Excel").split("\n")[0]
    display = [
        line
        for line in _output("pmset", "-g", "log", seconds=60.0).split("\n")
        if "Display is turned" in line
    ]
    lines = [
        f"process: {_output('ps', '-o', 'pid=,etime=,%cpu=,rss=', '-p', pid) if pid else 'none'}",
        "  (pid, elapsed time, CPU %, resident KB)",
        f"version answers: {_output('osascript', '-e', _VERSION, seconds=ANSWER_SECONDS)}",
        f"display, last change: {display[-1] if display else 'none logged'}",
    ]
    return "\n".join(lines)


def _sample(path: Path) -> str:
    pid = _output("pgrep", "-x", "Microsoft Excel").split("\n")[0]
    if not pid:
        return "no Excel process to sample"
    seconds = SAMPLE_SECONDS + 60.0
    _output("sample", pid, str(SAMPLE_SECONDS), "-file", str(path), seconds=seconds)
    return str(path) if path.exists() else "sample wrote no file"


def _log(record: Path, text: str) -> None:
    """Add to the record at once: a long wait may be interrupted."""
    with record.open("a", encoding="utf-8") as file:
        file.write(text.rstrip("\n") + "\n")


def collect(what: str, arguments: tuple[str, ...], seconds: float) -> tuple[Path, bool]:
    """Save the evidence for a request past its limit.

    Returns:
        The record's path, and whether Excel answered: at once it is only slow, and later the hang
        cleared by itself. It waits at most 5 minutes, and stops asking at the first answer.
    """
    EVIDENCE.mkdir(parents=True, exist_ok=True)
    stamp = dt.datetime.now().strftime("%Y%m%d-%H%M%S")
    record = EVIDENCE / f"{stamp}.txt"
    print(f"Excel: {what}. Saving evidence in {record}", file=sys.stderr, flush=True)
    _log(record, f"at {stamp}: {what}\nlimit: {seconds:.0f} s\nrequest: {' '.join(arguments)}")
    answered = answers()
    _log(record, f"answers about workbooks at once: {answered}")
    _log(record, _state())
    _log(record, awake())
    _log(record, f"sample: {_sample(EVIDENCE / f'{stamp}.sample.txt')}")
    waited = 0.0
    for wait in ASK_AGAIN:
        if answered:
            break
        print(f"Excel does not answer: asking again in {wait:.0f} s", file=sys.stderr, flush=True)
        time.sleep(wait)
        waited += wait
        answered = answers()
        _log(record, f"answers after {waited:.0f} s: {answered}")
    if answered:
        _log(record, tidy(arguments))
    return record, answered


@dataclass
class ExcelTimeoutError(Exception):
    """A request to Excel passed its limit."""

    what: str
    answered: bool
    record: Path

    def __str__(self) -> str:
        state = (
            "Excel answers now."
            if self.answered
            else "Excel does not answer, and is left as it is: run `python -m capture.excel_watch` "
            "to ask again before anyone quits it."
        )
        return f"Excel: {self.what}. {state} Evidence: {self.record}"


def osascript(*arguments: str, seconds: float) -> subprocess.CompletedProcess[str]:
    """Run ``osascript`` with a time limit.

    Raises:
        ExcelTimeoutError: past the limit, after the evidence is saved.
    """
    keep_awake()
    try:
        return subprocess.run(
            ["osascript", *arguments], capture_output=True, text=True, timeout=seconds
        )
    except subprocess.TimeoutExpired:
        what = f"{Path(arguments[0]).name} past {seconds:.0f} s"
        record, answered = collect(what, arguments, seconds)
        raise ExcelTimeoutError(what, answered, record) from None


def osascript_or_exit(*arguments: str, seconds: float) -> subprocess.CompletedProcess[str]:
    """The same for a script: past the limit it ends with the message."""
    try:
        return osascript(*arguments, seconds=seconds)
    except ExcelTimeoutError as timeout:
        sys.exit(str(timeout))


def main() -> None:
    """Ask Excel again, and add the answer to the last record."""
    answered = answers()
    records = sorted(p for p in EVIDENCE.glob("*.txt") if not p.name.endswith(".sample.txt"))
    now = dt.datetime.now().strftime("%Y%m%d-%H%M%S")
    if records:
        _log(records[-1], f"asked again at {now}: answers {answered}\n{_state()}")
    print(f"answers about workbooks: {answered}; {records[-1] if records else 'no record'}")


if __name__ == "__main__":
    main()
