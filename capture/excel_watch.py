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
until the process ends): every unexplained hang so far happened with the display off. The record
says whether it was.

Before a process's first request, a window of Excel's own open on screen beside its workbooks (the
About window, Preferences, a dialog) stops the run with a message: such a window blocks every
request about workbooks, and one left open caused a hang that looked like the others.
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
_display_at_start = "not checked"


def _display_line() -> str:
    """The display's last logged change (``pmset -g log``), or "none logged"."""
    lines = [
        line
        for line in _output("pmset", "-g", "log", seconds=60.0).split("\n")
        if "Display is turned" in line
    ]
    return lines[-1] if lines else "none logged"


def keep_awake() -> None:
    """Keep the display awake until this process ends (``caffeinate -d -w``), on a Mac.

    ``-d`` keeps a display on but does not wake one that is off, so the first call also declares
    user activity once (``caffeinate -u -t 1``), and notes the display's state before it.
    """
    global _awake, _display_at_start
    if _awake is not None and _awake.poll() is None:
        return
    tool = shutil.which("caffeinate")
    if tool is None:
        return
    _display_at_start = _display_line()
    _output(tool, "-u", "-t", "1", seconds=10.0)
    _awake = subprocess.Popen(
        [tool, "-d", "-w", str(os.getpid())],
        stdout=subprocess.DEVNULL,
        stderr=subprocess.DEVNULL,
    )


_WINDOWS_JXA = r"""
ObjC.import("CoreGraphics");
const all = $.CGWindowListCopyWindowInfo($.kCGWindowListOptionOnScreenOnly, 0);
const list = ObjC.castRefToObject(all);
const out = [];
for (let i = 0; i < list.count; i++) {
  const w = list.objectAtIndex(i);
  if (ObjC.unwrap(w.objectForKey("kCGWindowOwnerName")) !== "Microsoft Excel") continue;
  if (ObjC.unwrap(w.objectForKey("kCGWindowLayer")) !== 0) continue;
  const b = ObjC.deepUnwrap(w.objectForKey("kCGWindowBounds"));
  const title = ObjC.unwrap(w.objectForKey("kCGWindowName")) || "";
  out.push([Math.round(b.Width), Math.round(b.Height), title].join("\t"));
}
out.join("\n");
"""
_DOCUMENT_WINDOWS = 'tell application "Microsoft Excel" to count of windows'
_checked = False


def excel_windows() -> list[str]:
    """Excel's windows on screen, as ``width x height`` and a title where macOS gives one."""
    try:
        result = subprocess.run(
            ["osascript", "-l", "JavaScript", "-e", _WINDOWS_JXA],
            capture_output=True, text=True, timeout=10.0,
        )  # fmt: skip
    except subprocess.TimeoutExpired:
        return []
    text = result.stdout if result.returncode == 0 else ""
    out = []
    for line in filter(None, text.splitlines()):
        width, height, title = [*line.split("\t"), "", ""][:3]
        out.append(f"{width} x {height}" + (f" ({title})" if title else ""))
    return out


def preflight() -> None:
    """Stop before a run when a window of Excel's own blocks its requests, once per process.

    Excel answers nothing about workbooks while a modal window is open (the About window left
    open hung a run after 78 opens). The windows on screen are compared with the document windows
    Excel reports; one more, or no answer, names the windows and ends the process.
    """
    global _checked
    if _checked or shutil.which("osascript") is None:
        return
    _checked = True
    if subprocess.run(["pgrep", "-x", "Microsoft Excel"], capture_output=True).returncode:
        return  # not running: the first request starts it
    on_screen = excel_windows()
    try:
        result = subprocess.run(
            ["osascript", "-e", _DOCUMENT_WINDOWS], capture_output=True, text=True, timeout=10.0
        )
        documents = int(result.stdout.strip()) if result.returncode == 0 else None
    except (subprocess.TimeoutExpired, ValueError):
        documents = None
    if documents is not None and len(on_screen) <= documents:
        return
    seen = ", ".join(on_screen) or "none listed"
    state = "does not answer" if documents is None else f"has {documents} workbook windows"
    sys.exit(
        f"Excel {state}, with these windows on screen: {seen}. A window of Excel's own (About "
        "Microsoft Excel, Preferences, a dialog) blocks every request: close it and run again."
    )


def awake() -> str:
    """Whether the display is being kept awake, for the record."""
    start = f"; display at the start of the run, last change: {_display_at_start}"
    if _awake is not None and _awake.poll() is None:
        return f"display kept awake: yes (caffeinate -d, pid {_awake.pid}){start}"
    return f"display kept awake: no{start}"


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
    display = _display_line()
    lines = [
        f"process: {_output('ps', '-o', 'pid=,etime=,%cpu=,rss=', '-p', pid) if pid else 'none'}",
        "  (pid, elapsed time, CPU %, resident KB)",
        f"version answers: {_output('osascript', '-e', _VERSION, seconds=ANSWER_SECONDS)}",
        f"display, last change: {display}",
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
    preflight()
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
