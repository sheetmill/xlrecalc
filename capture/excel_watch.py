"""A time limit on each request to Excel, and evidence when one passes (``CAPTURING.md`` §2).

    uv run python -m capture.excel_watch [RECORD]   # asks Excel again; the answer goes in RECORD

Past its limit a request is stopped, and so is one that AppleScript gave up on first (``AppleEvent
timed out``, -1712), the same hang seen from the other side; what can be learned is saved under
``capture/build``: a ``sample`` of Excel, how long it has been up, its CPU and memory, the
display's last change, the request and its file, and whether Excel answers a question about its
workbooks now and after 1 and 5 minutes. Excel is ended only when nothing of anyone else's can be
in it (``may_end``): the run began with no workbook open, and every window of Excel's on screen is
the request's own. Otherwise it is left as it is, since it may hold what the person using it
needs. When it answers again, the workbooks the request named in Excel's container are closed
without saving and alerts are turned back on, so a recovered hang leaves nothing behind. Imports
only the standard library, so that the default test run can import it.

The display is kept awake while a process drives Excel (``caffeinate -d``, from the first request
until the process ends): each unexplained hang recorded came with the display off. The record
says whether it was.

Before a process's first request, a window of Excel's own open on screen beside its workbooks (the
About window, Preferences, a dialog) stops the run with a message: such a window blocks every
request about workbooks, and the run would look like an unexplained hang.
"""

from __future__ import annotations

import datetime as dt
import os
import shutil
import signal
import subprocess
import sys
import time
from dataclasses import dataclass
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
EVIDENCE = ROOT / "capture" / "build" / "excel_hangs"
ANSWER_SECONDS = 15.0
"""The limit on a short question to Excel: its workbook count, its version, closing workbooks."""
END_GRACE_SECONDS = 10.0
"""Between TERM and KILL, when Excel is ended."""
ASK_AGAIN = (60.0, 240.0)
"""The waits before each further question about workbooks, after the first at once: Excel is
asked again after 1 minute, and after 5 minutes in all."""
RECENT_SECONDS = 1800.0
"""Asking again adds the answer to the latest record only if that record is this recent."""
APPLE_EVENT_TIMEOUT = "(-1712)"
"""AppleScript's error when it gives up on a request, which can come before our limit."""
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
_began_empty: bool | None = None  # whether this process's run began with no workbook in Excel


def _window_list() -> list[tuple[str, str, str]] | None:
    """Excel's windows on screen as width, height and title (empty where macOS gives none).

    None when the window server does not answer: then nothing is known about them.
    """
    try:
        result = subprocess.run(
            ["osascript", "-l", "JavaScript", "-e", _WINDOWS_JXA],
            capture_output=True, text=True, timeout=10.0,
        )  # fmt: skip
    except subprocess.TimeoutExpired:
        return None
    if result.returncode:
        return None
    return [
        (width, height, title)
        for width, height, title in (
            [*line.split("\t"), "", ""][:3] for line in result.stdout.splitlines() if line
        )
    ]


def excel_windows() -> list[str]:
    """Excel's windows on screen, as ``width x height`` and a title where macOS gives one."""
    return [
        f"{width} x {height}" + (f" ({title})" if title else "")
        for width, height, title in _window_list() or []
    ]


def may_end(
    began_empty: bool | None, windows: list[tuple[str, str, str]] | None, names: list[str]
) -> bool:
    """Whether a request past its limit may end Excel: nothing of anyone else's can be in it.

    The run began with Excel not running or holding no workbook (``began_empty``), and every
    window of Excel's on screen is the request's own: no more of them than the workbooks it named,
    and each title macOS gives is one of their names. Anything unknown is a no.
    """
    if began_empty is not True or windows is None or not names:
        return False
    if len(windows) > len(names):
        return False
    stems = {Path(name).stem for name in names} | set(names)
    return all(not title or title in stems for _, _, title in windows)


def end_excel(record: Path) -> bool:
    """End Excel (TERM, then KILL after a grace period), in the record. Whether it has ended."""
    found = subprocess.run(["pgrep", "-x", "Microsoft Excel"], capture_output=True, text=True)
    pids = [int(p) for p in found.stdout.split()]
    if len(pids) != 1:
        _log(record, f"not ended: {len(pids)} processes named Microsoft Excel")
        return False
    pid = pids[0]
    os.kill(pid, signal.SIGTERM)
    deadline = time.monotonic() + END_GRACE_SECONDS
    while time.monotonic() < deadline and _alive(pid):
        time.sleep(0.5)
    how = "TERM"
    if _alive(pid):
        os.kill(pid, signal.SIGKILL)
        how = f"TERM, then KILL after {END_GRACE_SECONDS:.0f} s"
        time.sleep(1.0)
    ended = not _alive(pid)
    _log(record, f"ended Excel (pid {pid}) by {how}: {'ended' if ended else 'still running'}")
    global _checked
    _checked = False  # the next request checks Excel's windows again (a recovery window, say)
    return ended


def _alive(pid: int) -> bool:
    try:
        os.kill(pid, 0)
    except ProcessLookupError:
        return False
    return True


def preflight() -> None:
    """Stop before a run when a window of Excel's own blocks its requests, once per process.

    Excel answers nothing about workbooks while a modal window is open (the About window,
    Preferences, a dialog). The windows on screen are compared with the document windows
    Excel reports; one more, or no answer, names the windows and ends the process.
    """
    global _checked, _began_empty
    if _checked or shutil.which("osascript") is None:
        return
    _checked = True
    if subprocess.run(["pgrep", "-x", "Microsoft Excel"], capture_output=True).returncode:
        if _began_empty is None:
            _began_empty = True
        return  # not running: the first request starts it
    on_screen = excel_windows()
    try:
        result = subprocess.run(
            ["osascript", "-e", _DOCUMENT_WINDOWS], capture_output=True, text=True, timeout=10.0
        )
        documents = int(result.stdout.strip()) if result.returncode == 0 else None
    except (subprocess.TimeoutExpired, ValueError):
        documents = None
    if _began_empty is None:
        _began_empty = documents == 0
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


def collect(
    what: str, arguments: tuple[str, ...], seconds: float, end_allowed: bool = False
) -> tuple[Path, bool, bool]:
    """Save the evidence for a request past its limit.

    Returns:
        The record's path; whether Excel answered: at once it is only slow, and later the hang
        cleared by itself (it waits at most 5 minutes, and stops asking at the first answer); and
        whether Excel was ended, which only a caller allowing it (``end_allowed``) and ``may_end``
        together permit, at once, without the waits.
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
    if not answered and end_allowed:
        windows = _window_list()
        names = [Path(a).name for a in arguments if a.endswith(".xlsx") and CONTAINER_MARK in a]
        allowed = may_end(_began_empty, windows, names)
        _log(record, f"may end Excel: {allowed} (run began empty: {_began_empty}; windows: "
                     f"{windows}; the request's workbooks: {names})")  # fmt: skip
        if allowed:
            print("Excel: ending it, as nothing else is open in it", file=sys.stderr, flush=True)
            return record, False, end_excel(record)
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
    return record, answered, False


@dataclass
class ExcelTimeoutError(Exception):
    """A request to Excel passed its limit."""

    what: str
    answered: bool
    record: Path
    ended: bool = False

    def __str__(self) -> str:
        state = (
            "Excel was ended: nothing else was open in it."
            if self.ended
            else "Excel answers now."
            if self.answered
            else "Excel does not answer, and is left as it is: run `python -m capture.excel_watch` "
            "to ask again before anyone quits it."
        )
        return f"Excel: {self.what}. {state} Evidence: {self.record}"


def osascript(
    *arguments: str, seconds: float, end_allowed: bool = False
) -> subprocess.CompletedProcess[str]:
    """Run ``osascript`` with a time limit.

    ``end_allowed``: past the limit Excel may be ended when ``may_end`` holds, for a probe whose
    case can keep Excel calculating without end (a function at an extreme input).

    Raises:
        ExcelTimeoutError: past the limit, after the evidence is saved.
    """
    preflight()
    keep_awake()
    try:
        result = subprocess.run(
            ["osascript", *arguments], capture_output=True, text=True, timeout=seconds
        )
    except subprocess.TimeoutExpired:
        what = f"{Path(arguments[0]).name} past {seconds:.0f} s"
        record, answered, ended = collect(what, arguments, seconds, end_allowed)
        raise ExcelTimeoutError(what, answered, record, ended) from None
    if result.returncode and APPLE_EVENT_TIMEOUT in result.stderr:
        what = f"{Path(arguments[0]).name}: AppleEvent timed out (-1712)"
        record, answered, ended = collect(what, arguments, seconds, end_allowed)
        raise ExcelTimeoutError(what, answered, record, ended)
    return result


def osascript_or_exit(*arguments: str, seconds: float) -> subprocess.CompletedProcess[str]:
    """The same for a script: past the limit it ends with the message."""
    try:
        return osascript(*arguments, seconds=seconds)
    except ExcelTimeoutError as timeout:
        sys.exit(str(timeout))


def main(argv: list[str]) -> None:
    """Ask Excel again, and add the answer to the record named, else to the latest if recent.

    A record older than ``RECENT_SECONDS`` belongs to an earlier hang: the answer then starts a
    record of its own rather than being added to it.
    """
    answered = answers()
    now = dt.datetime.now().strftime("%Y%m%d-%H%M%S")
    if argv:
        record = Path(argv[0])
    else:
        records = sorted(p for p in EVIDENCE.glob("*.txt") if not p.name.endswith(".sample.txt"))
        recent = records and time.time() - records[-1].stat().st_mtime <= RECENT_SECONDS
        record = records[-1] if recent else EVIDENCE / f"{now}-asked.txt"
    EVIDENCE.mkdir(parents=True, exist_ok=True)
    _log(record, f"asked again at {now}: answers {answered}\n{_state()}")
    print(f"answers about workbooks: {answered}; {record}")


if __name__ == "__main__":
    main(sys.argv[1:])
