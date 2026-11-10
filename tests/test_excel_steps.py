"""The limits on Excel steps and the evidence they save, without Excel: commands are replaced."""

from __future__ import annotations

import os
import subprocess
import sys
from pathlib import Path

import pytest
from capture import excel_watch

from tests import excel_steps


def _times_out(*args, **kwargs):
    raise subprocess.TimeoutExpired(args[0], kwargs.get("timeout", 0))


@pytest.fixture(autouse=True)
def waits(monkeypatch, tmp_path):
    """The evidence goes to a temporary directory, and its waits are only noted."""
    slept: list[float] = []
    monkeypatch.setattr(excel_watch, "EVIDENCE", tmp_path)
    monkeypatch.setattr(excel_watch.time, "sleep", slept.append)
    monkeypatch.setattr(excel_watch.shutil, "which", lambda name: None)  # no real caffeinate
    monkeypatch.setattr(excel_watch, "_awake", None)
    return slept


def test_a_step_that_hangs_on_a_silent_excel_ends_the_run(monkeypatch, waits):
    monkeypatch.setattr(subprocess, "run", _times_out)
    monkeypatch.setenv("PYTEST_CURRENT_TEST", "tests/test_excel_live.py::test_x (call)")
    with pytest.raises(pytest.exit.Exception) as stopped:
        excel_steps.step("capture/save_copy.applescript", "a", "b", seconds=1.0)
    message = str(stopped.value)
    assert "tests/test_excel_live.py::test_x" in message
    assert "save_copy.applescript past 1 s" in message
    assert stopped.value.returncode == 3
    assert waits == [60.0, 240.0]


def test_a_step_that_is_slow_on_an_excel_that_answers_only_fails(monkeypatch, waits):
    calls = []

    def run(command, **kwargs):
        calls.append(command)
        if len(calls) == 1:
            _times_out(command, **kwargs)
        return subprocess.CompletedProcess(command, 0, "0", "")

    monkeypatch.setattr(subprocess, "run", run)
    with pytest.raises(AssertionError, match="took too long"):
        excel_steps.step("capture/save_copy.applescript", seconds=1.0)
    assert waits == []


def test_a_step_passes_on_what_osascript_returned(monkeypatch):
    done = subprocess.CompletedProcess(["osascript"], 1, "", "error -50")
    monkeypatch.setattr(subprocess, "run", lambda *a, **k: done)
    assert excel_steps.step("capture/save_copy.applescript") is done


def test_the_record_holds_the_request_the_state_and_each_answer(monkeypatch, tmp_path, waits):
    """A hang that clears at the 1 minute question stops the asking, and the record says so."""
    questions = []

    def run(command, **kwargs):
        if command[0] == "osascript" and command[-1] == "count":
            questions.append(command)
            if len(questions) == 1:
                _times_out(command, **kwargs)
            return subprocess.CompletedProcess(command, 0, "0", "")
        if command[0] == "pgrep":
            return subprocess.CompletedProcess(command, 0, "4242\n", "")
        if command[0] == "ps":
            return subprocess.CompletedProcess(command, 0, "4242 06:00:00 0.0 900000", "")
        if command[0] == "sample":
            (tmp_path / "written.sample.txt").write_text("x")
        if command[0] == "pmset":
            return subprocess.CompletedProcess(command, 0, "t Display is turned off\nz", "")
        return subprocess.CompletedProcess(command, 0, "", "")

    monkeypatch.setattr(subprocess, "run", run)
    container = "/Users/u/Library/Containers/com.microsoft.Excel/Data/Documents/xlrecalc/"
    request = ("s.applescript", "book.xlsx", container + "probe.xlsx", container + "copy.xlsx")
    record, answered, _ = excel_watch.collect("x past 9 s", request, 9.0)
    text = record.read_text()
    assert answered
    assert waits == [60.0]
    assert "display kept awake: no" in text
    assert "after the answer: closed without saving" in text
    for expected in (
        "x past 9 s",
        "limit: 9 s",
        "s.applescript book.xlsx",
        "answers about workbooks at once: False",
        "4242 06:00:00",
        "Display is turned off",
        "answers after 60 s: True",
    ):
        assert expected in text


def test_a_failing_command_does_not_lose_the_record(monkeypatch):
    def run(command, **kwargs):
        raise OSError("no such command")

    monkeypatch.setattr(subprocess, "run", run)
    monkeypatch.setattr(excel_watch, "answers", lambda seconds=0: False)
    record, answered, _ = excel_watch.collect("x", ("s",), 1.0)
    assert not answered
    assert "answers after 300 s: False" in record.read_text()


def test_the_module_imports_without_openpyxl():
    code = "import sys, capture.excel_watch; sys.exit('openpyxl' in sys.modules)"
    assert subprocess.run([sys.executable, "-c", code], cwd=excel_watch.ROOT).returncode == 0


def test_a_capture_limit_grows_with_its_cases_and_is_generous():
    pytest.importorskip("openpyxl")  # capture.pipeline builds workbooks with it
    from capture.pipeline import capture_limit

    assert capture_limit(10) >= 600
    assert capture_limit(100_000) > capture_limit(10) + 5_000


def test_a_recovered_hang_closes_only_the_requests_container_workbooks(monkeypatch):
    """Their names reach AppleScript as arguments, never as source, and alerts come back on."""
    commands = []
    monkeypatch.setattr(
        subprocess, "run", lambda c, **k: commands.append(c) or subprocess.CompletedProcess(c, 0)
    )
    container = "/Users/u/Library/Containers/com.microsoft.Excel/Data/Documents/xlrecalc/"
    excel_watch.tidy(("s.applescript", "/tmp/other.xlsx", container + 'a "b".xlsx'))
    (command,) = commands
    source = [command[i + 1] for i, part in enumerate(command) if part == "-e"]
    assert command[-1] == 'a "b".xlsx'
    assert "/tmp/other.xlsx" not in command
    assert "set display alerts to true" in source
    assert not any('"b"' in line for line in source)


def test_the_display_is_kept_awake_for_the_process(monkeypatch):
    started = []

    class Fake:
        pid = 7

        def __init__(self, command, **kwargs):
            started.append(command)

        def poll(self):
            return None

    monkeypatch.setattr(excel_watch.shutil, "which", lambda name: "/usr/bin/caffeinate")
    monkeypatch.setattr(excel_watch.subprocess, "Popen", Fake)
    ran = []

    def run(command, **kwargs):
        ran.append(command)
        text = "t Display is turned off\n" if command[0] == "pmset" else ""
        return subprocess.CompletedProcess(command, 0, text, "")

    monkeypatch.setattr(subprocess, "run", run)
    excel_watch.keep_awake()
    excel_watch.keep_awake()
    assert started == [["/usr/bin/caffeinate", "-d", "-w", str(excel_watch.os.getpid())]]
    assert ("/usr/bin/caffeinate", "-u", "-t", "1") in ran  # wakes a display already off
    assert excel_watch.awake() == (
        "display kept awake: yes (caffeinate -d, pid 7); display at the start of the run, "
        "last change: t Display is turned off"
    )


class _Ran:
    def __init__(self, returncode: int, stdout: str = "") -> None:
        self.returncode, self.stdout = returncode, stdout


def _preflight(monkeypatch, windows, documents):
    """Run the preflight with Excel's windows and its answer to ``count of windows`` given."""
    monkeypatch.setattr(excel_watch, "_checked", False)
    monkeypatch.setattr(excel_watch.shutil, "which", lambda tool: "/usr/bin/" + tool)
    monkeypatch.setattr(excel_watch, "excel_windows", lambda: windows)

    def run(command, **kwargs):
        if command[0] == "pgrep":
            return _Ran(0)
        if documents is None:
            raise excel_watch.subprocess.TimeoutExpired(command, 10.0)
        return _Ran(0, f"{documents}\n")

    monkeypatch.setattr(excel_watch.subprocess, "run", run)
    excel_watch.preflight()


def test_preflight_passes_with_only_workbook_windows(monkeypatch):
    _preflight(monkeypatch, ["1200 x 800"], 1)
    _preflight(monkeypatch, [], 0)


@pytest.mark.parametrize(("windows", "documents"), [(["520 x 300"], 0), (["520 x 300"], None)])
def test_preflight_stops_on_a_window_of_excels_own(monkeypatch, windows, documents):
    with pytest.raises(SystemExit, match="A window of Excel's own"):
        _preflight(monkeypatch, windows, documents)


_BOX = "/Users/u/Library/Containers/com.microsoft.Excel/Data/Documents/xlrecalc/"


def test_may_end_only_when_nothing_else_can_be_in_excel():
    names = ["probe.xlsx", "probe_saved.xlsx"]
    may_end = excel_watch.may_end
    assert may_end(True, [("800", "600", "probe")], names)
    assert may_end(True, [("800", "600", "")], names)  # no title from macOS: counted
    assert may_end(True, [], names)
    assert not may_end(False, [("800", "600", "probe")], names)  # a workbook was open at the start
    assert not may_end(None, [("800", "600", "probe")], names)  # the start is not known
    assert not may_end(True, None, names)  # the window server did not answer
    assert not may_end(True, [("800", "600", "Budget")], names)  # a window of someone else's
    assert not may_end(True, [("1", "1", ""), ("2", "2", ""), ("3", "3", "")], names)  # one more
    assert not may_end(True, [], [])  # a request naming no workbook of the probe's


def test_collect_ends_excel_only_when_allowed_and_may_end_holds(monkeypatch, tmp_path):
    monkeypatch.setattr(excel_watch, "EVIDENCE", tmp_path)
    monkeypatch.setattr(excel_watch, "answers", lambda seconds=0: False)
    monkeypatch.setattr(excel_watch, "_state", lambda: "state")
    monkeypatch.setattr(excel_watch, "_sample", lambda path: "sample")
    monkeypatch.setattr(excel_watch, "_window_list", lambda: [("800", "600", "probe")])
    monkeypatch.setattr(excel_watch.time, "sleep", lambda s: None)
    ended: list[Path] = []
    monkeypatch.setattr(excel_watch, "end_excel", lambda record: ended.append(record) or True)
    request = ("s.applescript", _BOX + "probe.xlsx", _BOX + "probe_saved.xlsx")

    monkeypatch.setattr(excel_watch, "_began_empty", True)
    record, answered, was_ended = excel_watch.collect("x", request, 1.0, end_allowed=True)
    assert (answered, was_ended, ended) == (False, True, [record])
    assert "may end Excel: True" in record.read_text()

    ended.clear()
    _, _, was_ended = excel_watch.collect("x", request, 1.0)  # the caller does not allow it
    assert (was_ended, ended) == (False, [])

    monkeypatch.setattr(excel_watch, "_began_empty", False)  # a workbook was open at the start
    record, _, was_ended = excel_watch.collect("x", request, 1.0, end_allowed=True)
    assert (was_ended, ended) == (False, [])
    assert "may end Excel: False" in record.read_text()


def test_an_apple_event_timeout_saves_a_record_and_stops(monkeypatch, tmp_path):
    """AppleScript's own limit (-1712) can come before ours: the same hang, recorded alike."""
    error = "save_copy.applescript:1288:1302: execution error: AppleEvent timed out. (-1712)"

    def run(command, **kwargs):
        if command[0] == "osascript" and command[1].endswith("save_copy.applescript"):
            return subprocess.CompletedProcess(command, 1, "", error)
        return subprocess.CompletedProcess(command, 0, "", "")

    monkeypatch.setattr(subprocess, "run", run)
    monkeypatch.setattr(excel_watch, "answers", lambda seconds=0: True)
    monkeypatch.setattr(excel_watch, "_checked", True)
    with pytest.raises(excel_watch.ExcelTimeoutError) as stopped:
        excel_watch.osascript("capture/save_copy.applescript", "a.xlsx", seconds=600.0)
    assert stopped.value.answered
    (record,) = [p for p in tmp_path.glob("*.txt") if not p.name.endswith(".sample.txt")]
    assert "save_copy.applescript: AppleEvent timed out (-1712)" in record.read_text()


def test_asking_again_adds_to_a_recent_record_only(monkeypatch, tmp_path):
    monkeypatch.setattr(excel_watch, "answers", lambda seconds=0: True)
    monkeypatch.setattr(excel_watch, "_state", lambda: "state")
    old = tmp_path / "20260101-000000.txt"
    old.write_text("an earlier hang\n")
    hour_ago = excel_watch.time.time() - 3600
    os.utime(old, (hour_ago, hour_ago))
    excel_watch.main([])
    assert old.read_text() == "an earlier hang\n"
    (asked,) = tmp_path.glob("*-asked.txt")
    assert "answers True" in asked.read_text()
    excel_watch.main([])
    assert asked.read_text().count("asked again") == 2  # now the latest, and recent
    excel_watch.main([str(old)])
    assert "asked again" in old.read_text()  # a record named takes the answer
