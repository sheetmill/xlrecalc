"""The limits on Excel steps and the evidence they save, without Excel: commands are replaced."""

from __future__ import annotations

import subprocess
import sys

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
    record, answered = excel_watch.collect("x past 9 s", request, 9.0)
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
    record, answered = excel_watch.collect("x", ("s",), 1.0)
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
    excel_watch.keep_awake()
    excel_watch.keep_awake()
    assert started == [["/usr/bin/caffeinate", "-d", "-w", str(excel_watch.os.getpid())]]
    assert excel_watch.awake() == "display kept awake: yes (caffeinate -d, pid 7)"
