"""The limits on Excel steps, without Excel: ``osascript`` is replaced."""

from __future__ import annotations

import subprocess

import pytest

from tests import excel_steps


def _times_out(*args, **kwargs):
    raise subprocess.TimeoutExpired(args[0], kwargs.get("timeout", 0))


def test_a_step_that_hangs_on_a_silent_excel_ends_the_run(monkeypatch):
    monkeypatch.setattr(subprocess, "run", _times_out)
    monkeypatch.setenv("PYTEST_CURRENT_TEST", "tests/test_excel_live.py::test_x (call)")
    with pytest.raises(pytest.exit.Exception) as stopped:
        excel_steps.step("capture/save_copy.applescript", "a", "b", seconds=1.0)
    message = str(stopped.value)
    assert "tests/test_excel_live.py::test_x" in message
    assert "save_copy.applescript past 1 s" in message
    assert stopped.value.returncode == 3


def test_a_step_that_is_slow_on_an_excel_that_answers_only_fails(monkeypatch):
    calls = []

    def run(command, **kwargs):
        calls.append(command)
        if len(calls) == 1:
            _times_out(command, **kwargs)
        return subprocess.CompletedProcess(command, 0, "0", "")

    monkeypatch.setattr(subprocess, "run", run)
    with pytest.raises(AssertionError, match="took too long"):
        excel_steps.step("capture/save_copy.applescript", seconds=1.0)


def test_a_step_passes_on_what_osascript_returned(monkeypatch):
    done = subprocess.CompletedProcess(["osascript"], 1, "", "error -50")
    monkeypatch.setattr(subprocess, "run", lambda *a, **k: done)
    assert excel_steps.step("capture/save_copy.applescript") is done
