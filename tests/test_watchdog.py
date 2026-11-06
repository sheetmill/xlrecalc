"""The watchdog itself: a run past its limit is stopped, quickly, and said to be."""

import os
import signal
import subprocess
import sys
import time
from pathlib import Path

import pytest

from tests.watchdog import GIB, bounded, in_child, run

ROOT = Path(__file__).resolve().parents[1]


def allocate(gib):
    data = bytearray(int(gib * GIB))
    data[::4096] = b"x" * len(data[::4096])  # touched, so it is really resident
    time.sleep(60)  # held: a process that frees it again can finish between two samples
    return len(data)


def double(value):
    return 2 * value


def fails():
    raise ValueError("from the child")


def test_a_command_past_the_limit_is_killed():
    touch = "b = bytearray(1 << 30); b[::4096] = b'x' * len(b[::4096])"
    code = f"import time; {touch}; time.sleep(60); print('done')"
    result = run([sys.executable, "-c", code], limit=GIB // 8)
    assert result.killed.startswith("used")
    assert "done" not in result.output
    assert result.peak < GIB


def test_a_command_past_its_time_is_killed():
    result = run([sys.executable, "-c", "import time; time.sleep(60)"], seconds=0.2)
    assert result.killed == "still running after 0 s"


def test_a_child_that_stays_small_is_left_alone():
    assert in_child(double, 21) == 42


def test_a_bomb_in_a_child_fails_the_test():
    with pytest.raises(AssertionError, match="past the limit"):
        in_child(allocate, 1, limit=GIB // 8)


def test_a_failure_in_a_child_is_reported():
    with pytest.raises(AssertionError, match="from the child"):
        in_child(fails)


@bounded
def test_a_bounded_test_runs_in_another_process():
    assert "tests.watchdog" in sys.argv or "--call" in sys.argv


HOLD = "import os, sys, time; open(sys.argv[1], 'w').write(str(os.getpid())); time.sleep(60)"


def _alive(pid):
    try:
        os.kill(pid, 0)
    except ProcessLookupError:
        return False
    return True


@pytest.mark.parametrize("sig", [signal.SIGTERM, signal.SIGKILL], ids=["SIGTERM", "SIGKILL"])
def test_what_it_runs_dies_with_the_watchdog(tmp_path, sig):
    # A time limit around the watchdog (gtimeout) must not leave the command running unwatched.
    pid_file = tmp_path / "pid"
    command = [sys.executable, "-m", "tests.watchdog", "--", sys.executable, "-c", HOLD]
    watchdog = subprocess.Popen([*command, str(pid_file)], cwd=ROOT)
    deadline = time.monotonic() + 20
    while not (pid_file.exists() and pid_file.read_text()):
        assert time.monotonic() < deadline, "the command did not start"
        time.sleep(0.05)
    child = int(pid_file.read_text())
    os.kill(watchdog.pid, sig)
    watchdog.wait()
    deadline = time.monotonic() + 10
    while _alive(child):
        if time.monotonic() > deadline:
            os.kill(child, signal.SIGKILL)
            pytest.fail("the command outlived the watchdog")
        time.sleep(0.05)


def test_a_finished_run_leaves_nothing_behind():
    result = run([sys.executable, "-c", "print('done')"])
    assert result.returncode == 0
    assert "done" in result.output
