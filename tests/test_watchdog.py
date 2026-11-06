"""The watchdog itself: a run past its limit is stopped, quickly, and said to be."""

import sys
import time

import pytest

from tests.watchdog import GIB, bounded, in_child, run


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
