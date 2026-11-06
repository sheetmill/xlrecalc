"""Run something that may try to use too much memory, and stop it before it does.

A test that feeds crafted input to the package must not be able to exhaust the machine if a guard
regresses. ``bounded`` runs a test in a child process and fails it when the child's memory passes a
limit. ``watch_session`` covers every other test, by stopping the whole run. Both sample the size
of the process group with ``ps``: a thread could not do this, since one huge allocation holds the
interpreter's lock, and ``ulimit -v`` is not a limit on macOS.

What the watchdog runs dies with it. The command runs in a session of its own, so that the whole
group can be killed; if the watchdog were killed in turn (a time limit around it, say), the command
would live on with no limit at all. A guardian process holds the read end of a pipe whose only
writer is the watchdog: when the watchdog dies, by any signal, SIGKILL included, the kernel closes
the pipe, and the guardian kills the command's group. A watchdog that ends normally first writes a
byte, and the guardian leaves the group alone.

    uv run python -m tests.watchdog -- uv run pytest   # anything else, under the same limit
"""

from __future__ import annotations

import contextlib
import functools
import importlib
import os
import pickle
import signal
import subprocess
import sys
import tempfile
import time
from collections.abc import Callable
from dataclasses import dataclass
from pathlib import Path
from typing import Any

GIB = 1 << 30
LIMIT = 4 * GIB
"""What anything may use. A run past it is killed, and that is a finding: find out why."""
TEST_LIMIT = 1 * GIB
"""What one crafted case may use: far more than a refusal needs, and quickly reached by a bomb."""
TEST_SECONDS = 120.0
_SAMPLE = 0.02
_ROOT = Path(__file__).resolve().parents[1]


def _group_size(group: int) -> int:
    """Bytes resident in a process group, or 0 when it is gone."""
    listing = subprocess.run(
        ["ps", "-A", "-o", "pgid=,rss="], capture_output=True, text=True, check=False
    ).stdout
    total = 0
    for line in listing.splitlines():
        fields = line.split()
        if len(fields) == 2 and int(fields[0]) == group:
            total += int(fields[1]) * 1024
    return total


@dataclass
class Result:
    """How a watched command ended; ``killed`` says why the watchdog stopped it, if it did."""

    returncode: int
    output: str
    peak: int
    killed: str | None


def run(
    command: list[str], *, limit: int = LIMIT, seconds: float = 3600.0, capture: bool = True
) -> Result:
    """Run a command in its own process group, and kill the group past ``limit`` or ``seconds``.

    The group is killed too if this process dies before the command ends (see the module's
    docstring).
    """
    start = time.monotonic()
    with tempfile.TemporaryFile("w+", encoding="utf-8", errors="replace") as log:
        process = subprocess.Popen(
            command,
            stdout=log if capture else None,
            stderr=subprocess.STDOUT if capture else None,
            start_new_session=True,
            cwd=_ROOT,
        )
        alive = _guard(process.pid)
        try:
            peak = 0
            killed = None
            while process.poll() is None:
                peak = max(peak, _group_size(process.pid))
                if peak > limit:
                    killed = f"used {peak / GIB:.1f} GiB, past the limit of {limit / GIB:.1f} GiB"
                elif time.monotonic() - start > seconds:
                    killed = f"still running after {seconds:.0f} s"
                if killed:
                    os.killpg(process.pid, signal.SIGKILL)
                    break
                time.sleep(_SAMPLE)
            process.wait()
        except BaseException:  # interrupted here: the group must not outlive the watchdog
            with contextlib.suppress(ProcessLookupError):
                os.killpg(process.pid, signal.SIGKILL)
            os.close(alive)
            raise
        os.write(alive, b"x")  # finished: the guardian leaves the group alone
        os.close(alive)
        log.seek(0)
        return Result(process.returncode, log.read(), peak, killed)


def _guard(group: int) -> int:
    """Start a guardian that kills ``group`` when this process dies; the pipe end to hold.

    The guardian has a session of its own, so a signal to this process or its group does not
    reach it, and it inherits only the read end: the pipe closes exactly when this process does.
    """
    read, write = os.pipe()
    subprocess.Popen(
        [sys.executable, str(Path(__file__)), "--guard", str(group), str(read)],
        pass_fds=(read,),
        start_new_session=True,
    )
    os.close(read)
    return write


def _guardian(group: int, fd: int) -> None:
    """Wait for the watchdog: a byte means it finished, end of file means it died."""
    try:
        finished = os.read(fd, 1) == b"x"
    except OSError:
        finished = False
    if not finished:
        with contextlib.suppress(ProcessLookupError, PermissionError):
            os.killpg(group, signal.SIGKILL)


def bounded(
    test: Callable[..., Any] | None = None,
    *,
    limit: int = TEST_LIMIT,
    seconds: float = TEST_SECONDS,
) -> Any:
    """Run a test in a child process, and fail it if the child passes ``limit`` or ``seconds``.

    Its fixtures must be values that can be pickled. A test that needs a result in this process
    calls :func:`in_child`.
    """

    def decorate(function: Callable[..., Any]) -> Callable[..., Any]:
        @functools.wraps(function)
        def wrapper(*args: Any, **kwargs: Any) -> None:
            in_child(function, *args, limit=limit, seconds=seconds, **kwargs)

        return wrapper

    return decorate(test) if test is not None else decorate


def in_child(
    function: Callable[..., Any],
    *args: Any,
    limit: int = TEST_LIMIT,
    seconds: float = TEST_SECONDS,
    **kwargs: Any,
) -> Any:
    """``function(*args, **kwargs)``, a module-level function, in a child under the watchdog."""
    with tempfile.TemporaryDirectory() as folder:
        call, answer = Path(folder) / "call", Path(folder) / "answer"
        call.write_bytes(pickle.dumps((args, kwargs)))
        child = [sys.executable, "-m", "tests.watchdog", "--call", function.__module__]
        child += [function.__qualname__, str(call), str(answer)]
        result = run(child, limit=limit, seconds=seconds)
        if result.killed:
            raise AssertionError(
                f"{function.__qualname__} {result.killed}\n{result.output[-2000:]}"
            )
        if result.returncode != 0:
            raise AssertionError(f"{function.__qualname__} failed:\n{result.output[-4000:]}")
        return pickle.loads(answer.read_bytes())


def watch_session(limit: int = LIMIT) -> None:
    """Stop this process, the whole test run, if its memory passes ``limit``.

    The watcher ends when this process does. A test's child under ``bounded`` runs in a session
    of its own, so killing this process does not reach it; its guardian ends it instead.
    """
    subprocess.Popen(
        [sys.executable, str(Path(__file__)), "--session", str(os.getpid()), str(limit)],
        start_new_session=True,
    )


def _watch_process(pid: int, limit: int) -> None:
    while True:
        listing = subprocess.run(
            ["ps", "-o", "rss=", "-p", str(pid)], capture_output=True, text=True, check=False
        ).stdout.strip()
        if not listing:
            return
        size = int(listing) * 1024
        if size > limit:
            sys.stderr.write(
                f"\ntests/watchdog.py: the test run used {size / GIB:.1f} GiB, past the limit of "
                f"{limit / GIB:.1f} GiB, and was killed. Find out why before running it again.\n"
            )
            os.kill(pid, signal.SIGKILL)
            return
        time.sleep(_SAMPLE * 5)


def _call(module: str, name: str, call: str, answer: str) -> None:
    function = importlib.import_module(module)
    for part in name.split("."):
        function = getattr(function, part)
    function = getattr(function, "__wrapped__", function)
    args, kwargs = pickle.loads(Path(call).read_bytes())
    Path(answer).write_bytes(pickle.dumps(function(*args, **kwargs)))


def main(arguments: list[str]) -> int:
    """The command line: ``-- COMMAND`` runs a command under :data:`LIMIT`."""
    if arguments[:1] == ["--call"]:
        _call(*arguments[1:5])
        return 0
    if arguments[:1] == ["--guard"]:
        _guardian(int(arguments[1]), int(arguments[2]))
        return 0
    if arguments[:1] == ["--session"]:
        _watch_process(int(arguments[1]), int(arguments[2]))
        return 0
    if arguments[:1] == ["--"]:
        result = run(arguments[1:], capture=False)
        if result.killed:
            sys.stderr.write(f"tests/watchdog.py: {result.killed}\n")
        sys.stderr.write(f"tests/watchdog.py: peak {result.peak / GIB:.2f} GiB\n")
        return result.returncode
    sys.stderr.write(__doc__ or "")
    return 2


if __name__ == "__main__":
    sys.exit(main(sys.argv[1:]))
