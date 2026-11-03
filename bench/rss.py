"""Memory of this process, the same way for every runner in ``bench/``.

On macOS, the physical footprint (what Activity Monitor shows: resident memory plus memory the
system has compressed, from ``proc_pid_rusage``), now and the largest it has been. Resident
memory alone drops whenever macOS compresses pages, which makes it erratic for a process holding
a gigabyte. Elsewhere, the resident set from ``ps`` and ``ru_maxrss``. Every runner records it
at the same points: at the start (interpreter and imports), after reading the input (``table``:
the pickled table, for the runners that start from it), after loading, and at the end, so rows
can be compared with the same baseline subtracted or not.
"""

from __future__ import annotations

import ctypes
import os
import resource
import subprocess
import sys

_V4 = 4  # RUSAGE_INFO_V4: after the 16-byte uuid, 64-bit fields
_FOOTPRINT, _LIFETIME_MAX = 7, 28  # ri_phys_footprint, ri_lifetime_max_phys_footprint


def _rusage() -> list[int]:
    lib = ctypes.CDLL("/usr/lib/libproc.dylib")
    buf = (ctypes.c_uint64 * 64)()
    if lib.proc_pid_rusage(os.getpid(), _V4, buf) != 0:
        raise OSError("proc_pid_rusage failed")
    return list(buf)[2:]


def current_mb() -> float:
    """Memory now, in MB."""
    if sys.platform == "darwin":
        return _rusage()[_FOOTPRINT] / 1e6
    out = subprocess.run(
        ["ps", "-o", "rss=", "-p", str(os.getpid())], capture_output=True, text=True, check=True
    )
    return int(out.stdout.strip()) / 1e3  # kilobytes


def peak_mb() -> float:
    """The most memory so far, in MB."""
    if sys.platform == "darwin":
        return _rusage()[_LIFETIME_MAX] / 1e6
    return resource.getrusage(resource.RUSAGE_SELF).ru_maxrss / 1e3
