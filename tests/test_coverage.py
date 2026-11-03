"""The committed coverage matrix is current (``tools/coverage.py``, ARCHITECTURE.md §8.4)."""

from __future__ import annotations

import sys

import pytest
from tools.coverage import OUTPUT, matrix


@pytest.mark.excel
@pytest.mark.skipif(
    sys.platform != "darwin",
    reason="the matrix is generated on macOS, the reference platform: the split between exact "
    "and within-tolerance rows depends on the math library's last bits",
)
def test_coverage_matrix_is_current():
    assert OUTPUT.read_text(encoding="utf-8") == matrix(), (
        "docs/coverage.md is out of date: run `uv run python tools/coverage.py --write`"
    )


def test_every_registered_function_appears():
    from xlrecalc._functions import SPECS

    text = OUTPUT.read_text(encoding="utf-8")
    for name in SPECS:
        bare = name.removeprefix("_XLFN.")
        assert f"| `{bare}` |" in text, name
