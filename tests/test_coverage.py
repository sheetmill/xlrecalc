"""The committed coverage matrix is current (``tools/coverage.py``, ARCHITECTURE.md §8.4)."""

from __future__ import annotations

import pytest
from tools.coverage import OUTPUT, matrix


@pytest.mark.excel
def test_coverage_matrix_is_current():
    """The page as generated now, on any platform: it counts no split that the math library sets."""
    assert OUTPUT.read_text(encoding="utf-8") == matrix(), (
        "docs/coverage.md is out of date: run `uv run python tools/coverage.py --write`"
    )


def test_every_registered_function_appears():
    from xlrecalc._functions import SPECS

    text = OUTPUT.read_text(encoding="utf-8")
    for name in SPECS:
        bare = name.removeprefix("_XLFN.").removeprefix("_XLWS.")  # _xlfn._xlws.SORT too
        assert f"| `{bare}` |" in text, name
