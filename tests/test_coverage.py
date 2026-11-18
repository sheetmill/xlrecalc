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


EXCEL_DRAW = [-2.0, -3.0, -1.0, -3.0]
"""Excel's saved draw for ``RANDARRAY(1,4,-3,-1,TRUE)`` (capture ``dyn_randarray``, case 15)."""


def _randarray(monkeypatch: pytest.MonkeyPatch, forced: bool) -> tuple[int, int, int]:
    """``RANDARRAY``'s rows, matches and listed rows in the dynamic captures, with its draw for
    case 15 forced to Excel's when ``forced``, else drawn as usual."""
    import collections
    import random as real

    from tools.coverage import Evidence, _dynamic_evidence

    from tests import test_dynamic_arrays as dyn
    from xlrecalc._functions import _makers

    draws = iter(EXCEL_DRAW)

    class Generator:
        """``random`` with ``randint(-3, -1)`` giving Excel's saved draw, in order."""

        def __getattr__(self, name: str) -> object:
            return getattr(real, name)

        def randint(self, a: int, b: int) -> int:
            return int(next(draws)) if forced and (a, b) == (-3, -1) else real.randint(a, b)

    monkeypatch.setattr(_makers, "random", Generator())
    monkeypatch.setattr(dyn, "_ENGINES", {})  # engines built here stay here
    found: dict[str, Evidence] = collections.defaultdict(Evidence)
    _dynamic_evidence(found)
    if forced:
        engine = dyn.engine_for("randarray")
        assert [engine.get(0, 180, c) for c in range(1, 5)] == EXCEL_DRAW  # B181:E181
    e = found["RANDARRAY"]
    return e.rows, e.match, e.listed


def test_a_random_draw_that_matches_excel_changes_no_count(monkeypatch: pytest.MonkeyPatch) -> None:
    """The page counts a random case as listed whatever its draw, Excel's own included."""
    assert _randarray(monkeypatch, forced=True) == _randarray(monkeypatch, forced=False)
