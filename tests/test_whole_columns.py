"""Whole columns and rows (captures ``whole_columns`` and ``dyn_whole_columns``).

``tests/test_captures.py`` and ``tests/test_dynamic_arrays.py`` do not include these suites: a few
of their cases make arrays of a million cells however they are computed, and those suites run
each formula more than once. Every row is compared here instead, on the row it was captured on
(several read their own row of a column through implicit intersection):

- ``whole_columns`` in the default run, but for four rows over column AB, whose only value is in
  its last row, so the array they make is the whole column (``slow``);
- ``dyn_whole_columns`` in the slow tier: ``SORT``, ``UNIQUE``, ``FILTER`` and ``TAKE`` of whole
  columns make or keep arrays of a million rows.
"""

from __future__ import annotations

import pytest

from tests.captures import Row, load
from tests.test_captures import _input, formula_text, same
from xlrecalc._engine import Engine
from xlrecalc._formula._template import cell_axes
from xlrecalc._quirks import Quirks

WHOLE_COLUMN_AB = (
    'SUMPRODUCT(--(AB:AB<>""))',
    'LOOKUP(2,1/(AB:AB<>""),AB:AB)',
    'LOOKUP(2,1/(AB:AB<>""),ROW(AB:AB))',
    "SUMPRODUCT(AB:AB*1)",
)
"""The rows whose array is the whole of column AB: about 10 s together."""


def _wrong(rows: dict[int, Row]) -> list[tuple[str, object, object]]:
    """Each row computed on its own captured row beside the suite's data; those Excel differs on."""
    suite = load("whole_columns")
    engine = Engine(Quirks.EXCEL, today=suite.captured)
    s = engine.add_sheet("cases")
    for ref, value in suite.fixed.items():
        r, c = cell_axes(ref, 0, 0)
        engine.sheets[s].set(r.value, c.value, _input(value))  # type: ignore[arg-type]
    for k, row in rows.items():
        engine.set_formula(s, k - 1, 0, formula_text(row, k))
    engine.recalculate(full=True)
    return [
        (row.formula, row.expected, engine.get(s, k - 1, 0))
        for k, row in rows.items()
        if not same(engine.get(s, k - 1, 0), row.expected)
    ]


def _rows(column_ab: bool) -> dict[int, Row]:
    rows = dict(enumerate(load("whole_columns").rows, 1))
    return {k: r for k, r in rows.items() if (r.formula in WHOLE_COLUMN_AB) == column_ab}


def test_every_row_but_column_ab_matches_excel() -> None:
    rows = _rows(column_ab=False)
    assert len(rows) == 82
    assert _wrong(rows) == []


@pytest.mark.slow
def test_the_rows_over_column_ab_match_excel() -> None:
    rows = _rows(column_ab=True)
    assert len(rows) == len(WHOLE_COLUMN_AB)
    assert _wrong(rows) == []


def test_match_converts_one_value_given_as_its_lookup_array_to_a_number() -> None:
    """Excel's ``MATCH`` finds nothing in one value (``#N/A``) once it converts it to a number,
    and gives ``#VALUE!`` for a text that is not one (``MATCH("a","a",0)``, ``MATCH("",Z:Z&"",0)``
    reduced to its own row's blank). Rows of ``whole_columns``, checked apart from the others."""
    rows = {k: r for k, r in _rows(column_ab=False).items() if r.tag == "match_scalar"}
    rows[18] = load("whole_columns").rows[17]  # MATCH("",Z:Z&"",0), the row that showed the rule
    assert rows[18].formula == 'MATCH("",Z:Z&"",0)'
    assert len(rows) == 27
    assert _wrong(rows) == []


@pytest.mark.slow
def test_the_dynamic_capture_matches_excel() -> None:
    """Every cell of ``dyn_whole_columns`` and every spill range, as ``test_dynamic_arrays``
    checks its own suites."""
    from tests import test_dynamic_arrays as dyn

    cells = [r[1] for r in dyn._rows("whole_columns") if r[0] == "cell" and dyn._cell(r[1])[1] < 27]
    assert len(cells) == 31
    for address in cells:
        dyn.test_cell_matches_excel("whole_columns", address)
    dyn.test_spill_ranges_match_excel("whole_columns")
