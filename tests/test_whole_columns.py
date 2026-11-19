"""Whole columns and rows (captures ``whole_columns`` and ``dyn_whole_columns``).

Most rows of these captures make arrays of a million cells each, and the two suites take about a
minute and a half to compute, so ``tests/test_captures.py`` and ``tests/test_dynamic_arrays.py``
do not include them. The rows that give ``MATCH`` one value as its lookup array are cheap and are
checked here.
"""

from __future__ import annotations

from tests.captures import load
from tests.test_captures import _input, formula_text, same
from xlrecalc._engine import Engine
from xlrecalc._formula._template import cell_axes
from xlrecalc._quirks import Quirks


def test_match_converts_one_value_given_as_its_lookup_array_to_a_number() -> None:
    """Excel's ``MATCH`` finds nothing in one value (``#N/A``) once it converts it to a number,
    and gives ``#VALUE!`` for a text that is not one (``MATCH("a","a",0)``, ``MATCH("",Z:Z&"",0)``
    reduced to its own row's blank)."""
    suite = load("whole_columns")
    engine = Engine(Quirks.EXCEL, today=suite.captured)
    s = engine.add_sheet("cases")
    for ref, value in suite.fixed.items():
        r, c = cell_axes(ref, 0, 0)
        engine.sheets[s].set(r.value, c.value, _input(value))  # type: ignore[arg-type]
    rows = {k: row for k, row in enumerate(suite.rows, 1) if row.tag == "match_scalar"}
    rows[18] = suite.rows[17]  # MATCH("",Z:Z&"",0), the row that showed the rule
    assert rows[18].formula == 'MATCH("",Z:Z&"",0)'
    for k, row in rows.items():  # each on its own row, as captured: some read their row of Z
        engine.set_formula(s, k - 1, 0, formula_text(row, k))
    engine.recalculate(full=True)
    wrong = [
        (row.formula, row.expected, engine.get(s, k - 1, 0))
        for k, row in rows.items()
        if not same(engine.get(s, k - 1, 0), row.expected)
    ]
    assert len(rows) == 27
    assert wrong == []
