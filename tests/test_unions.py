"""Unions of references (``(A1:A3,C1:C3)``): every area as written (capture ``unions``).

The capture is the acceptance (``tests/test_captures.py``); here, a union against the same
function over its areas one by one, overlapping areas, the forms Excel gives an error, and the
dependencies a union gives.
"""

from __future__ import annotations

import random

import pytest

from xlrecalc._engine import Engine, node
from xlrecalc._values import ErrorValue


def _random_sheet(rng: random.Random) -> dict[str, object]:
    cells: dict[str, object] = {}
    for r in range(1, 7):
        for c in "ABCDE":
            roll = rng.random()
            if roll < 0.55:
                cells[f"{c}{r}"] = float(rng.randint(-50, 50))
            elif roll < 0.7:
                cells[f"{c}{r}"] = rng.choice(["x", "", "7"])
            elif roll < 0.8:
                cells[f"{c}{r}"] = rng.random() < 0.5
    return cells


def _random_area(rng: random.Random) -> str:
    c1, c2 = sorted(rng.choices("ABCDE", k=2))
    r1, r2 = sorted(rng.choices(range(1, 7), k=2))
    return f"{c1}{r1}" if (c1, r1) == (c2, r2) else f"{c1}{r1}:{c2}{r2}"


@pytest.mark.parametrize("seed", range(30))
def test_a_union_reads_each_area_as_written(seed: int) -> None:
    """SUM, COUNT, COUNTA, MAX and SUMSQ of a union equal the same over its areas as arguments of
    their own, overlapping areas counted once for each; LARGE reads them as one list, INDEX the
    area its area number names, AREAS counts them, nested unions flattened."""
    rng = random.Random(seed)
    data = _random_sheet(rng)
    areas = [_random_area(rng) for _ in range(rng.randint(2, 4))]
    union = f"({','.join(areas)})"
    nested = f"(({','.join(areas[:2])}),{','.join(areas[2:]) or areas[0]})"
    flat = areas + ([] if areas[2:] else [areas[0]])
    k = rng.randrange(len(areas))
    pairs = [
        (f"SUM({union})", "+".join(f"SUM({a})" for a in areas)),
        (f"COUNT({union})", "+".join(f"COUNT({a})" for a in areas)),
        (f"COUNTA({union})", "+".join(f"COUNTA({a})" for a in areas)),
        (f"MAX({union})", f"MAX({','.join(areas)})"),
        (f"SUMSQ({union})", f"SUMSQ({','.join(areas)})"),
        (f"SUM({nested})", "+".join(f"SUM({a})" for a in flat)),
        (f"AREAS({union})", f"{len(areas)}"),
        (f"AREAS({nested})", f"{len(flat)}"),
        (f"SUM(INDEX({union},0,0,{k + 1}))", f"SUM({areas[k]})"),
        (
            f"IF(COUNT({union})>0,LARGE({union},1),0)",
            f"IF(COUNT({','.join(areas)})>0,MAX({','.join(areas)}),0)",
        ),
    ]
    engine = Engine.from_cells({"cases": data})
    for i, (left, right) in enumerate(pairs):  # columns K and L, beside the data
        engine.set_formula(0, i, 10, left)
        engine.set_formula(0, i, 11, right)
    engine.recalculate()
    for i, (left, _) in enumerate(pairs):
        assert engine.get(0, i, 10) == engine.get(0, i, 11), (left, seed)
    assert not engine.report.unsupported, dict(engine.report.unsupported)


def test_the_forms_excel_gives_an_error() -> None:
    """A union as a value, in an operator or in array arithmetic is #VALUE!; areas on two sheets
    #VALUE!, a 3-D reference among them #NAME?; INDEX's area number below 1 #VALUE!, past the
    last area #REF!; a function no capture shows reading a union is reported, not guessed."""
    engine = Engine.from_cells(
        {"cases": {"A1": "=(B1,C1)", "A2": "=(B1,C1)+1", "A3": "=SUM((B1:B2,C1)*1)",
                   "A4": "=SUM((B1,Jan!A1))", "A5": "=SUM((Jan:Feb!A1,B1))",
                   "A6": "=INDEX((B1,C1),1,1,0)", "A7": "=INDEX((B1,C1),1,1,3)",
                   "A8": "=ABS((B1,C1))", "B1": 1.0, "B2": 2.0, "C1": 10.0},
         "Jan": {"A1": 1.0}, "Feb": {"A1": 2.0}}
    )  # fmt: skip
    engine.recalculate()
    v, ref = ErrorValue.VALUE, ErrorValue.REF
    assert [engine.get(0, r, 0) for r in range(7)] == [v, v, v, v, ErrorValue.NAME, v, ref]
    assert dict(engine.report.unsupported) == {"a union of ranges in ABS": 1}


def test_an_edit_inside_an_area_wakes_the_formula_and_one_outside_does_not() -> None:
    """The dependencies of a union are its areas: an edit in any of them makes the formula
    stale, an edit between them does not; a formula inside an area is computed first."""
    engine = Engine.from_cells(
        {"cases": {"A1": "=SUM((B1:B3,D1:D3,F2))", "A2": "=A1*2", "B1": 1.0, "D1": 10.0,
                   "F2": "=H1*3", "H1": 1.0}}
    )  # fmt: skip
    engine.recalculate()
    assert engine.get(0, 1, 0) == 2 * (1.0 + 10.0 + 3.0)
    formula, dependent = node(0, 0, 0), node(0, 1, 0)
    for row, col, wakes in ((2, 1, True), (1, 3, True), (1, 2, False), (4, 1, False)):
        engine.stale.clear()
        engine.set_value(0, row, col, 5.0)
        assert (formula in engine.stale) is wakes, (row, col)
        assert (dependent in engine.stale) is wakes, (row, col)
        engine.recalculate()
    assert engine.get(0, 0, 0) == 1.0 + 5.0 + 5.0 + 10.0 + 3.0
    engine.set_value(0, 0, 7, 2.0)  # H1, read by F2 inside the union
    engine.recalculate(targets=[node(0, 0, 0)])
    assert engine.get(0, 0, 0) == 1.0 + 5.0 + 5.0 + 10.0 + 6.0
