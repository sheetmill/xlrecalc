"""``SUMIF``, ``SUMIFS`` and ``COUNTIFS`` visit only non-blank cells (ARCHITECTURE.md §7.7).

Their results are compared with a walk over every position, and their reads are counted.
"""

from __future__ import annotations

import random
from collections.abc import Iterator
from types import SimpleNamespace

import pytest

from tests.watchdog import bounded
from xlrecalc._compare import criterion
from xlrecalc._engine import Engine
from xlrecalc._functions._math import COUNTIFS, SUMIF, SUMIFS, _total
from xlrecalc._quirks import Quirks
from xlrecalc._store import MAX_COL, MAX_ROW
from xlrecalc._values import BLANK, ErrorValue, Ref

CTX = SimpleNamespace(quirks=Quirks.EXCEL)
CRITERIA = ["", "=", "<>", "<>1", "<>abc", ">0", "<=2", "1", "abc", "a*", "#N/A", 1.0, True]
VALUES = [1.0, 2.0, -1.5, 0.0, "1", "abc", "", True, False, ErrorValue.NA, ErrorValue.DIV0]


class Grid:
    """A source for ``Ref``: a sparse grid that counts the cells it is asked for."""

    def __init__(self, cells: dict[tuple[int, int], object]) -> None:
        self.cells = cells
        self.reads = 0

    def get(self, sheet: int, row: int, col: int) -> object:
        assert 0 <= row <= MAX_ROW, "off the sheet"
        assert 0 <= col <= MAX_COL, "off the sheet"
        self.reads += 1
        return self.cells.get((row, col), BLANK)

    def box(
        self, sheet: int, r1: int, c1: int, r2: int, c2: int
    ) -> Iterator[tuple[int, int, object]]:
        for (r, c), v in sorted(self.cells.items()):
            if r1 <= r <= r2 and c1 <= c <= c2:
                self.reads += 1
                yield r, c, v

    walk = box


def _grid(rng: random.Random, rows: int, cols: int, density: float) -> Grid:
    cells = {}
    for r in range(rows):
        for c in range(cols):
            if rng.random() < density:
                cells[r, c] = rng.choice(VALUES)
    return Grid(cells)


def _matching(conditions: list[tuple[Ref, object]]) -> list[tuple[int, int]]:
    """Every position of the ranges, each tested in every range: what the walks must equal."""
    first = conditions[0][0]
    tests = [(ref, criterion(c, Quirks.EXCEL)) for ref, c in conditions]
    return [
        (r, c)
        for r in range(first.rows)
        for c in range(first.columns)
        if all(test(ref.value(r, c)) for ref, test in tests)
    ]


def _sum_at(target: Ref, positions: list[tuple[int, int]]) -> object:
    values = (target.value(r, c) for r, c in positions)
    return _total((v for v in values if isinstance(v, float | ErrorValue)), CTX)


@pytest.mark.parametrize("density", [0.1, 0.5, 0.95])
@pytest.mark.parametrize("seed", range(40))
def test_the_walks_match_every_position(seed, density):
    rng = random.Random(seed)
    rows, cols = rng.randint(1, 6), rng.randint(1, 4)
    grid = _grid(rng, rows + 4, cols * 3 + 1, density)
    refs = [
        Ref(grid, 0, 2 * k % 4, k * cols, 2 * k % 4 + rows - 1, k * cols + cols - 1)
        for k in range(3)
    ]
    n = rng.randint(1, 3)
    conditions = [(refs[k], rng.choice(CRITERIA)) for k in range(n)]
    pairs = [x for pair in conditions for x in pair]
    positions = _matching(conditions)

    assert COUNTIFS(CTX, *pairs) == float(len(positions))
    assert SUMIFS(CTX, refs[-1], *pairs) == _sum_at(refs[-1], positions)
    ref, crit = conditions[0]
    alone = _matching([(ref, crit)])
    assert SUMIF(CTX, ref, crit) == _sum_at(ref, alone)
    assert SUMIF(CTX, ref, crit, refs[-1]) == _sum_at(refs[-1], alone)


WHOLE_SHEET = {
    "countifs blank": (lambda r: COUNTIFS(CTX, r, ""), 17179869182.0),
    "countifs number": (lambda r: COUNTIFS(CTX, r, ">0"), 2.0),
    "countifs two": (lambda r: COUNTIFS(CTX, r, "<>1", r, "<>2"), 17179869182.0),
    "sumifs blank": (lambda r: SUMIFS(CTX, r, r, "<>5"), 3.0),
    "sumifs number": (lambda r: SUMIFS(CTX, r, r, ">0"), 3.0),
    "sumif blank": (lambda r: SUMIF(CTX, r, "<>5"), 3.0),
    "sumif number": (lambda r: SUMIF(CTX, r, ">0"), 3.0),
}


@bounded
def test_reads_follow_the_cells_held_not_the_range():
    for case, (call, expected) in WHOLE_SHEET.items():
        grid = Grid({(0, 0): 1.0, (1, 1): 2.0})
        assert call(Ref(grid, 0, 0, 0, MAX_ROW, MAX_COL)) == expected, case
        assert grid.reads <= 16, case


def test_a_sum_range_is_cut_off_at_the_edge_of_the_sheet():
    grid = Grid({(0, 0): 1.0, (0, 1): 1.0, (1, 0): 5.0, (2, 0): 1.0})
    ref = Ref(grid, 0, 0, 0, 2, 1)
    edge = Ref(grid, 0, 0, MAX_COL, 0, MAX_COL)
    assert SUMIF(CTX, ref, ">0", edge) == 0.0
    assert SUMIF(CTX, ref, "<>z", edge) == 0.0


@pytest.mark.parametrize(
    "formula",
    [
        '=SUMIFS(OFFSET(C1,0,0,3),OFFSET(A1,0,0,3),">0")',
        '=COUNTIFS(OFFSET(A1,0,0,3),"<>1")',
        '=SUMIF(OFFSET(A1,0,0,3),"<>1",C1:C3)',
    ],
)
def test_a_range_read_only_in_part_is_still_checked_for_a_cycle(formula):
    report = Engine.from_cells({"S": {"A2": formula, "C1": 1.0}}).recalculate()
    assert {(r.row, r.col) for r in report.circular} == {(2, 1)}
