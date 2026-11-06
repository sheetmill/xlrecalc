"""Lookups visit a range's non-blank cells: their results against a walk of every position."""

from __future__ import annotations

import random
from types import SimpleNamespace

import pytest

from tests.test_criteria_walk import Grid
from tests.watchdog import bounded
from xlrecalc._compare import compare, lookup_equal, wildcard
from xlrecalc._formula import MISSING
from xlrecalc._functions._lookup import HLOOKUP, LOOKUP, MATCH, VLOOKUP, XLOOKUP, XMATCH, _x_compare
from xlrecalc._quirks import Quirks
from xlrecalc._store import MAX_ROW
from xlrecalc._values import BLANK, ErrorValue, Ref

CTX = SimpleNamespace(quirks=Quirks.EXCEL)
NA = ErrorValue.NA
VALUES = [0.0, 1.0, 2.0, 2.5, 3.0, 7.0, "", "a", "ab", "B", "*", True, False, NA, ErrorValue.DIV0]
KEYS = [0.0, 1.0, 2.0, 2.2, 9.0, -1.0, "", "a", "b", "a*", "?b", "~*", True, False]


def _full_exact(key, line):
    if isinstance(key, str) and any(ch in key for ch in "*?~"):
        pattern = wildcard(key)
        return next(
            (i for i, v in enumerate(line) if isinstance(v, str) and pattern.fullmatch(v)), None
        )
    return next((i for i, v in enumerate(line) if lookup_equal(key, v)), None)


def _kind(v):
    if isinstance(v, bool):
        return 2
    if isinstance(v, str):
        return 1
    return 0 if isinstance(v, float) else -1


def _full_approximate(key, line, descending=False):
    candidates = [i for i, v in enumerate(line) if _kind(v) == _kind(key)]
    lo, hi, found = 0, len(candidates) - 1, None
    while lo <= hi:
        mid = (lo + hi) // 2
        v = line[candidates[mid]]
        if compare(v, key, ">=" if descending else "<=", Quirks.EXCEL) is True:
            found, lo = candidates[mid], mid + 1
        else:
            hi = mid - 1
    return found


def _position(found):
    return NA if found is None else float(found + 1)


def _value(cell):
    return cell  # a blank stays one: ISBLANK of a lookup sees it (capture criteria)


def _grid(rng, rows, cols, density):
    cells = {}
    for r in range(rows):
        for c in range(cols):
            if rng.random() < density:
                cells[r, c] = rng.choice(VALUES)
    return Grid(cells)


def _key(rng, grid, rows):
    """A key as a lookup is given it: a value, or a cell's value, which may be blank."""
    kind = rng.random()
    if kind < 0.8:
        return rng.choice(KEYS)
    return grid.cells.get((rng.randrange(rows), 0), BLANK)


@pytest.mark.parametrize("density", [0.1, 0.5, 0.95])
@pytest.mark.parametrize("seed", range(40))
def test_match_and_the_table_lookups_match_a_walk_of_every_position(seed, density):
    rng = random.Random(seed)
    rows, cols = rng.randint(1, 12), rng.randint(2, 6)
    grid = _grid(rng, rows, cols, density)
    table = Ref(grid, 0, 0, 0, rows - 1, cols - 1)
    column = Ref(grid, 0, 0, 0, rows - 1, 0)
    row = Ref(grid, 0, 0, 0, 0, cols - 1)
    for _ in range(12):
        key = _key(rng, grid, rows)
        if isinstance(key, ErrorValue):
            continue
        lookup = 0.0 if key is BLANK else key  # a blank key is 0, as a lookup reads it
        down = [grid.cells.get((r, 0), BLANK) for r in range(rows)]
        across = [grid.cells.get((0, c), BLANK) for c in range(cols)]

        assert MATCH(CTX, key, column, 0.0) == _position(_full_exact(lookup, down))
        assert MATCH(CTX, key, column, 1.0) == _position(_full_approximate(lookup, down))
        assert MATCH(CTX, key, column, -1.0) == _position(_full_approximate(lookup, down, True))
        assert MATCH(CTX, key, row, 0.0) == _position(_full_exact(lookup, across))
        assert MATCH(CTX, key, row, 1.0) == _position(_full_approximate(lookup, across))

        for approximate, search in ((False, _full_exact), (True, _full_approximate)):
            found = search(lookup, down)
            want = NA if found is None else _value(grid.cells.get((found, 1), BLANK))
            assert VLOOKUP(CTX, key, table, 2.0, approximate) == want
            found = search(lookup, across)
            want = NA if found is None else _value(grid.cells.get((1, found), BLANK))
            if rows > 1:
                assert HLOOKUP(CTX, key, table, 2.0, approximate) == want


@pytest.mark.parametrize("seed", range(20))
def test_lookup_matches_a_walk_of_every_position(seed):
    rng = random.Random(seed)
    rows = rng.randint(1, 12)
    grid = _grid(rng, rows, 2, 0.5)
    column = Ref(grid, 0, 0, 0, rows - 1, 0)
    down = [grid.cells.get((r, 0), BLANK) for r in range(rows)]
    for key in KEYS:
        found = _full_approximate(key, down)
        want = NA if found is None else _value(grid.cells.get((found, 1), BLANK))
        assert LOOKUP(CTX, key, column, Ref(grid, 0, 0, 1, rows - 1, 1)) == want


LOOKUPS = {
    "match exact": lambda k, r: MATCH(CTX, k, r, 0.0),
    "match sorted": lambda k, r: MATCH(CTX, k, r, 1.0),
    "lookup": lambda k, r: LOOKUP(CTX, k, r, r),
}


@bounded
def test_reads_follow_the_cells_held_not_the_column():
    for case, lookup in LOOKUPS.items():
        for key in (2.0, "", "a*", 0.0):
            grid = Grid({(0, 0): 1.0, (5, 0): 2.0, (900_000, 0): 3.0})
            lookup(key, Ref(grid, 0, 0, 0, MAX_ROW, 0))
            assert grid.reads <= 8, (case, key)


@bounded
def test_a_table_lookup_reads_what_it_holds():
    grid = Grid({(0, 0): 1.0, (5, 0): 2.0, (5, 1): 20.0, (900_000, 0): 3.0})
    table = Ref(grid, 0, 0, 0, MAX_ROW, 1)
    assert VLOOKUP(CTX, 2.0, table, 2.0, False) == 20.0
    assert VLOOKUP(CTX, 2.0, table, 2.0, True) == 20.0
    assert grid.reads <= 16


def _x_full(key, line, match, search):
    """``XMATCH`` by a walk of every position of the line, blanks included."""
    order = list(enumerate(line))
    if search < 0:
        order.reverse()
    if key is BLANK:
        gap = next((i for i, v in order if v is BLANK), None)
        if gap is not None:
            return gap
    if match == 2 and isinstance(key, str):
        pattern = wildcard(key)
        return next((i for i, v in order if isinstance(v, str) and pattern.fullmatch(v)), None)
    best, best_value = None, None
    for i, v in order:
        c = _x_compare(key, v, CTX)
        if c is None:
            continue
        if c == 0:
            return i
        if c == match and (best is None or _x_compare(best_value, v, CTX) == -match):
            best, best_value = i, v
    return best


@pytest.mark.parametrize("density", [0.1, 0.5, 0.95])
@pytest.mark.parametrize("seed", range(30))
def test_xmatch_and_xlookup_match_a_walk_of_every_position(seed, density):
    rng = random.Random(seed)
    rows = rng.randint(1, 12)
    grid = _grid(rng, rows, 2, density)
    column = Ref(grid, 0, 0, 0, rows - 1, 0)
    results = Ref(grid, 0, 0, 1, rows - 1, 1)
    line = [grid.cells.get((r, 0), BLANK) for r in range(rows)]
    for _ in range(12):
        key = _key(rng, grid, rows)
        if isinstance(key, ErrorValue):
            continue
        for match in (0.0, -1.0, 1.0, 2.0):
            for search in (1.0, -1.0):
                want = _x_full(key, line, int(match), int(search))
                assert XMATCH(CTX, key, column, match, search) == _position(want)
                got = XLOOKUP(CTX, key, column, results, MISSING, match, search)
                if want is None:
                    assert got == NA
                else:
                    assert (got.r1, got.c1) == (want, 1)


@bounded
def test_xlookup_reads_what_a_whole_column_holds():
    for key in (2.0, "a*", 9.0):
        for search in (1.0, -1.0, 2.0, -2.0):
            grid = Grid({(0, 0): 1.0, (5, 0): 2.0, (900_000, 0): 3.0})
            column = Ref(grid, 0, 0, 0, MAX_ROW, 0)
            XMATCH(CTX, key, column, 0.0, search)
            assert grid.reads <= 8, (key, search)
