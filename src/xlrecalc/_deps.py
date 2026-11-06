"""Reverse dependencies from templates and geometry, and dirty sets as runs (ARCHITECTURE.md §6).

No cell stores its dependents. A template's reference covers, for a formula at (row, col), a
box whose corners are offsets from (row, col) or absolute. For a changed cell (r, c), the
formula positions whose box contains it form a rectangle, found per axis:

- both ends relative (offsets d1 <= d2): positions ``[r - d2, r - d1]``;
- both ends absolute (a, b): every position when ``a <= r <= b``, none otherwise;
- one end absolute ``a``, one relative ``d`` (``$A$1:A1``, a growing range): positions
  ``[r - d, MAX]`` when ``r > a``, ``[0, r - d]`` when ``r < a``, every position when ``r = a``.

The formula cells of that template inside the rectangle are its dependents. A template's cells
are kept as row runs per column, so a rectangle yields runs, not cells. Single-cell references
with both axes relative, the common case, are grouped by offset, so a changed cell is probed
once per distinct offset (§6.1).

Dirty sets travel as runs: changing a run of cells in one column dirties runs. A template that
refers to its own column above (``C2 = C1 + B2``, a running total) makes every later cell of its
run depend on any earlier one, so a dirty cell of it dirties the rest of its run at once
(§6.4).
"""

from __future__ import annotations

import bisect
from collections import defaultdict
from collections.abc import Iterable, Iterator
from dataclasses import dataclass
from typing import TYPE_CHECKING

from xlrecalc._store import COLS, MAX_COL, MAX_ROW

if TYPE_CHECKING:
    from xlrecalc._engine import Engine
    from xlrecalc._formula._codegen import RefSpec

__all__ = ["DependencyIndex"]

SHEET_SHIFT = 34
Runs = dict[tuple[int, int], list[tuple[int, int]]]  # (sheet, col) -> sorted disjoint row runs


@dataclass(frozen=True, slots=True)
class _Edge:
    """One reference of one template: the formula sheets it is used on, the sheet it reads."""

    template: int
    target: int | None  # None: the formula's own sheet
    r1: int
    r1_abs: bool
    r2: int
    r2_abs: bool
    c1: int
    c1_abs: bool
    c2: int
    c2_abs: bool


def _axis(
    x1: int, x2: int, d1: int, a1: bool, d2: int, a2: bool, top: int
) -> tuple[int, int] | None:
    """Formula positions on one axis whose box covers [x1, x2] of the changed run."""
    if not a1 and not a2:
        lo, hi = min(d1, d2), max(d1, d2)
        first, last = x1 - hi, x2 - lo
    elif a1 and a2:
        lo, hi = min(d1, d2), max(d1, d2)
        if x2 < lo or x1 > hi:
            return None
        first, last = 0, top
    else:
        a, d = (d1, d2) if a1 else (d2, d1)
        first, last = top + 1, -1
        if x2 > a:  # changed positions above a: formulas at pos >= x - d
            first, last = min(first, max(x1, a + 1) - d), top
        if x1 < a:  # below a: formulas at pos <= x - d
            first, last = 0, max(last, min(x2, a - 1) - d)
        if x1 <= a <= x2:
            first, last = 0, top
    first, last = max(first, 0), min(last, top)
    return (first, last) if first <= last else None


class DependencyIndex:
    """Templates' references and cells, for reverse lookups; built from an engine's state."""

    def __init__(self, engine: Engine) -> None:
        self.engine = engine
        self.offsets: dict[tuple[int, int], list[int]] = defaultdict(list)
        self.edges: list[_Edge] = []
        self.runs: dict[int, Runs] = defaultdict(lambda: defaultdict(list))
        self.sheets_of: dict[int, set[int]] = defaultdict(set)
        self.climbs: set[int] = set()  # templates referring to their own column above
        self._build()

    def _build(self) -> None:
        engine = self.engine
        for s, sheet in enumerate(engine.sheets):
            for col in sorted(sheet.formula_columns):
                templates = sheet.templates
                current_tid, start, previous = -1, -1, -2
                for row in sheet.formula_columns[col]:
                    tid = templates[row * COLS + col]
                    if tid != current_tid or row != previous + 1:
                        if current_tid >= 0:
                            self.runs[current_tid][(s, col)].append((start, previous))
                        current_tid, start = tid, row
                        self.sheets_of[tid].add(s)
                    previous = row
                if current_tid >= 0:
                    self.runs[current_tid][(s, col)].append((start, previous))
        for tid, template in enumerate(engine.templates):
            if tid not in self.runs:
                continue
            for spec in template.refs:
                for edge in self._edges(tid, spec):
                    self.edges.append(edge)
                    single = (
                        edge.target is None
                        and not (edge.r1_abs or edge.r2_abs or edge.c1_abs or edge.c2_abs)
                        and edge.r1 == edge.r2
                        and edge.c1 == edge.c2
                    )
                    if single:
                        self.offsets[(edge.r1, edge.c1)].append(tid)
                        if edge.c1 == 0 and edge.r1 < 0:
                            self.climbs.add(tid)
        self.edges = [e for e in self.edges if not self._is_single(e)]

    @staticmethod
    def _is_single(e: _Edge) -> bool:
        return (
            e.target is None
            and not (e.r1_abs or e.r2_abs or e.c1_abs or e.c2_abs)
            and e.r1 == e.r2
            and e.c1 == e.c2
        )

    def _edges(self, tid: int, spec: RefSpec) -> Iterator[_Edge]:
        engine = self.engine
        if spec.kind == "name":
            for s in self.sheets_of[tid]:
                template = engine._name_template(spec.name or "", s)
                if template is None:
                    continue
                for inner in template.refs:
                    if inner.kind == "box":
                        target = s if inner.sheet is None else engine.sheet_index(inner.sheet)
                        if target is not None:
                            yield _make(tid, target, inner, absolute=True)
            return
        target = None if spec.sheet is None else engine.sheet_index(spec.sheet)
        if spec.sheet is not None and target is None:
            return
        yield _make(tid, target, spec, absolute=False)

    # Queries ---------------------------------------------------------------------------------

    def _template_runs(
        self, tid: int, sheet: int, rows: tuple[int, int], cols: tuple[int, int]
    ) -> Iterator[tuple[int, int, int, int]]:
        """Runs (sheet, col, first, last) of a template's cells inside a rectangle."""
        by_column = self.runs[tid]
        c1, c2 = cols
        r1, r2 = rows
        keys = (
            [(sheet, c) for c in range(c1, c2 + 1) if (sheet, c) in by_column]
            if c2 - c1 < len(by_column)
            else [k for k in by_column if k[0] == sheet and c1 <= k[1] <= c2]
        )
        for key in keys:
            runs = by_column[key]
            i = max(bisect.bisect_right(runs, (r1, MAX_ROW + 1)) - 1, 0)
            while i < len(runs) and runs[i][0] <= r2:
                a, b = runs[i]
                lo, hi = max(a, r1), min(b, r2)
                if lo <= hi:
                    if tid in self.climbs:
                        hi = b  # a running total: the rest of the run depends on it
                    yield key[0], key[1], lo, hi
                i += 1

    def direct(
        self, sheet: int, col: int, first: int, last: int
    ) -> Iterator[tuple[int, int, int, int]]:
        """Runs of formula cells reading any cell of the run (sheet, col, first..last)."""
        for (dr, dc), tids in self.offsets.items():
            c = col - dc
            if not 0 <= c <= MAX_COL:
                continue
            rows = (max(first - dr, 0), min(last - dr, MAX_ROW))
            if rows[0] > rows[1]:
                continue
            for tid in tids:
                yield from self._template_runs(tid, sheet, rows, (c, c))
        for e in self.edges:
            if e.target is not None and e.target != sheet:
                continue
            row_span = _axis(first, last, e.r1, e.r1_abs, e.r2, e.r2_abs, MAX_ROW)
            if row_span is None:
                continue
            col_span = _axis(col, col, e.c1, e.c1_abs, e.c2, e.c2_abs, MAX_COL)
            if col_span is None:
                continue
            formula_sheets = [sheet] if e.target is None else sorted(self.sheets_of[e.template])
            for s in formula_sheets:
                yield from self._template_runs(e.template, s, row_span, col_span)

    def dirty(self, cells: Iterable[int], ranges: Iterable[tuple[int, int, int, int]] = ()) -> Runs:
        """Every formula cell depending, directly or not, on the cells, as runs.

        ``ranges`` adds whole runs of changed cells, as (sheet, col, first row, last row).
        """
        found: Runs = defaultdict(list)
        self.spread(cells, ranges, found)
        return found

    def spread(
        self,
        cells: Iterable[int],
        ranges: Iterable[tuple[int, int, int, int]],
        found: Runs,
    ) -> list[tuple[int, int, int, int]]:
        """Add to ``found`` what depends on the cells; returns the runs it did not hold yet.

        What ``found`` already holds is not walked again, so that many calls sharing one
        ``found`` cost what one walk over all their cells costs.
        """
        new: list[tuple[int, int, int, int]] = []
        queue: list[tuple[int, int, int, int]] = list(ranges)
        for n in cells:
            sheet = n >> SHEET_SHIFT
            row, col = divmod(n & ((1 << SHEET_SHIFT) - 1), COLS)
            queue.append((sheet, col, row, row))
        while queue:
            sheet, col, first, last = queue.pop()
            for s, c, a, b in self.direct(sheet, col, first, last):
                for lo, hi in _add(found[(s, c)], a, b):
                    run = (s, c, lo, hi)
                    new.append(run)
                    queue.append(run)
        return new


def _make(tid: int, target: int | None, spec: RefSpec, absolute: bool) -> _Edge:
    r1, c1, r2, c2 = spec.r1, spec.c1, spec.r2, spec.c2
    assert r1 is not None
    assert c1 is not None
    assert r2 is not None
    assert c2 is not None
    return _Edge(
        tid,
        target,
        r1.value,
        r1.absolute or absolute,
        r2.value,
        r2.absolute or absolute,
        c1.value,
        c1.absolute or absolute,
        c2.value,
        c2.absolute or absolute,
    )


def _add(runs: list[tuple[int, int]], a: int, b: int) -> list[tuple[int, int]]:
    """Merge [a, b] into sorted disjoint runs; return the parts that were new."""
    new: list[tuple[int, int]] = []
    i = bisect.bisect_left(runs, (a, -1))
    if i > 0 and runs[i - 1][1] >= a - 1:
        i -= 1
    lo, hi = a, b
    cursor = a
    j = i
    while j < len(runs) and runs[j][0] <= b + 1:
        ra, rb = runs[j]
        if ra > cursor:
            new.append((cursor, min(ra - 1, b)))
        cursor = max(cursor, rb + 1)
        lo, hi = min(lo, ra), max(hi, rb)
        j += 1
    if cursor <= b:
        new.append((cursor, b))
    runs[i:j] = [(lo, hi)]
    return new
