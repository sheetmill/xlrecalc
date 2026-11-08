"""Iterative calculation against Excel (capture ``iterate_calc``; ARCHITECTURE.md §7.2).

Each variant of ``iterate_calc`` is laid out as an engine whose formula cells hold the cached
values the captured file held, with its ``iterateCount`` and ``iterateDelta``. One recalculation
is a run, as Excel's opening was: its values must be what Excel showed. A second is another run,
as Excel's ``calculate`` and its saving were: its values must be the saved copy's, bit for bit.
An edit variant enters a number, as Excel did, and the run that follows must give what Excel
showed after it.
"""

from __future__ import annotations

import math
import re
from collections import defaultdict
from pathlib import Path

import pytest

from xlrecalc._engine import Engine
from xlrecalc._formula._template import cell_axes
from xlrecalc._store import COLS
from xlrecalc._values import BLANK, ErrorValue, Value

CAPTURES = Path(__file__).parent / "captures"


def rows() -> dict[str, list[list[str]]]:
    by_variant: dict[str, list[list[str]]] = defaultdict(list)
    for line in (CAPTURES / "iterate_calc.tsv").read_text(encoding="utf-8").splitlines():
        if line and not line.startswith("#"):
            fields = line.split("\t")
            by_variant[fields[0]].append(fields)
    return dict(by_variant)


VARIANTS = rows()


def place(ref: str) -> tuple[str, int, int]:
    sheet, cell = ref.split("!")
    r, c = cell_axes(cell, 0, 0)
    return sheet, r.value, c.value


def engine_for(cells: list[list[str]]) -> Engine:
    calc = cells[0][1]
    count = re.search(r'iterateCount="(\d+)"', calc)
    delta = re.search(r'iterateDelta="([^"]+)"', calc)
    sheets: dict[str, dict[str, object]] = {}
    for _, _, ref, formula, cached, *_ in cells:
        sheet, cell = ref.split("!")
        sheets.setdefault(sheet, {})[cell] = "=" + formula if formula else float(cached)
    engine = Engine.from_cells(
        sheets,
        iterate=True,
        iterate_count=int(count[1]) if count else (3 if "edit" in calc else 100),
        iterate_delta=float(delta[1]) if delta else 0.001,
    )
    for _, _, ref, formula, cached, *_ in cells:
        if formula:  # the value iteration starts from
            sheet, row, col = place(ref)
            index = engine.sheet_index(sheet)
            assert index is not None
            engine.sheets[index].values[row * COLS + col] = float(cached)
    return engine


def value(engine: Engine, ref: str) -> Value:
    sheet, row, col = place(ref)
    index = engine.sheet_index(sheet)
    assert index is not None
    return engine.sheets[index].values.get(row * COLS + col, BLANK)


def shown_matches(got: Value, shown: str) -> bool:
    """What AppleScript showed: a number as text, ``missing value`` for an error."""
    if shown == "missing value":
        return isinstance(got, ErrorValue)
    if shown in ("true", "false"):
        return got is (shown == "true")
    try:
        number = float(shown)
    except ValueError:
        return got == shown
    return type(got) is float and math.isclose(got, number, rel_tol=1e-11)


def saved_matches(got: Value, kind: str, saved: str) -> bool:
    if kind == "n":
        return type(got) is float and got == float(saved)
    if kind == "e":
        return isinstance(got, ErrorValue) and got.code == saved
    if kind == "b":
        return got is (saved == "1")
    return got == saved


@pytest.mark.parametrize("variant", [v for v in VARIANTS if not v.startswith("edit_")])
def test_a_run_and_another_give_excels_values(variant: str) -> None:
    cells = VARIANTS[variant]
    engine = engine_for(cells)
    engine.recalculate(full=True)
    first = [(r[2], value(engine, r[2]), r[5]) for r in cells if r[3]]
    assert [f for f in first if not shown_matches(f[1], f[2])] == []
    engine.recalculate()
    second = [(r[2], value(engine, r[2]), r[9], r[10]) for r in cells if r[3]]
    assert [s for s in second if not saved_matches(s[1], s[2], s[3])] == []
    assert not engine.report.uncomputed
    assert not engine.report.circular


@pytest.mark.parametrize("variant", [v for v in VARIANTS if v.startswith("edit_")])
def test_an_edit_runs_once_more(variant: str) -> None:
    cells = VARIANTS[variant]
    engine = engine_for(cells)
    engine.recalculate(full=True)
    assert [r[2] for r in cells if not shown_matches(value(engine, r[2]), r[5])] == []
    target, number = cells[0][1].removeprefix("edit ").split("=")
    sheet, row, col = place(target)
    index = engine.sheet_index(sheet)
    assert index is not None
    engine.set_value(index, row, col, float(number))
    engine.recalculate()
    assert [r[2] for r in cells if not shown_matches(value(engine, r[2]), r[7])] == []


def _public_cases() -> list[tuple[str, str, str, list[list[str]]]]:
    """(suite, variant, form, rows) of the captures of files with no cached values."""
    out = []
    for suite in ("iterate_uncached", "iterate_opening"):
        grouped: dict[tuple[str, str], list[list[str]]] = defaultdict(list)
        for line in (CAPTURES / f"{suite}.tsv").read_text(encoding="utf-8").splitlines():
            if line and not line.startswith("#"):
                fields = line.split("\t")
                if fields[1] in ("flag", "no_flag"):
                    grouped[(fields[0], fields[1])].append(fields)
        out += [(suite, v, f, rows) for (v, f), rows in grouped.items()]
    return out


@pytest.mark.parametrize(
    ("suite", "variant", "form", "cells"),
    _public_cases(),
    ids=[f"{s}:{v}:{f}" for s, v, f, _ in _public_cases()],
)
def test_the_public_path_gives_excels_saved_copy(
    tmp_path: Path, suite: str, variant: str, form: str, cells: list[list[str]]
) -> None:
    """An openpyxl file without cached values, then ``recalculate(path)``: Excel's saved copy.

    Bit for bit. That copy is Excel's opening and its saving: two runs, the first from the empty
    text.
    """
    pytest.importorskip("openpyxl")
    from capture.iterate_probe import build
    from capture.iterate_uncached_probe import CALC

    import xlrecalc
    from xlrecalc._report import column_letters
    from xlrecalc._xlsx import Book

    constants = {"tl_conv": {"B1": "3"}, "tl_interest": {"B1": "1000"}}
    constants["tl_sumself"] = {"A1": "1", "A2": "2", "A3": "3", "A4": "4"}
    sheets: dict[str, list[tuple[str, str, str | None]]] = {}
    for row in cells:
        sheet, ref = row[3].split("!")
        value = constants.get(variant, {}).get(ref, "5")
        sheets.setdefault(sheet, []).append((ref, row[4], None if row[4] else value))
    path = tmp_path / "book.xlsx"
    build(path, CALC[form].format(cells[0][2]), list(sheets.items()))
    report = xlrecalc.recalculate(path)
    assert not report.uncomputed
    with Book(path) as book:  # the written file's cached values, as a reader sees them
        written = {
            f"s!{column_letters(c.column)}{c.row}": ErrorValue(c.value)
            if c.type == "e"
            else c.value
            for c in book.cells("s")
        }
    wrong = [
        (row[3], written.get(row[3]), row[8])
        for row in cells
        if row[4] and not saved_matches(written.get(row[3], BLANK), row[7], row[8])
    ]
    assert wrong == []
