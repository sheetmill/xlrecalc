"""3-D references (``Jan:Mar!A1:A3``): every sheet of the span, in tab order (capture ``three_d``).

The capture is the acceptance (``tests/test_captures.py``); here, a 3-D reference against the same
reference on each sheet of its span, the dependencies it gives, and a lazy read of a span.
"""

from __future__ import annotations

import io
import random

import pytest
import xlio

import xlrecalc
from tests.xlsx_files import DOC, MAIN, PKG, package
from xlrecalc._engine import Engine, node
from xlrecalc._values import BLANK, ErrorValue

NAMES = ["cases", "Jan", "Feb", "Mar", "Apr", "May"]


def _random_engine(rng: random.Random) -> tuple[Engine, dict[str, dict[str, object]]]:
    sheets: dict[str, dict[str, object]] = {"cases": {}}
    for name in NAMES[1:]:
        cells: dict[str, object] = {}
        for r in range(1, 5):
            for c in "AB":
                roll = rng.random()
                if roll < 0.55:
                    cells[f"{c}{r}"] = float(rng.randint(-50, 50))
                elif roll < 0.7:
                    cells[f"{c}{r}"] = rng.choice(["x", "", "7"])
                elif roll < 0.8:
                    cells[f"{c}{r}"] = rng.random() < 0.5
        sheets[name] = cells
    return Engine.from_cells(sheets), sheets


@pytest.mark.parametrize("seed", range(30))
def test_a_3d_reference_reads_each_sheet_of_its_span(seed: int) -> None:
    """SUM, COUNT, COUNTA, MAX and CONCAT over a span equal the same over each sheet in turn, in
    tab order whichever endpoint comes first; one sheet's range is a span of one."""
    rng = random.Random(seed)
    engine, _ = _random_engine(rng)
    i, j = sorted(rng.sample(range(1, len(NAMES)), 2))
    first, last = (NAMES[i], NAMES[j]) if rng.random() < 0.5 else (NAMES[j], NAMES[i])
    box = rng.choice(["A1:A4", "A1:B4", "B2", "A:A"])
    span = [NAMES[k] for k in range(i, j + 1)]
    pairs = [
        (f"SUM({first}:{last}!{box})", "+".join(f"SUM({s}!{box})" for s in span)),
        (f"COUNT({first}:{last}!{box})", "+".join(f"COUNT({s}!{box})" for s in span)),
        (f"COUNTA({first}:{last}!{box})", "+".join(f"COUNTA({s}!{box})" for s in span)),
        (f"MAX({first}:{last}!{box})", f"MAX({','.join(f'{s}!{box}' for s in span)})"),
        (
            f"_xlfn.CONCAT({first}:{last}!{box})",
            f"_xlfn.CONCAT({','.join(f'{s}!{box}' for s in span)})",
        ),
    ]
    for k, (three_d, per_sheet) in enumerate(pairs):
        engine.set_formula(0, k, 0, three_d)
        engine.set_formula(0, k, 1, per_sheet)
    engine.recalculate()
    for k, (three_d, _) in enumerate(pairs):
        assert engine.get(0, k, 0) == engine.get(0, k, 1), (three_d, seed)
        assert not engine.report.unsupported, dict(engine.report.unsupported)


def test_a_3d_reference_as_a_value_or_with_a_missing_sheet_is_ref() -> None:
    engine = Engine.from_cells(
        {"cases": {"A1": "=Jan:Feb!A1+1", "A2": "=SUM(Jan:Nope!A1)", "A3": "=Jan:Feb!A1"},
         "Jan": {"A1": 1.0}, "Feb": {"A1": 2.0}}
    )  # fmt: skip
    engine.recalculate()
    assert [engine.get(0, r, 0) for r in range(3)] == [ErrorValue.REF] * 3


def test_an_edit_on_a_sheet_of_the_span_wakes_the_formula_and_one_outside_does_not() -> None:
    """The dependencies of a 3-D reference are its box on each sheet of its span: an edit there
    makes the formula stale, an edit on a sheet outside the span or outside the box does not."""
    engine = Engine.from_cells(
        {"cases": {"A1": "=SUM(Jan:Mar!A1:A3)", "B1": "=A1*2"},
         "Jan": {"A1": 1.0}, "Feb": {"A1": 10.0}, "Mar": {"A1": 100.0}, "Apr": {"A1": 1000.0}}
    )  # fmt: skip
    engine.recalculate()
    assert engine.get(0, 0, 1) == 222.0
    formula, dependent = node(0, 0, 0), node(0, 0, 1)
    for sheet, row, wakes in ((2, 1, True), (3, 2, True), (4, 1, False), (1, 3, False)):
        engine.stale.clear()
        engine.set_value(sheet, row, 0, 5.0)
        assert (formula in engine.stale) is wakes, (sheet, row)
        assert (dependent in engine.stale) is wakes, (sheet, row)
        engine.recalculate()
    assert engine.get(0, 0, 0) == 1.0 + 10.0 + 5.0 + 100.0 + 5.0
    engine.set_value(2, 0, 0, BLANK)  # Feb!A1 cleared
    engine.recalculate()
    assert engine.get(0, 0, 1) == 2 * (1.0 + 5.0 + 100.0 + 5.0)


def _book(sheets: dict[str, str]) -> bytes:
    """A package of several sheets, each given as its sheetData."""
    listing = "".join(
        f'<sheet name="{n}" sheetId="{i}" r:id="rId{i}"/>' for i, n in enumerate(sheets, 1)
    )
    rels = "".join(
        f'<Relationship Id="rId{i}" Type="{DOC}/worksheet" Target="worksheets/sheet{i}.xml"/>'
        for i in range(1, len(sheets) + 1)
    )
    parts: dict[str, str | bytes] = {
        "xl/_rels/workbook.xml.rels": f'<Relationships xmlns="{PKG}">{rels}</Relationships>'
    }
    data = list(sheets.values())
    for i, rows in enumerate(data[1:], 2):
        parts[f"xl/worksheets/sheet{i}.xml"] = (
            f'<worksheet xmlns="{MAIN}"><sheetData>{rows}</sheetData></worksheet>'
        )
    return package(data[0], sheets=listing, extra_parts=parts)


def test_a_lazy_read_loads_every_sheet_of_a_span() -> None:
    """Reading one sheet whose formula spans three others loads the middle one too."""
    value = lambda v: f'<row r="1"><c r="A1"><v>{v}</v></c></row>'  # noqa: E731
    data = _book(
        {"cases": '<row r="1"><c r="A1"><f>SUM(Jan:Mar!A1)</f><v>0</v></c></row>',
         "Jan": value(1), "Feb": value(10), "Mar": value(100), "Apr": value(1000)}
    )  # fmt: skip
    with xlio.open(io.BytesIO(data)) as book:
        loader = xlrecalc.Loader(book, recalc_on_load="always")
        list(loader.cells(book.sheets[0]))
        assert loader.workbook().value("cases!A1") == 111.0


def test_a_formula_inside_the_span_is_computed_before_the_3d_reference_reads_it() -> None:
    """The span's formulas are the 3-D reference's precedents: after an edit, Mar!A1 is computed
    first and the SUM reads its new value, on the last sheet of the span as on the first."""
    engine = Engine.from_cells(
        {"cases": {"A1": "=SUM(Jan:Mar!A1)", "B1": 1.0},
         "Jan": {"A1": "=cases!B1*3"}, "Feb": {"A1": 10.0}, "Mar": {"A1": "=cases!B1*2"}}
    )  # fmt: skip
    engine.recalculate()
    assert engine.get(0, 0, 0) == 3.0 + 10.0 + 2.0
    engine.set_value(0, 0, 1, 5.0)
    engine.recalculate(targets=[node(0, 0, 0)])
    assert engine.get(0, 0, 0) == 15.0 + 10.0 + 10.0
