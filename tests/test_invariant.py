"""After a full recalculation, every formula cell has a new value or is reported.

Every formula of the capture suites and the array captures goes into a file with a sentinel as
its cached value; ``recalculate`` runs in full, and the file it wrote is read back. A cell still
holding the sentinel must be in the report. A file's old value left in place and not reported is
the failure this looks for: it is a wrong answer nobody is told about.
"""

from __future__ import annotations

import dataclasses
import tempfile
from pathlib import Path
from xml.sax.saxutils import escape

import pytest

import xlrecalc
from tests import test_array_refs, test_arrays
from tests.captures import Formula, load
from tests.formula_corpus import CORPUS as FORMULAS
from tests.test_array_limit import two_sheets
from tests.test_captures import SUITES, column_letters, formula_text
from tests.test_xlsx import MAIN, package
from xlrecalc._formula import FormulaSyntaxError, TokenizeError, parse
from xlrecalc._formula._nodes import Call, Invoke, Union
from xlrecalc._formula._template import cell_axes
from xlrecalc._values import BLANK, ErrorValue
from xlrecalc._xlsx import Book

SENTINEL = 987654.125


def _cell(address: str, value: object) -> str:
    if isinstance(value, Formula):
        return f'<c r="{address}"><f>{escape(value.text[1:])}</f><v>{SENTINEL}</v></c>'
    if value is None or value is BLANK:
        return ""
    if isinstance(value, bool):
        return f'<c r="{address}" t="b"><v>{int(value)}</v></c>'
    if isinstance(value, ErrorValue):
        return f'<c r="{address}" t="e"><v>{value.code}</v></c>'
    if isinstance(value, str):
        text = f'<is><t xml:space="preserve">{escape(value)}</t></is>'
        return f'<c r="{address}" t="inlineStr">{text}</c>'
    return f'<c r="{address}"><v>{float(value)!r}</v></c>'  # type: ignore[arg-type]


def suite_file(name: str) -> bytes:
    """A capture suite laid out as ``tests/test_captures.py`` lays it out, in a file."""
    suite = load(name)
    cells: dict[tuple[int, int], str] = {}
    for ref, value in suite.fixed.items():
        r, c = cell_axes(ref, 0, 0)
        cells[r.value, c.value] = _cell(ref, value)
    for k, row in enumerate(suite.rows, 1):
        for i, value in enumerate(row.inputs, 1):
            cells[k - 1, i] = _cell(f"{column_letters(i + 1)}{k}", value)
        text = formula_text(row, k)
        cells[k - 1, 0] = _cell(f"A{k}", Formula("=" + text))
    rows: dict[int, list[str]] = {}
    for (r, _c), xml in sorted(cells.items()):
        rows.setdefault(r, []).append(xml)
    data = "".join(f'<row r="{r + 1}">{"".join(x)}</row>' for r, x in sorted(rows.items()))
    pr = '<workbookPr date1904="1"/>' if suite.date1904 else ""
    sheet = f'<worksheet xmlns="{MAIN}"><sheetData>{data}</sheetData></worksheet>'
    return package(sheet_xml=sheet, workbook_extra=pr)


def stale_and_unreported(data: bytes) -> tuple[list[str], int]:
    """Cells still holding the sentinel that the report does not cover, and how many it does.

    A reported array formula, named by its first cell, covers the cells of its range: one refused
    for its size may declare billions of them.
    """
    with tempfile.TemporaryDirectory() as folder:
        path = Path(folder) / "book.xlsx"
        path.write_bytes(data)
        report = xlrecalc.recalculate(path, full=True)
        reported = {
            str(c)
            for c in report.uncomputed | report.circular | report.from_kept | set(report.refused)
        }
        stale = []
        kept = 0
        with Book(path) as book:
            for name in (s.name for s in book.sheets):
                ranges = []
                held = []
                for c in book.cells(name):
                    where = f"{name}!{column_letters(c.column)}{c.row}"
                    if c.formula and c.formula.ref and where in reported:
                        ranges.append(_span(c.formula.ref))
                    if c.value == SENTINEL:
                        held.append((where, c.row, c.column, c.formula))
                for where, row, column, formula in held:
                    covered = any(
                        r1 <= row <= r2 and c1 <= column <= c2 for r1, c1, r2, c2 in ranges
                    )
                    if where in reported or (formula is None and covered):
                        kept += 1
                    else:
                        stale.append(f"{where}: {formula.text if formula else ''}")
    return stale, kept


def _span(ref: str) -> tuple[int, int, int, int]:
    first, _, last = ref.partition(":")
    top, left = cell_axes(first, 0, 0)
    bottom, right = cell_axes(last or first, 0, 0)
    return top.value + 1, left.value + 1, bottom.value + 1, right.value + 1


@pytest.mark.parametrize("name", SUITES)
def test_a_capture_suite_leaves_no_old_value_unreported(name):
    stale, _ = stale_and_unreported(suite_file(name))
    assert stale[:10] == []


def test_an_array_of_references_read_by_a_function_is_reported():
    data = two_sheets("N(OFFSET(D!A1,{0;1},0))", cached=str(SENTINEL))
    stale, kept = stale_and_unreported(data)
    assert (stale, kept) == ([], 1)


def test_arrays_leave_no_old_value_unreported():
    stale, _ = stale_and_unreported(test_arrays.workbook(cached=str(SENTINEL)))
    assert stale[:10] == []


@pytest.mark.slow
def test_array_refs_leave_no_old_value_unreported():
    stale, kept = stale_and_unreported(test_array_refs.workbook(cached=str(SENTINEL)))
    assert stale[:10] == []
    assert kept > 0  # the arrays of references in test_array_refs.KNOWN are reported


NOT_IMPLEMENTED = [
    "SUM((D!A1,D!B2))",
    "COUNT((D!A1,D!B2))",
    "LARGE((D!A1,D!B2),1)",
    "(D!A1,D!B2)",
    "SUM(D!A1:A2,(D!A1,D!B2))",
    "SUM(1)(2)",
    "SUM(S:D!A1)",
    "SUM(Table1[a])",
    "_xlfn.MAP(D!A1:A2,_xlfn.LAMBDA(_xlpm.v,_xlpm.v*2))",
    "N(OFFSET(D!A1,{0;1},0))",
]


@pytest.mark.parametrize("formula", NOT_IMPLEMENTED)
def test_a_form_not_implemented_keeps_its_value_and_is_reported(formula):
    """A wrong value would not hold the sentinel: these must keep it, not merely be listed."""
    stale, kept = stale_and_unreported(two_sheets(escape(formula), cached=str(SENTINEL)))
    assert (stale, kept) == ([], 1)


def _has_form_not_implemented(node: object) -> bool:
    """Whether a parsed formula has a union (not as ``ROW``'s reference) or a call of a result."""
    stack = [node]
    while stack:
        item = stack.pop()
        if isinstance(item, Invoke):
            return True
        if isinstance(item, Call) and item.name.text.upper() in ("ROW", "COLUMN"):
            continue
        if isinstance(item, Union):
            return True
        if isinstance(item, list | tuple):
            stack.extend(item)
        elif dataclasses.is_dataclass(item):
            stack.extend(getattr(item, f.name) for f in dataclasses.fields(item))
    return False


def test_every_such_form_in_the_formula_corpus_is_reported():
    """Generated from the corpus's parse trees, so a new form of this kind cannot be missed."""
    formulas = []
    for text in FORMULAS:
        try:
            if _has_form_not_implemented(parse(text).root):
                formulas.append(text)
        except (FormulaSyntaxError, TokenizeError):
            continue
    assert formulas
    wrong = []
    for text in formulas:
        data = two_sheets(escape(text), cached=str(SENTINEL))
        with tempfile.TemporaryDirectory() as folder:
            path = Path(folder) / "book.xlsx"
            path.write_bytes(data)
            report = xlrecalc.recalculate(path, full=True)
        if not (report.uncomputed or report.from_kept or report.refused or report.circular):
            wrong.append(text)
    assert wrong == []
