"""Replay the LAMBDA captures: each formula Excel computed, entered as typed, against the engine.

``capture/lambda_probe.py --paste SUITE`` pasted each formula as text into a workbook with
1, 2, 3 in Z1:Z3 and 10, 20, 30 in AA1:AA3 (and the suite's defined names), formula k at
A<5k+1>, and recorded every cell of its result. Here the same workbook is built in file form,
loaded and recalculated, and every recorded cell compared with Excel's value bit for bit.
"""

from __future__ import annotations

import ast
import io
import zipfile
from pathlib import Path

import pytest
from capture._file import text_escape

from tests.test_arrays import DOC, _cell, _metadata
from tests.test_captures import same
from tests.watchdog import bounded, in_child
from tests.xlsx_files import package
from xlrecalc._engine import node
from xlrecalc._load import load
from xlrecalc._values import BLANK, ErrorValue
from xlrecalc._workbook import _may_be_array

CAPTURES = Path(__file__).with_name("captures")
SUITES = ("lambda_calls", "lambda_helpers", "lambda_names")
DEPTH = "lambda_depth"
"""Recursion to Excel's limit and past it: replayed in a child process under the watchdog."""
DATA = {"Z1": 1.0, "Z2": 2.0, "Z3": 3.0, "AA1": 10.0, "AA2": 20.0, "AA3": 30.0}

_BODY = (
    "Excel's recursion budget counts what a body evaluates at its deepest call too: a body"
    " calling SUM with four or eight arguments stops one or two calls earlier; the engine charges"
    " calls only"
)
_TAIL = (
    "Excel runs a call in tail position (nothing after it in the body) past the budget, to between"
    " 8,000 and 10,000 calls; the engine charges it as any other call"
)
KNOWN: dict[tuple[str, str], str] = {
    (DEPTH, "A196"): _BODY,  # CDBIG(5460): SUM of four arguments
    (DEPTH, "A301"): _BODY,  # BSUMEIGHT(5459): SUM of eight
    (DEPTH, "A306"): _BODY,  # BSUMEIGHT(5460)
    (DEPTH, "A476"): _TAIL,  # BPLAIN(5461)
    (DEPTH, "A481"): _TAIL,  # BPLAIN(8000)
}
"""Cells where the engine differs from Excel, with the reason (``QUIRKS.md``). Strict: a listed
cell that matches fails."""
MORE = "lambda_more"
"""The budget's optional parameters and LET names, recursion through MAP, and the choices
without a capture before it: replayed in a child process under the watchdog, as is DEPTH."""
DEEP = (DEPTH, MORE)
BOUNDARY = ("A11", "A16", "A31", "A36", "A66", "A71", "A111", "A116")
"""The rows at the edge of the budget for one, two and three parameters, in the default run."""


def _rows(suite: str) -> list[list[object]]:
    lines = (CAPTURES / f"{suite}.tsv").read_text(encoding="utf-8").splitlines()
    return [[ast.literal_eval(f) for f in line.split("\t")] for line in lines if line[:1] != "#"]


def workbook(suite: str, only: tuple[str, ...] = ()) -> bytes:
    """The probe's workbook for a suite: its data, its names and its formulas in file form
    (``only`` those at these addresses, if given)."""
    cells = {where: f'<c r="{where}"><v>{value!r}</v></c>' for where, value in DATA.items()}
    names = ""
    for row in _rows(suite):
        if row[0] == "name":
            names += f'<definedName name="{row[1]}">{text_escape(str(row[2]))}</definedName>'
            continue
        _, _, _, where, _, _, text, f_t, _, _, _ = row
        if text is None or (only and where not in only):  # spilled into, or left out
            continue
        f = text_escape(str(text))
        if f_t == "array":
            cells[str(where)] = f'<c r="{where}" cm="1"><f t="array" ref="{where}">{f}</f></c>'
        else:
            cells[str(where)] = f'<c r="{where}"><f>{f}</f></c>'
    rows: dict[int, list[tuple[int, str]]] = {}
    for where, xml in cells.items():
        r, c = _cell(where)
        rows.setdefault(r, []).append((c, xml))
    sheet_data = "".join(
        f'<row r="{r}">{"".join(x for _, x in sorted(items))}</row>'
        for r, items in sorted(rows.items())
    )
    data = package(
        sheet_data,
        extra_parts={"xl/metadata.xml": _metadata()},
        workbook_extra=f"<definedNames>{names}</definedNames>" if names else "",
    )
    with zipfile.ZipFile(io.BytesIO(data)) as z:
        parts = {n: z.read(n) for n in z.namelist()}
    rels = parts["xl/_rels/workbook.xml.rels"].decode()
    parts["xl/_rels/workbook.xml.rels"] = rels.replace(
        "</Relationships>",
        f'<Relationship Id="rId8" Type="{DOC}/sheetMetadata" Target="metadata.xml"/>'
        "</Relationships>",
    ).encode()
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as z:
        for n, d in parts.items():
            z.writestr(n, d)
    return buf.getvalue()


_ENGINES: dict[str, object] = {}


def engine_for(suite: str):  # type: ignore[no-untyped-def]
    if suite not in _ENGINES:
        e = load(io.BytesIO(workbook(suite)))
        e.recalculate()
        _ENGINES[suite] = e
    return _ENGINES[suite]


def expected(row: list[object]) -> object:
    """Excel's value: an error is the one its rich value holds (``#CALC!``), if any."""
    kind, value, rich = row[4], row[5], row[10]
    if kind == "e":
        return ErrorValue(str(rich or value))
    if kind == "n":
        return float(value)  # type: ignore[arg-type]
    if kind == "b":
        return bool(value)
    return value


def _cases() -> list[tuple[str, str]]:
    return [(s, str(r[3])) for s in SUITES for r in _rows(s) if r[0] == "cell"]


@pytest.mark.parametrize(("suite", "address"), _cases(), ids=lambda x: str(x))
def test_cell_matches_excel(suite: str, address: str) -> None:
    e = engine_for(suite)
    row = next(r for r in _rows(suite) if r[0] == "cell" and r[3] == address)
    r, c = _cell(address)
    k = int(row[1])  # type: ignore[call-overload]
    tid = e.sheets[0].templates.at(5 * k, 0)  # the case's formula, at A<5k+1>
    if tid is not None and e.templates[tid].unsupported:
        pytest.skip(f"not implemented: {', '.join(sorted(e.templates[tid].unsupported))}")
    got = e.get(0, r - 1, c - 1)
    want = expected(row)
    ok = got is BLANK if want is None else same(got, want)
    reason = KNOWN.get((suite, address))
    if reason is not None:
        assert not ok, f"{suite} {address}: listed ({reason}) but matches Excel"
        return
    assert ok, (suite, address, row[2], got, want)


@pytest.mark.parametrize("suite", SUITES)
def test_spill_ranges_match_excel(suite: str) -> None:
    """Each dynamic array spills over the range Excel recorded in its ``ref``."""
    e = engine_for(suite)
    wrong = []
    for row in _rows(suite):
        if row[0] != "cell" or row[7] != "array":
            continue
        r, c = _cell(str(row[3]))
        tid = e.sheets[0].templates.at(r - 1, c - 1)
        if tid is not None and e.templates[tid].unsupported:
            continue
        first, _, last = str(row[8]).partition(":")
        r2, c2 = _cell(last or first)
        box = e.anchors[node(0, r - 1, c - 1)].box
        if box != (r - 1, c - 1, r2 - 1, c2 - 1) and (suite, row[3]) not in KNOWN:
            wrong.append((row[3], box, row[8]))
    assert not wrong


@pytest.mark.parametrize("suite", SUITES)
def test_every_formula_is_computed(suite: str) -> None:
    """No formula of a LAMBDA capture is kept as not implemented."""
    report = engine_for(suite).recalculate(full=True)
    assert not report.unsupported, dict(report.unsupported)


def _depth_failures(only: tuple[str, ...], suite: str = DEPTH) -> list[str]:
    """A deep suite's cells (``only`` these, if given) that differ from Excel unlisted, or
    match it though listed; and calls reported."""
    e = load(io.BytesIO(workbook(suite, only)))
    report = e.recalculate(full=True)
    problems = [f"reported: {k}" for k in report.unsupported]
    for row in _rows(suite):
        if row[0] != "cell" or (only and row[3] not in only):
            continue
        r, c = _cell(str(row[3]))
        got, want = e.get(0, r - 1, c - 1), expected(row)
        listed = (suite, row[3]) in KNOWN
        if same(got, want) == listed:
            problems.append(f"{row[3]} {row[2]}: {got!r}, Excel {want!r}, listed {listed}")
    return problems


@bounded
def test_recursion_at_the_budget_matches_excel() -> None:
    """One, two and three parameters, each at the last call Excel takes and the first it stops."""
    assert _depth_failures(BOUNDARY) == []


@pytest.mark.slow
@bounded(seconds=600.0)
def test_recursion_matches_excel_at_every_captured_depth() -> None:
    assert _depth_failures(()) == []


@bounded
def test_the_further_lambda_capture_matches_excel() -> None:
    """Capture ``lambda_more``: every cell, recursion with optional parameters and LET names
    to the budget, and through MAP to its limit, included."""
    assert _depth_failures((), MORE) == []


def _outcomes_here() -> list[tuple[str, str, str]]:
    """Each typed formula of the four captures (suite, file form) and how it fares: ``match``
    when every cell of its result matches Excel, ``listed`` when one is a known deviation,
    ``not evaluated`` when the engine reports it, else ``differs``."""
    out = []
    for suite in (*SUITES, *DEEP):
        e = load(io.BytesIO(workbook(suite)))
        report = e.recalculate(full=True)
        reported = {(ref.row, ref.col) for ref in report.uncomputed}
        cases: dict[int, list[list[object]]] = {}
        for row in _rows(suite):
            if row[0] == "cell":
                cases.setdefault(int(row[1]), []).append(row)  # type: ignore[call-overload]
        for rows in cases.values():
            cells = [_cell(str(r[3])) for r in rows]
            if cells[0] in reported:
                outcome = "not evaluated"
            elif any((suite, r[3]) in KNOWN for r in rows):
                outcome = "listed"
            elif all(
                same(e.get(0, r - 1, c - 1), expected(row))
                for (r, c), row in zip(cells, rows, strict=True)
            ):
                outcome = "match"
            else:
                outcome = "differs"
            out.append((suite, str(rows[0][6]), outcome))
    return out


_OUTCOMES: list[tuple[str, str, str]] = []


def formula_outcomes() -> list[tuple[str, str, str]]:
    """``_outcomes_here``, run once in a child process under the watchdog (deep recursion)."""
    if not _OUTCOMES:
        _OUTCOMES.extend(in_child(_outcomes_here, seconds=600.0))
    return _OUTCOMES


def _book(names: dict[str, str], formulas: dict[str, str]) -> bytes:
    """A workbook of plain formulas (file form) and defined names."""
    defined = "".join(
        f'<definedName name="{n}">{text_escape(t)}</definedName>' for n, t in names.items()
    )
    rows = "".join(
        f'<row r="{_cell(w)[0]}"><c r="{w}"><f>{text_escape(f)}</f></c></row>'
        for w, f in sorted(formulas.items(), key=lambda item: _cell(item[0]))
    )
    return package(rows, workbook_extra=f"<definedNames>{defined}</definedNames>")


def _computed(names: dict[str, str], formulas: dict[str, str]) -> tuple[dict[str, object], list]:
    e = load(io.BytesIO(_book(names, formulas)))
    report = e.recalculate(full=True)
    values = {w: e.get(0, _cell(w)[0] - 1, _cell(w)[1] - 1) for w in formulas}
    return values, sorted(report.unsupported)


_RUNAWAY = {
    "TAIL": "_xlfn.LAMBDA(_xlpm.n,TAIL(_xlpm.n+1))",
    "GROW": "_xlfn.LAMBDA(_xlpm.n,1+GROW(_xlpm.n+1))",
    "LETS": "_xlfn.LAMBDA(_xlpm.n,_xlfn.LET(_xlpm.m,_xlpm.n-1,IF(_xlpm.n<=0,0,1+LETS(_xlpm.m))))",
    "ISEV": "_xlfn.LAMBDA(_xlpm.n,IF(_xlpm.n=0,TRUE,ISOD(_xlpm.n-1)))",
    "ISOD": "_xlfn.LAMBDA(_xlpm.n,IF(_xlpm.n=0,FALSE,ISEV(_xlpm.n-1)))",
    "VIAMAP": "_xlfn.LAMBDA(_xlpm.n,IF(_xlpm.n<=0,0,"
    "SUM(_xlfn.MAP({1},_xlfn.LAMBDA(_xlpm.x,1+VIAMAP(_xlpm.n-1))))))",
}


@bounded
def test_runaway_recursion_stops_at_the_budget() -> None:
    """Recursion that never ends stops with ``#NUM!``, on the explicit stack, in bounded memory;
    so does recursion through a ``LET`` binding (a name costing 4, as a parameter) and mutual
    recursion past the budget."""
    values, reported = _computed(
        _RUNAWAY,
        {
            "A1": "TAIL(1)",
            "A2": "GROW(1)",
            "A3": "LETS(4094)",
            "A4": "LETS(4095)",
            "A5": "ISEV(5000)",
            "A6": "ISEV(6000)",
        },
    )
    num = ErrorValue.NUM
    assert values == {"A1": num, "A2": num, "A3": 4094.0, "A4": num, "A5": True, "A6": num}
    assert reported == []


@bounded
def test_recursion_through_a_function_that_calls_a_lambda_is_bounded() -> None:
    """Each ``MAP`` inside the recursion is a call stack inside the one before, so Python's own:
    20 deep computes, past ``NESTED_HELPERS`` it is ``#VALUE!``, as in Excel, however deep."""
    values, reported = _computed(
        _RUNAWAY, {"A1": "VIAMAP(20)", "A2": "VIAMAP(21)", "A3": "VIAMAP(100000)"}
    )
    assert values == {"A1": 20.0, "A2": ErrorValue.VALUE, "A3": ErrorValue.VALUE}
    assert reported == []


_FAMILY = ("LAMBDA", "ISOMITTED", "MAP", "REDUCE", "SCAN", "BYROW", "BYCOL", "MAKEARRAY")


def _prefix_rows() -> list[tuple[str, str, str, str, str]]:
    """Capture ``prefixes``' rows for the LAMBDA family (function, prefix, formula, type,
    value): plain formulas, with and without ``_xlfn.``, over 1, 2, 3 in Z1:Z3."""
    lines = (CAPTURES / "prefixes.tsv").read_text(encoding="utf-8").splitlines()
    rows = [line.split("\t") for line in lines if line[:1] != "#"]
    return [
        (f, ast.literal_eval(p), ast.literal_eval(t), k, v)
        for f, p, t, k, v in rows
        if f in _FAMILY
    ]


@pytest.mark.parametrize("row", _prefix_rows(), ids=lambda r: r[2])
def test_plain_formulas_match_excel_or_are_reported(row: tuple[str, str, str, str, str]) -> None:
    """With ``_xlfn.`` each computes Excel's value; without it, a name Excel does not know, each
    computes Excel's value too (``#NAME?`` where the bare call runs) and the report names the
    cause."""
    _, prefix, text, kind, value = row
    first = f'<row r="1"><c r="A1"><f>{text_escape(text)}</f></c><c r="Z1"><v>1.0</v></c></row>'
    rest = "".join(f'<row r="{i}"><c r="Z{i}"><v>{i}.0</v></c></row>' for i in (2, 3))
    data = package(first + rest)
    e = load(io.BytesIO(data))
    report = e.recalculate(full=True)
    got = e.get(0, 0, 0)
    if prefix:
        assert not report.unsupported, dict(report.unsupported)
        want = ErrorValue(value) if kind == "e" else float(value)
        assert same(got, want), (text, got, want)
        return
    labels = " ".join(report.unsupported)
    assert "written without the _xlfn. prefix Excel requires; Excel shows #NAME?" in labels, labels
    want = ErrorValue(value) if kind == "e" else float(value)
    assert same(got, want), (text, got, want)
    assert (value != "#NAME?") == (text in _BARE_INSIDE), text


_BARE_INSIDE = frozenset({"_xlfn.LAMBDA(_xlpm.x,_xlpm.y,ISOMITTED(_xlpm.y))(1)"})
"""A bare name inside a body Excel never runs (one argument for two parameters: ``#VALUE!``):
the engine gives Excel's value and counts the label, as for any formula holding a bare name."""


@pytest.mark.parametrize(
    ("names", "formula"),
    [
        ({"OUTER": "INNER*2", "INNER": "Data!$B$1"}, "OUTER"),
        ({"TWICEB": "_xlfn.LAMBDA(_xlpm.x,_xlpm.x*Data!$B$1)", "SELF": "TWICEB"}, "SELF(2)"),
    ],
)
def test_an_edit_reaches_a_formula_through_names(names: dict[str, str], formula: str) -> None:
    """A name's dependencies are those of the names it refers to or calls, so an edit of a
    cell they read recalculates the formula."""
    defined = "".join(
        f'<definedName name="{n}">{text_escape(t)}</definedName>' for n, t in names.items()
    )
    row = f'<row r="1"><c r="A1"><f>{text_escape(formula)}</f></c><c r="B1"><v>1</v></c></row>'
    e = load(io.BytesIO(package(row, workbook_extra=f"<definedNames>{defined}</definedNames>")))
    e.recalculate(full=True)
    before = e.get(0, 0, 0)
    e.set_value(0, 0, 1, 5.0)
    e.recalculate()
    assert (before, e.get(0, 0, 0)) == (2.0, 10.0)


def _typed() -> list[tuple[str, str, bool]]:
    """Every typed formula of the LAMBDA captures: suite, file form, and whether Excel marked it."""
    return [
        (suite, str(row[6]), row[7] == "array")
        for suite in (*SUITES, *DEEP)
        for row in _rows(suite)
        if row[0] == "cell" and row[6] is not None
    ]


@pytest.mark.parametrize(("suite", "text", "marked"), _typed(), ids=lambda x: str(x))
def test_marker_as_excel_writes_it_typed(suite: str, text: str, marked: bool) -> None:
    """``set_formula`` marks a formula as a dynamic array exactly where Excel marked it typed."""
    names = {str(r[1]).casefold(): str(r[2]) for r in _rows(suite) if r[0] == "name"}
    assert _may_be_array(text, lambda n: names.get(n.casefold())) is marked
