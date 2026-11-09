"""The functions that make arrays against Excel (captures ``dyn_*``; ARCHITECTURE.md §7.7).

Each workbook ``capture/dynamic_probe.py`` gave Excel is rebuilt from the capture's input rows
(data, constants in a result's way, dynamic-array formulas with the dynamic-array metadata,
plain formulas reading them), loaded, recalculated, and every cell Excel saved must hold the
same value here, and every dynamic array must spill over the range Excel recorded. A known
deviation is listed in ``KNOWN`` with its reason, strictly: a listed cell that matches fails.
"""

from __future__ import annotations

import ast
import io
import xml.etree.ElementTree as ET
import zipfile
from pathlib import Path

import pytest
from capture.dynamic_suites import BLOCK, REFUSED

from tests.test_arrays import DOC, _cell, _metadata
from tests.test_captures import same
from tests.test_xlsx import package
from xlrecalc._engine import Engine, node
from xlrecalc._formula import parse
from xlrecalc._formula._nodes import Call, Group, Leaf, Reference
from xlrecalc._functions import SPECS
from xlrecalc._load import load
from xlrecalc._quirks import Quirks
from xlrecalc._values import BLANK, ErrorValue
from xlrecalc._xlsx._write import text_escape

CAPTURES = Path(__file__).with_name("captures")
SUITES = (
    "sequence", "mmult", "unique", "sort", "sortby", "filter",
    "choose", "droptake", "stack", "tocol", "wrap", "index", "matrix", "frequency", "textsplit",
    "randarray",
)  # fmt: skip
RICH = "{http://schemas.microsoft.com/office/spreadsheetml/2017/richdata}"
_ERROR_TYPES = {"8": "#SPILL!", "13": "#CALC!"}

# (suite, cell) -> reason
KNOWN: dict[tuple[str, str], str] = {}
# Cases whose values are random draws, exempt from the cell-by-cell check (each must still be a
# number, as Excel's is); the plain formulas reading them, in column N, count the elements within
# the bounds, and whole ones with the integer flag, and are checked strictly.
RANDOM = "RANDARRAY's values are random draws; its readers check their number, bounds and kind"
RANDOM_CASES = frozenset(("randarray", case) for case in (0, 1, 2, 3, 11, 15))


def _rows(suite: str) -> list[list[str]]:
    text = (CAPTURES / f"dyn_{suite}.tsv").read_text(encoding="utf-8")
    return [line.split("\t") for line in text.splitlines() if not line.startswith("#")]


def _rich_errors(suite: str) -> dict[str, str]:
    """vm (1-based, as cells carry it) -> the error its rich value holds (none without parts)."""
    if not (CAPTURES / f"dyn_{suite}_rdrichvalue.xml").exists():
        return {}
    structures = ET.parse(CAPTURES / f"dyn_{suite}_rdrichvaluestructure.xml").getroot()
    keys = [[k.get("n") for k in s.iter(f"{RICH}k")] for s in structures.iter(f"{RICH}s")]
    values = ET.parse(CAPTURES / f"dyn_{suite}_rdrichvalue.xml").getroot()
    out = {}
    for i, rv in enumerate(values.iter(f"{RICH}rv"), 1):
        fields = dict(
            zip(keys[int(rv.get("s", "0"))], (v.text for v in rv.iter(f"{RICH}v")), strict=False)
        )
        out[str(i)] = _ERROR_TYPES.get(fields.get("errorType") or "", "?")
    return out


def workbook(suite: str) -> bytes:
    """The probe's workbook for a suite, from the capture's input rows."""
    cells: dict[tuple[int, int], str] = {}
    for _, kind, where, text in (r for r in _rows(suite) if r[0] == "input"):
        value = ast.literal_eval(text)
        r, c = _cell(where)
        if kind == "dynamic":
            f = text_escape(value)
            cells[r, c] = f'<c r="{where}" cm="1"><f t="array" ref="{where}">{f}</f></c>'
        elif kind in ("legacy", "reader"):
            cells[r, c] = f'<c r="{where}"><f>{text_escape(value)}</f></c>'
        elif value.startswith("ErrorValue("):
            code = ast.literal_eval(value[len("ErrorValue(") : -1])
            cells[r, c] = f'<c r="{where}" t="e"><v>{text_escape(code)}</v></c>'
        else:
            constant = ast.literal_eval(value)
            if isinstance(constant, bool):
                cells[r, c] = f'<c r="{where}" t="b"><v>{int(constant)}</v></c>'
            elif isinstance(constant, float):
                cells[r, c] = f'<c r="{where}"><v>{constant!r}</v></c>'
            else:
                t = text_escape(constant)
                cells[r, c] = f'<c r="{where}" t="inlineStr"><is><t>{t}</t></is></c>'
    rows: dict[int, list[str]] = {}
    for (r, _c), xml in sorted(cells.items()):
        rows.setdefault(r, []).append(xml)
    sheet_data = "".join(f'<row r="{r}">{"".join(x)}</row>' for r, x in sorted(rows.items()))
    data = package(sheet_data, extra_parts={"xl/metadata.xml": _metadata()})
    with zipfile.ZipFile(io.BytesIO(data)) as z:
        parts = {n: z.read(n) for n in z.namelist()}
    rels = parts["xl/_rels/workbook.xml.rels"].decode()
    rels = rels.replace(
        "</Relationships>",
        f'<Relationship Id="rId8" Type="{DOC}/sheetMetadata" Target="metadata.xml"/>'
        "</Relationships>",
    )
    parts["xl/_rels/workbook.xml.rels"] = rels.encode()
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


def expected(cell: list[str], rich: dict[str, str]) -> object:
    _, _, kind, text, _formula, _t, _ref, _cm, vm = cell
    value = ast.literal_eval(text)
    if kind == "e":
        return ErrorValue(rich[vm] if vm else value)
    if kind == "n":
        return float(value)
    if kind == "b":
        return bool(value)
    return value


def _case_cells() -> list[tuple[str, str]]:
    out = []
    for suite in SUITES:  # results and readers, not the data in columns AA on
        out.extend((suite, r[1]) for r in _rows(suite) if r[0] == "cell" and _cell(r[1])[1] < 27)
    return out


def _case(address: str) -> int:
    return (_cell(address)[0] - 1) // BLOCK


@pytest.mark.parametrize(("suite", "address"), _case_cells(), ids=lambda x: str(x))
def test_cell_matches_excel(suite: str, address: str) -> None:
    e = engine_for(suite)
    cells = {r[1]: r for r in _rows(suite) if r[0] == "cell"}
    row, col = _cell(address)
    tid = e.sheets[0].templates.at(_case(address) * BLOCK, 1)  # the case's formula, in B
    if tid is not None and e.templates[tid].unsupported:
        names = ", ".join(sorted(e.templates[tid].unsupported))
        pytest.skip(f"not implemented: {names}")
    got = e.get(0, row - 1, col - 1)
    want = expected(cells[address], _rich_errors(suite))
    if (suite, _case(address)) in RANDOM_CASES and col < 14:
        assert type(want) is float, (RANDOM, address, want)
        assert type(got) is float, (RANDOM, address, got)
        return
    ok = got is BLANK if want is None else same(got, want)
    reason = KNOWN.get((suite, address))
    if reason is not None:
        assert not ok, f"{suite} {address}: listed ({reason}) but matches Excel"
        return
    assert ok, (suite, address, cells[address][4], got, want)


@pytest.mark.parametrize("suite", SUITES)
def test_spill_ranges_match_excel(suite: str) -> None:
    """Each dynamic array spills over the range Excel recorded in its ``ref``."""
    e = engine_for(suite)
    cells = {r[1]: r for r in _rows(suite) if r[0] == "cell"}
    wrong = []
    for _, kind, where, _text in (r for r in _rows(suite) if r[0] == "input"):
        if kind != "dynamic":
            continue
        r, c = _cell(where)
        tid = e.sheets[0].templates.at(r - 1, c - 1)
        if tid is not None and e.templates[tid].unsupported:
            continue
        first, _, last = cells[where][6].partition(":")
        r2, c2 = _cell(last or first)
        box = e.anchors[node(0, r - 1, c - 1)].box
        if box != (r - 1, c - 1, r2 - 1, c2 - 1) and (suite, where) not in KNOWN:
            wrong.append((where, box, cells[where][6]))
    assert not wrong


def test_refused_forms_are_refused() -> None:
    """A file holding a form Excel refuses keeps the value it holds and reports it."""
    for formula in REFUSED:
        sheet = f'<row r="1"><c r="B1"><f>{text_escape(formula)}</f><v>7</v></c></row>'
        e = load(io.BytesIO(package(sheet)))
        e.recalculate()
        assert e.report.refused, formula
        assert e.get(0, 0, 1) == 7.0, formula


def _outermost(text: str) -> str | None:
    """The function a formula (file form) calls outermost, through parentheses, or ``None``."""
    root = parse(text).root
    while isinstance(root, Group):
        root = root.inner
    if not isinstance(root, Call):
        return None
    return ((root.prefix.text if root.prefix is not None else "") + root.name.text).upper()


_SIZED = ("INDEX", "OFFSET")  # marked or not by their arguments (_workbook._may_be_array)


def test_array_makers_are_flagged_as_the_captures_show() -> None:
    """``makes_array`` holds exactly for the functions the captures show making arrays.

    Evidence for a flag: a dynamic array whose outermost call spilled over more than one cell
    (``dyn_*``), or a typed call of the function on literals and references alone that Excel
    marked although it neither spills nor intersects (``marker_typed``). Evidence against: such
    a call Excel did not mark. ``INDEX`` and ``OFFSET``, marked or not by their arguments, are
    left to ``test_api``'s end-to-end test of every typed formula.
    """
    flagged = {name for name, spec in SPECS.items() if spec.makes_array}
    makes, plain = set(), set()
    for suite in SUITES:
        cells = {r[1]: r for r in _rows(suite) if r[0] == "cell"}
        for _, kind, where, text in (r for r in _rows(suite) if r[0] == "input"):
            if kind != "dynamic" or ":" not in cells[where][6]:
                continue
            name = _outermost(ast.literal_eval(text))
            if name is not None and name not in _SIZED:
                makes.add(name)
    engine = Engine(Quirks.EXCEL)
    engine.add_sheet("typed")
    for line in (CAPTURES / "marker_typed.tsv").read_text(encoding="utf-8").splitlines():
        if line.startswith("#"):
            continue
        typed, form, cm, _kind, ref, _value = map(ast.literal_eval, line.split("\t"))
        if typed.startswith("=@"):  # the file drops an explicit @: =@SEQUENCE(3) is SEQUENCE(3)
            continue
        name = _outermost(form)
        root = parse(form).root
        if name is None or name in _SIZED or not isinstance(root, Call):
            continue
        if not all(isinstance(a, Leaf | Reference) for a in root.args):
            continue
        if cm is None:
            plain.add(name)
        elif ":" not in ref and not engine.templates[engine.template_for(form, 0, 0)].intersects:
            makes.add(name)
    makes &= SPECS.keys()  # a function captured before it is implemented has no flag yet
    assert not makes & plain
    assert flagged == makes, (sorted(flagged - makes), sorted(makes - flagged))
