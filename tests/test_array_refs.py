"""Functions taking a reference that may return an array, against Excel (capture ``array_refs``).

The workbook ``capture/array_ref_probe.py`` gave Excel is rebuilt from the capture's input rows,
loaded and recalculated, and every cell Excel saved must hold the same value, but for the cells of
formulas listed in ``KNOWN`` with their reason (strict: a listed formula that matches fails).
Formulas making arrays of a million rows, about 40 of them, run only with ``--tier excel``; two
of them run in every test run, under the memory watchdog.
"""

from __future__ import annotations

import ast
import functools
import io
import re
import zipfile
from pathlib import Path

import pytest

import xlrecalc
from tests.test_array_limit import two_sheets
from tests.test_captures import column_letters, same
from tests.test_xlsx import MAIN, PKG, package
from tests.watchdog import GIB, bounded, in_child
from xlrecalc._engine import Engine, node
from xlrecalc._formula._tokens import column_number
from xlrecalc._load import load
from xlrecalc._values import BLANK, ErrorValue
from xlrecalc._xlsx._write import text_escape

CAPTURES = Path(__file__).with_name("captures")
DOC = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
ROWS = [
    line.rstrip("\n").split("\t")
    for line in (CAPTURES / "array_refs.tsv").read_text(encoding="utf-8").splitlines()
    if not line.startswith("#")
]
INPUTS = [(r[1], r[2], ast.literal_eval(r[3])) for r in ROWS if r[0] == "input"]
CELLS = {r[1]: r[2:] for r in ROWS if r[0] == "cell"}
NOT_IMPLEMENTED = ("_xlfn.ISFORMULA", "_xlfn.FORMULATEXT", "_xlfn.SHEET", "ISREF", "AREAS")

_REFS = "an array of references read by a function: not implemented, and the cell is reported"
OFFSETS = "OFFSET(Z1,{0;1;2;3},0)"
TEXTS = 'INDIRECT({"Z1";"Z2";"Z3";"Z4"})'
KNOWN: dict[str, str] = dict.fromkeys(
    (
        f"N({OFFSETS})",
        f"SUM(N({OFFSETS}))",
        f"SUMPRODUCT(N({OFFSETS}))",
        f"SUBTOTAL(9,{OFFSETS})",
        f"SUMPRODUCT(SUBTOTAL(9,{OFFSETS}))",
        f'COUNTIF({OFFSETS},">0")',
        f'SUMPRODUCT(COUNTIF({OFFSETS},">0"))',
        f"SUM({OFFSETS})",
        f"ROW({OFFSETS})",
        f"N({TEXTS})",
        f"SUMPRODUCT(N({TEXTS}))",
        f'SUMPRODUCT(COUNTIF({TEXTS},">0"))',
        "SUM(INDEX(Z1:Z4,ROW(Z1:Z2)))",
        "SUMPRODUCT(_xlfn.XLOOKUP(Z1:Z4,Z1:Z4,AA1:AA4))",
        "SUM(_xlfn.XLOOKUP(Z1:Z4,Z1:Z4,AA1:AA4))",
        "SUMPRODUCT(INDEX(Z1:Z4,ROW(Z1:Z2)))",
    ),
    _REFS,
)


def _cell(ref: str) -> tuple[int, int]:
    m = re.fullmatch(r"\$?([A-Z]+)\$?(\d+)", ref)
    assert m is not None
    return int(m[2]), column_number(m[1])


def _data() -> dict[tuple[int, int], str]:
    cells = {}
    for r, (z, aa, ab) in enumerate(
        zip([1, 2, 2, 3, 5], range(10, 60, 10), "abcde", strict=True), 1
    ):
        cells[r, 26] = f'<c r="Z{r}"><v>{z}</v></c>'
        cells[r, 27] = f'<c r="AA{r}"><v>{aa}</v></c>'
        cells[r, 28] = f'<c r="AB{r}" t="inlineStr"><is><t>{ab}</t></is></c>'
    return cells


def _sheet(cells: dict[tuple[int, int], str]) -> str:
    rows: dict[int, list[str]] = {}
    for (r, _c), xml in sorted(cells.items()):
        rows.setdefault(r, []).append(xml)
    data = "".join(f'<row r="{r}">{"".join(x)}</row>' for r, x in sorted(rows.items()))
    return f'<worksheet xmlns="{MAIN}"><sheetData>{data}</sheetData></worksheet>'


def _large(formula: str) -> bool:
    return "ROW(" in formula and any(
        f"({ref})" in formula for ref in ("Z:Z", "1:1048576", "A:XFD", "A:J")
    )


def workbook(large: bool = False, cached: str | None = None) -> bytes:
    """The probe's workbook: the data on two sheets, the name, and the input formulas.

    With ``cached``, every formula and every cell of a declared array holds that as its value.
    """
    cells = _data()
    v = f"<v>{cached}</v>" if cached else ""
    for kind, where, formula in INPUTS:
        if _large(formula) and not large:
            continue
        text = text_escape(formula)
        first = where.split(":")[0]
        r, c = _cell(first)
        if kind == "cse":
            r1, c1, r2, c2 = _span(where)
            for member_row in range(r1, r2 + 1):
                for member_col in range(c1, c2 + 1):
                    address = f"{column_letters(member_col)}{member_row}"
                    cells[member_row, member_col] = f'<c r="{address}">{v}</c>'
            cells[r, c] = f'<c r="{first}"><f t="array" ref="{where}">{text}</f>{v}</c>'
        elif kind == "dynamic":
            cells[r, c] = f'<c r="{first}" cm="1"><f t="array" ref="{first}">{text}</f>{v}</c>'
        else:
            cells[r, c] = f'<c r="{first}"><f>{text}</f>{v}</c>'
    rels = (
        f'<Relationships xmlns="{PKG}">'
        f'<Relationship Id="rId1" Type="{DOC}/worksheet" Target="worksheets/sheet1.xml"/>'
        f'<Relationship Id="rId3" Type="{DOC}/worksheet" Target="worksheets/sheet2.xml"/>'
        f'<Relationship Id="rId8" Type="{DOC}/sheetMetadata" Target="metadata.xml"/>'
        "</Relationships>"
    )
    data = package(
        sheet_xml=_sheet(cells),
        sheets='<sheet name="probe" sheetId="1" r:id="rId1"/>'
        '<sheet name="other" sheetId="2" r:id="rId3"/>',
        extra_parts={
            "xl/_rels/workbook.xml.rels": rels,
            "xl/worksheets/sheet2.xml": _sheet(_data()),
            "xl/metadata.xml": (CAPTURES / "arrays_metadata.xml").read_text(encoding="utf-8"),
        },
    )
    with zipfile.ZipFile(io.BytesIO(data)) as z:
        parts = {n: z.read(n) for n in z.namelist()}
    name = '<definedNames><definedName name="Block">probe!$Z$1:$Z$4</definedName></definedNames>'
    parts["xl/workbook.xml"] = parts["xl/workbook.xml"].replace(
        b"</sheets>", b"</sheets>" + name.encode()
    )
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as z:
        for n, d in parts.items():
            z.writestr(n, d)
    return buf.getvalue()


def _span(where: str) -> tuple[int, int, int, int]:
    first, _, last = where.partition(":")
    r1, c1 = _cell(first)
    r2, c2 = _cell(last or first)
    return r1, c1, r2, c2


def _owners() -> dict[tuple[int, int], str]:
    """Each cell of the workbook and the formula that fills it, as Excel filled it."""
    owners = {}
    for kind, where, formula in INPUTS:
        address = where.split(":")[0]
        if kind == "dynamic" and address in CELLS:
            where = CELLS[address][4] or where  # the range Excel spilled into
        r1, c1, r2, c2 = _span(where)
        for r in range(r1, r2 + 1):
            for c in range(c1, c2 + 1):
                owners[r, c] = formula
    return owners


def _expected(cell: list[str]) -> object:
    kind, text = cell[0], cell[1]
    value = ast.literal_eval(text)
    if kind == "e":
        return ErrorValue(value)
    if kind == "n":
        return float(value)
    return value


def _not_implemented(formula: str) -> bool:
    return any(name in formula for name in NOT_IMPLEMENTED)


@functools.cache
def _engine(large: bool) -> Engine:
    engine = load(io.BytesIO(workbook(large)))
    engine.recalculate()
    return engine


def differences(large: bool = False) -> dict[str, list[str]]:
    """Formula to the cells where the engine differs from Excel."""
    engine = _engine(large)
    owners = _owners()
    out: dict[str, list[str]] = {}
    for address, cell in CELLS.items():
        r, c = _cell(address)
        formula = owners.get((r, c))
        if formula is None or _not_implemented(formula) or (_large(formula) and not large):
            continue
        got = engine.get(0, r - 1, c - 1)
        want = _expected(cell)
        if want is None:
            if got is not BLANK:
                out.setdefault(formula, []).append(f"{address}: {got!r}, not blank")
        elif not same(got, want):
            out.setdefault(formula, []).append(f"{address}: {got!r}, Excel {want!r}")
    return out


def _check(found: dict[str, list[str]]) -> None:
    assert {f: v[:3] for f, v in found.items() if f not in KNOWN} == {}
    assert [f for f in KNOWN if f not in found] == []


@pytest.mark.excel
def test_every_cell_matches_excel_or_is_listed():
    _check(differences())


@pytest.mark.excel
def test_rows_of_a_million_match_excel(tier):
    if tier != "excel":
        pytest.skip("a million rows, about 40 times: run with --tier excel")
    _check(in_child(differences, True, limit=4 * GIB, seconds=300))


@pytest.mark.parametrize(
    ("formula", "array"),
    [("SUMPRODUCT(ROW(1:1048576))", None), ("SUM(ROW(1:1048576))", "A1:A1")],
)
@bounded
def test_a_million_rows_in_an_array_context(formula, array):
    book = xlrecalc.load(io.BytesIO(two_sheets(formula, array=array)))
    book.recalculate(full=True)
    assert book.value("S!A1") == 549756338176.0


@pytest.mark.excel
def test_spill_ranges_match_excel():
    engine = _engine(False)
    wrong = []
    for kind, anchor, formula in INPUTS:
        if kind != "dynamic" or formula in KNOWN or _not_implemented(formula) or _large(formula):
            continue
        r, c = _cell(anchor)
        filling = engine.anchors.get(node(0, r - 1, c - 1))
        r1, c1, r2, c2 = _span(CELLS[anchor][4] or anchor)
        if filling is not None and filling.box != (r1 - 1, c1 - 1, r2 - 1, c2 - 1):
            wrong.append((formula, filling.box))
    assert wrong == []
