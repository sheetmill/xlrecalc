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

from tests.test_arrays import DOC, _cell, _metadata
from tests.test_captures import same
from tests.test_xlsx import package
from xlrecalc._engine import node
from xlrecalc._load import load
from xlrecalc._values import BLANK, ErrorValue
from xlrecalc._xlsx._write import text_escape

CAPTURES = Path(__file__).with_name("captures")
SUITES = ("lambda_calls",)
DATA = {"Z1": 1.0, "Z2": 2.0, "Z3": 3.0, "AA1": 10.0, "AA2": 20.0, "AA3": 30.0}

KNOWN: dict[tuple[str, str], str] = {}
"""Cells where the engine differs from Excel, with the reason (``QUIRKS.md``). Strict: a listed
cell that matches fails."""


def _rows(suite: str) -> list[list[object]]:
    lines = (CAPTURES / f"{suite}.tsv").read_text(encoding="utf-8").splitlines()
    return [[ast.literal_eval(f) for f in line.split("\t")] for line in lines if line[:1] != "#"]


def workbook(suite: str) -> bytes:
    """The probe's workbook for a suite: its data, its names and its formulas in file form."""
    cells = {where: f'<c r="{where}"><v>{value!r}</v></c>' for where, value in DATA.items()}
    names = ""
    for row in _rows(suite):
        if row[0] == "name":
            names += f'<definedName name="{row[1]}">{text_escape(str(row[2]))}</definedName>'
            continue
        _, _, _, where, _, _, text, f_t, _, _, _ = row
        if text is None:  # a cell the formula above spilled into
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
