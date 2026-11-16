"""When a load recalculates, against capture ``open_calc`` (Excel 16.115 opening each variant).

Each row is a workbook whose cached values were planted wrong, with one change to its ``calcPr``
or cells. Here the same cells and ``calcPr`` are written, loaded with the default
``recalc_on_load="excel"``, and each cell compared with the copy Excel saved after opening it:
a planted value kept where Excel kept it, the right value where Excel recalculated.
"""

from __future__ import annotations

import ast
import io
import re
import warnings
from pathlib import Path

import pytest

import xlrecalc
from tests.test_captures import same
from tests.xlsx_files import package
from xlrecalc._load import VERIFIED_CALC_ID
from xlrecalc._values import ErrorValue

CAPTURE = Path(__file__).with_name("captures") / "open_calc.tsv"
FORMULAS = {
    "A2": "A1*3",
    "A3": '"x"&amp;A1',
    "A4": "A1&gt;1",
    "A5": "1/0",
    "A6": "SUM(A1:A2)",
    "A7": "RAND()*0+A1",
}


def _rows() -> list[list[str]]:
    lines = CAPTURE.read_text(encoding="utf-8").splitlines()
    return [line.split("\t") for line in lines if line and not line.startswith("#")]


def _cells(text: str) -> dict[str, tuple[str | None, str | None]]:
    """``A2=999 A3=str:SENT A4=b:0 A7=-``: each cell's type and value text (``-``: none)."""
    out: dict[str, tuple[str | None, str | None]] = {}
    for item in text.split():
        ref, _, rest = item.partition("=")
        kind, _, value = rest.rpartition(":") if ":" in rest[:4] else ("", "", rest)
        out[ref] = (kind or None, None if value == "-" else value)
    return out


def _value(kind: str | None, text: str) -> object:
    if kind == "str":
        return text
    if kind == "b":
        return text == "1"
    if kind == "e":
        return ErrorValue(text)
    return float(text)


def _workbook(row: list[str]) -> bytes:
    calc = row[1] if row[1].startswith("<calcPr") else ""
    cells = []
    for ref, (kind, value) in _cells(row[2]).items():
        t = f' t="{kind}"' if kind and kind != "n" else ""
        v = f"<v>{value}</v>" if value is not None else ""
        f = f"<f>{FORMULAS[ref]}</f>" if ref in FORMULAS else ""
        cells.append(f'<row r="{ref[1:]}"><c r="{ref}"{t}>{f}{v}</c></row>')
    return package("".join(cells), workbook_extra=calc)


def _loaded(row: list[str], **options: object) -> dict[str, object]:
    wb = xlrecalc.load(io.BytesIO(_workbook(row)), **options)  # type: ignore[arg-type]
    return {ref: wb.value(ref) for ref in _cells(row[2]) if ref in FORMULAS}


def _shown(row: list[str]) -> list[str]:
    return ast.literal_eval("[" + row[3].replace("' '", "', '") + "]")  # type: ignore[no-any-return]


@pytest.mark.parametrize("row", _rows(), ids=lambda r: r[0])
def test_a_load_recalculates_where_excel_does_on_opening(row: list[str]) -> None:
    """A load gives what Excel shows on opening: the planted values or the right ones."""
    kept = _shown(row)[0] == "999"
    cells = _cells(row[2] if kept else row[5])
    got = _loaded(row)
    want = {ref: _value(*cells[ref]) for ref in got}  # type: ignore[arg-type]
    assert all(same(got[r], want[r]) for r in got), (row[0], got, want)


@pytest.mark.parametrize("row", _rows(), ids=lambda r: r[0])
def test_recalculate_writes_what_excel_saves_after_opening(tmp_path: Path, row: list[str]) -> None:
    """``recalculate(path)`` is an opening and a save: Excel's saved copy, manual mode included."""
    path = tmp_path / "book.xlsx"
    path.write_bytes(_workbook(row))
    xlrecalc.recalculate(path, full_calc_on_load=False)
    wb = xlrecalc.load(path, recalc_on_load="never")
    saved = _cells(row[5])
    for ref in (r for r in saved if r in FORMULAS and r != "A7"):  # A7: RAND, drawn anew
        assert same(wb.value(ref), _value(*saved[ref])), (row[0], ref, wb.value(ref))


@pytest.mark.parametrize("row", _rows(), ids=lambda r: r[0])
def test_never_and_always(row: list[str]) -> None:
    """``"never"`` keeps every planted value (a formula without one is computed, and a volatile
    one, as at every calculation); ``"always"`` computes every formula."""
    planted = _cells(row[2])
    right = _cells(next(r for r in _rows() if r[0] == "no_calcpr")[5])
    never, always = _loaded(row, recalc_on_load="never"), _loaded(row, recalc_on_load="always")
    for ref in never:
        kind, value = planted[ref]
        if value is not None and ref != "A7":  # A7 is volatile
            assert same(never[ref], _value(kind, value)), (row[0], ref, never[ref])
        if ref != "A7":
            assert same(always[ref], _value(*right[ref])), (row[0], ref, always[ref])  # type: ignore[arg-type]


def test_the_threshold_is_the_capture_s() -> None:
    """``VERIFIED_CALC_ID`` is the lowest calcId the capture shows Excel trusting."""
    trusted = [
        int(m[1])
        for r in _rows()
        if (m := re.search(r'calcId="(\d+)"', r[1]))
        and "fullCalcOnLoad" not in r[1]
        and "calcMode" not in r[1]
        and r[0].startswith(("calcid", "patched"))
        and _cells(r[5])["A2"][1] == "999"
    ]
    assert min(trusted) == VERIFIED_CALC_ID


MANUAL_WITH_A_REASON = frozenset({"manual_fullcalconload", "manual_calcid_124519"})


def test_displayed_values_are_the_saved_ones_but_in_manual_mode_with_a_reason() -> None:
    """What Excel shows on opening is what it saves after it, except in manual calculation with a
    reason to recalculate: it shows the cached values and saves them recalculated."""
    for row in _rows():
        shown_kept = _shown(row)[0] == "999"
        saved_kept = _cells(row[5])["A2"][1] == "999"
        assert (shown_kept != saved_kept) == (row[0] in MANUAL_WITH_A_REASON), row[0]


@pytest.mark.parametrize(
    ("call", "alias", "mode"),
    [
        ("load", {"trust_cached": False}, "always"),
        ("load", {"trust_cached": True}, "excel"),
        ("recalculate", {"full": True}, "always"),
        ("recalculate", {"full": False}, "excel"),
    ],
)
def test_the_deprecated_aliases_warn_and_map(
    tmp_path: Path, call: str, alias: dict[str, bool], mode: str
) -> None:
    row = next(r for r in _rows() if r[0] == "patched")  # trusted on opening: 999 kept
    path = tmp_path / "book.xlsx"
    path.write_bytes(_workbook(row))
    run = xlrecalc.load if call == "load" else xlrecalc.recalculate
    with pytest.warns(DeprecationWarning, match=f'recalc_on_load="{mode}"'):
        result = run(path, **alias)  # type: ignore[operator]
    loaded = result if call == "load" else xlrecalc.load(path, recalc_on_load="never")
    got = loaded.value("A2")
    assert got == (6.0 if mode == "always" else 999.0)


def test_the_new_option_does_not_warn(tmp_path: Path) -> None:
    path = tmp_path / "book.xlsx"
    path.write_bytes(_workbook(next(r for r in _rows() if r[0] == "patched")))
    with warnings.catch_warnings():
        warnings.simplefilter("error")
        xlrecalc.load(path, recalc_on_load="always")
        xlrecalc.recalculate(path, recalc_on_load="excel")


def test_an_unknown_mode_is_refused() -> None:
    row = next(r for r in _rows() if r[0] == "patched")
    with pytest.raises(ValueError, match="recalc_on_load"):
        xlrecalc.load(io.BytesIO(_workbook(row)), recalc_on_load="sometimes")  # type: ignore[arg-type]
