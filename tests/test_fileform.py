"""The file-form printer and its inverse (ARCHITECTURE.md §5.2).

Captures ``formula_bar`` (each corpus formula typed in the formula bar, and Excel's file form of
it) and ``formula_forms`` (each corpus formula written into a file, and the text Excel saved
back): the printer must give Excel's text except for the differences listed with a reason.
"""

from __future__ import annotations

import ast
from pathlib import Path

import pytest
from tools.prefix_table import prefix_table

from xlrecalc._formula._fileform import display_form, file_form
from xlrecalc._formula._prefixes import PREFIXES

CAPTURES = Path(__file__).with_name("captures")


def rows(name: str) -> list[list[str]]:
    text = (CAPTURES / f"{name}.tsv").read_text(encoding="utf-8")
    return [line.split("\t") for line in text.splitlines() if not line.startswith("#")]


# Typed text -> why the printer's form differs from what the formula-bar capture recorded.
BAR_DIFFERENCES = {
    "=LAMBDA(x,x+1)(2)": "AppleScript's legacy formula does not know LAMBDA and stores the text "
    "as typed; a file holding _xlfn.LAMBDA(_xlpm.x,...) opens and is kept (formula_forms)",
    "=LAMBDA(x,y,x*y)(2,3)": "as above",
    "=eval(1)": "a function the package does not know keeps the case it was typed in",
    "=Table1[[Col'[1']]]": "Excel simplifies a single-column structured reference; both open",
}
# Written text -> why file_form(display_form(written)) differs from what Excel saved.
FORM_DIFFERENCES = {
    " 1": "Excel drops leading whitespace when it reads a file; the formula bar keeps it",
    " 1 + 2 ": "as above",
    "\n\nA1": "as above",
    "      6:6": "as above",
    "Table1[@Price]": "outside its table Excel makes it #REF!; the printer writes its file form",
    "[@Price]": "as above",
    "Table1[['#Hash]]": "as above",
    "Table1[[Col'[1']]]": "Excel simplifies a single-column structured reference",
    "_xlfn.EVAL(1)": "Excel's formula bar writes EVAL unprefixed (formula_bar), so its display "
    "form loses the prefix",
}


def test_prefix_table_is_derived_from_the_captures():
    assert prefix_table() == PREFIXES, "rerun: uv run python tools/prefix_table.py"


def test_prefixes():
    assert PREFIXES["IFNA"] == "_xlfn."
    assert PREFIXES["SORT"] == PREFIXES["FILTER"] == "_xlfn._xlws."
    assert PREFIXES["LET"] == PREFIXES["LAMBDA"] == PREFIXES["MAXIFS"] == "_xlfn."
    assert "SUM" not in PREFIXES
    assert "SUMIFS" not in PREFIXES  # Excel 2007: no prefix (capture prefixes)


@pytest.mark.excel
@pytest.mark.parametrize(
    "row", [r for r in rows("formula_bar") if r[0] != "function"], ids=lambda r: r[1][:40]
)
def test_printer_matches_the_formula_bar(row):
    _tag, typed, saved = row
    typed, saved = ast.literal_eval(typed), ast.literal_eval(saved)
    if saved == "REJECTED":
        pytest.skip("the formula bar does not take it")
    if len(typed) > 256:
        pytest.skip("AppleScript sets at most 255 characters of a formula")
    if typed in BAR_DIFFERENCES:
        assert file_form(typed) != saved
        return
    assert file_form(typed) == saved


@pytest.mark.excel
@pytest.mark.parametrize("row", rows("formula_forms"), ids=lambda r: r[0][:40])
def test_printer_gives_what_excel_saves_from_a_file(row):
    written, saved = map(ast.literal_eval, row)
    if saved == "REFUSED":
        return
    printed = file_form("=" + display_form(written))
    if written in FORM_DIFFERENCES:
        assert printed != saved
        return
    assert printed == saved


@pytest.mark.parametrize(
    ("typed", "printed"),
    [
        ("=A1#", "_xlfn.ANCHORARRAY(A1)"),
        ("=Sheet2!A1#", "_xlfn.ANCHORARRAY(Sheet2!A1)"),
        ("=SORT(A1:A3)#", "_xlfn.ANCHORARRAY(_xlfn._xlws.SORT(A1:A3))"),
        ("=@A1:A3", "A1:A3"),
        ("=IFNA(1,2)", "_xlfn.IFNA(1,2)"),
        ("=_xlfn.IFNA(1,2)", "_xlfn.IFNA(1,2)"),
        ("=LET(x, 1, x+1)", "_xlfn.LET(_xlpm.x, 1, _xlpm.x+1)"),
        ("=.5+1e3", "0.5+1000"),
        ("=0.123456789012345678", "0.123456789012345"),
        ("=sum(a1:b2)", "SUM(A1:B2)"),
        ('="a\tb"&true', '"a\tb"&TRUE'),
        ("=+1", "1"),
        ("={1 , +2 ; 3 , 4}", "{1,2;3,4}"),
        ("=SUM(A1 , B1)", "SUM(A1, B1)"),
        ("=IF(A1,\r\n1,\r2)", "IF(A1,\n1,\n2)"),
        ("=1 + 2  ", "1 + 2"),
        ("='Résumé'!A1", "Résumé!A1"),
        ("='v1.2'!A1", "'v1.2'!A1"),
        ("='My Sheet'!A1", "'My Sheet'!A1"),
        ("=Table1[@Price]", "Table1[[#This Row],[Price]]"),
        ("=Table1[@[Unit Price]]", "Table1[[#This Row],[Unit Price]]"),
        ("=_xll.ADDIN(1)", "_xll.ADDIN(1)"),
        ("=Sheet2!MyFunc(1)", "Sheet2!MyFunc(1)"),
        ('="_x0041_"', '"_x0041_"'),
    ],
)
def test_file_form(typed, printed):
    assert file_form(typed) == printed


@pytest.mark.parametrize(
    ("written", "shown"),
    [
        ("_xlfn.ANCHORARRAY(A1)", "A1#"),
        ("_xlfn.ANCHORARRAY(Sheet2!A1)", "Sheet2!A1#"),
        ("_xlfn.SINGLE(A1:A10)", "@A1:A10"),
        ("_xlfn.SINGLE(A1+1)", "@(A1+1)"),
        ("_xlfn.LET(_xlpm.x,1,_xlpm.x+1)", "LET(x,1,x+1)"),
        ("_xlfn._xlws.SORT(A1:A3)", "SORT(A1:A3)"),
        ("SUM(A1)", "SUM(A1)"),
    ],
)
def test_display_form(written, shown):
    assert display_form(written) == shown
    assert file_form("=" + shown) == written.replace("_xlfn.SINGLE(A1:A10)", "A1:A10").replace(
        "_xlfn.SINGLE(A1+1)", "(A1+1)"
    )
