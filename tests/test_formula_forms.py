"""Formulas as Excel writes them into a file round-trip through the parser (ARCHITECTURE.md §5.2).

``tests/captures/formula_forms.tsv`` holds each corpus formula as written and the text Excel
saved for it, captured by ``capture/formula_forms.py``.
"""

import ast
from pathlib import Path

import pytest

from xlrecalc._formula import parse

CAPTURE = Path(__file__).with_name("captures") / "formula_forms.tsv"


def rows() -> list[tuple[str, str | None]]:
    out = []
    for line in CAPTURE.read_text(encoding="utf-8").splitlines():
        if not line.startswith("#"):
            written, excel = line.split("\t")
            out.append((ast.literal_eval(written), ast.literal_eval(excel)))
    return out


ROWS = rows()
SAVED = [excel for _, excel in ROWS if excel not in (None, "REFUSED")]


@pytest.mark.excel
@pytest.mark.parametrize("text", SAVED, ids=repr)
def test_excel_file_form_round_trips(text):
    assert parse(text).text() == text


@pytest.mark.excel
def test_what_excel_changes():
    saved = dict(ROWS)
    # Excel writes number literals in its own form, at most 15 significant digits.
    assert saved[".5"] == "0.5"
    assert saved["1E3"] == "1000"
    assert saved["1.5E-10"] == "0.00000000015"
    assert saved["123456789012345678901234567890"] == "1.23456789012345E+29"
    # A leading unary plus, outer whitespace and the explicit intersection are dropped.
    assert saved["+1"] == "1"
    assert saved[" 1 + 2 "] == "1 + 2"
    assert saved["@A1:A10"] == "A1:A10"
    assert saved["_xlfn.SINGLE(A1:A10)"] == "A1:A10"
    # Line breaks become LF; inner spaces survive except before a comma.
    assert saved["IF(A1>0,\r\n  1,\r\n  2)"] == "IF(A1>0,\n  1,\n  2)"
    assert saved["SUM(A1 , B1)"] == "SUM(A1, B1)"
    # The spill operator has no file form (it is _xlfn.ANCHORARRAY), nor do the newer errors.
    for text in ("A1#", "Sheet2!A1#", "#SPILL!", "#CALC!", "-\t1", "SUM( )"):
        assert saved[text] == "REFUSED", text
    assert saved["_xlfn.ANCHORARRAY(A1)"] == "_xlfn.ANCHORARRAY(A1)"
    # Number literals keep their first 15 significant digits; unneeded quotes go; case is upper.
    assert saved["0.123456789012345678"] == "0.123456789012345"
    assert saved["1234567890123456789.5"] == "1234567890123450000"
    assert saved["'Résumé'!A1"] == "Résumé!A1"
    assert saved["sum(a1:a3)"] == "SUM(A1:A3)"
    # Excel's limits are accepted as written.
    for text in saved:
        if text.startswith("IF(A1>0," * 64) or len(text) > 7000 or text.count(",") == 254:
            assert saved[text] == text
