"""Number forms Excel never writes, and rows and columns off its grid, read as Excel reads them.

Capture ``value_forms`` (``capture/value_probe.py``): each case is the workbook the probe built,
with A1 rewritten; Excel either refused it (error -50) or opened it, and the TSV holds A1's and
B1's (``=A1*2``) values from the copy it saved. Here a refused case must be a ``PackageError``,
and an opened one must give Excel's A1 and B1.
"""

from __future__ import annotations

import io
from pathlib import Path

import pytest

import xlrecalc
from xlrecalc import ErrorValue, PackageError

openpyxl = pytest.importorskip("openpyxl")
from capture.value_probe import cases  # noqa: E402

CAPTURE = Path(__file__).with_name("captures") / "value_forms.tsv"
ROWS = [
    line.split("\t")
    for line in CAPTURE.read_text(encoding="utf-8").splitlines()[4:]
    if line.strip()
]
BUILT = cases()


def expected(text: str) -> object:
    kind, _, value = text.partition(":")
    if text == "absent":
        return None
    if kind == "e":
        return ErrorValue(value)
    return float(value)


@pytest.mark.parametrize("row", ROWS, ids=[r[0] for r in ROWS])
def test_read_as_excel_reads_it(row):
    case, outcome = row[0], row[1]
    data = BUILT[case]
    if outcome == "refused":
        if case == "row 99999999999":
            pytest.fail("Excel opened it")
        with pytest.raises(PackageError):
            xlrecalc.load(io.BytesIO(data))
        return
    a1, b1 = row[3], row[4]
    if case == "row 99999999999":
        # Excel opens it and drops the cell; the reader refuses a row off the sheet (QUIRKS.md).
        with pytest.raises(PackageError, match="off the sheet"):
            xlrecalc.load(io.BytesIO(data))
        return
    book = xlrecalc.load(io.BytesIO(data))
    book.recalculate(full=True)
    assert book.value("A1") == expected(a1)
    assert book.value("B1") == expected(b1)
