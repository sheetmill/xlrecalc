"""Whitespace in text, as Excel reads it (capture ``whitespace``, ``capture/space_probe.py``).

Without ``xml:space="preserve"`` Excel drops the leading and trailing whitespace of each ``<t>``
(each run of a rich text separately), for inline and shared strings alike; the reader must
read every captured case as Excel did.
"""

import ast
from pathlib import Path

import pytest

from tests.test_xlsx import cells, package

CAPTURE = Path(__file__).with_name("captures") / "whitespace.tsv"
ROWS = [
    line.split("\t")
    for line in CAPTURE.read_text(encoding="utf-8").splitlines()
    if not line.startswith("#")
]


@pytest.mark.excel
@pytest.mark.parametrize("row", ROWS, ids=lambda r: r[1])
def test_reader_reads_text_as_excel_does(row):
    _number, _tag, written, inline_text, inline_len, shared_text, shared_len = row
    xml = ast.literal_eval(written)
    data = package(
        f'<row r="1"><c r="A1" t="inlineStr"><is>{xml}</is></c><c r="B1" t="s"><v>0</v></c></row>',
        strings=[f"<si>{xml}</si>"],
    )
    got = {c.column: c.value for c in cells(data)}
    assert got[1] == ast.literal_eval(inline_text)
    assert got[2] == ast.literal_eval(shared_text)
    assert len(got[1]) == ast.literal_eval(inline_len)
    assert len(got[2]) == ast.literal_eval(shared_len)
