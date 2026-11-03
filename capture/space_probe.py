"""What Excel keeps of whitespace in a ``<t>`` without ``xml:space="preserve"`` (§9.1, §9.2).

    uv run python -m capture.space_probe

Non-interactive. openpyxl writes the workbook's skeleton; its sheet is then replaced by one
whose column A holds each case as an inline string (``<is><t>``) and column B the same text as
a shared string (``<si><t>``), both written without ``xml:space``, and whose columns C and D
hold ``=LEN(A<row>)`` and ``=LEN(B<row>)``. Excel opens it (alerts off: a file it would
repair fails with -50), calculates and saves a copy; the copy's text and lengths are what Excel
read. Writes ``tests/captures/whitespace.tsv``.
"""

from __future__ import annotations

import datetime as dt
import re
import subprocess
import sys
import zipfile
from pathlib import Path

import openpyxl

from capture.pipeline import CAPTURES, CONTAINER, excel_version
from xlrecalc._xlsx import Book

HERE = Path(__file__).resolve().parent
MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"

# (tag, the <t> content or <r> runs as written, XML)
CASES: list[tuple[str, str]] = [
    ("leading", "<t> lead</t>"),
    ("trailing", "<t>trail </t>"),
    ("spaces", "<t>  </t>"),
    ("space", "<t> </t>"),
    ("inner_newline", "<t>a\nb</t>"),
    ("leading_newline", "<t>\nlead</t>"),
    ("trailing_newline", "<t>trail\n</t>"),
    ("leading_tab", "<t>\tlead</t>"),
    ("inner_spaces", "<t>mid  dle</t>"),
    ("empty", "<t/>"),
    ("runs", "<r><t>a </t></r><r><t> b</t></r>"),
    ("runs_spaces", "<r><t> </t></r><r><t>b</t></r>"),
    ("preserved", '<t xml:space="preserve"> kept </t>'),
    ("char_ref", "<t>&#32;ref&#32;</t>"),
    ("escaped_cr", "<t>_x000D_cr</t>"),
]


def build(path: Path) -> None:
    wb = openpyxl.Workbook()
    wb.active.title = "probe"
    wb.active["A1"] = "placeholder"
    wb.save(path)
    with zipfile.ZipFile(path) as z:
        parts = {name: z.read(name) for name in z.namelist()}
    rows = []
    for r, (_, xml) in enumerate(CASES, 1):
        rows.append(
            f'<row r="{r}"><c r="A{r}" t="inlineStr"><is>{xml}</is></c>'
            f'<c r="B{r}" t="s"><v>{r - 1}</v></c>'
            f'<c r="C{r}"><f>LEN(A{r})</f></c><c r="D{r}"><f>LEN(B{r})</f></c></row>'
        )
    sheet = parts["xl/worksheets/sheet1.xml"].decode()
    sheet = re.sub(r"<sheetData>.*</sheetData>", f"<sheetData>{''.join(rows)}</sheetData>", sheet)
    parts["xl/worksheets/sheet1.xml"] = sheet.encode()
    items = "".join(f"<si>{xml}</si>" for _, xml in CASES)
    parts["xl/sharedStrings.xml"] = (
        f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<sst xmlns="{MAIN}" '
        f'count="{len(CASES)}" uniqueCount="{len(CASES)}">{items}</sst>'
    ).encode()
    rels = (
        parts["xl/_rels/workbook.xml.rels"]
        .decode()
        .replace(
            "</Relationships>",
            '<Relationship Id="rIdSst" Type="http://schemas.openxmlformats.org/officeDocument/2006/'
            'relationships/sharedStrings" Target="sharedStrings.xml"/></Relationships>',
        )
    )
    parts["xl/_rels/workbook.xml.rels"] = rels.encode()
    types = (
        parts["[Content_Types].xml"]
        .decode()
        .replace(
            "</Types>",
            '<Override PartName="/xl/sharedStrings.xml" ContentType="application/'
            'vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/></Types>',
        )
    )
    parts["[Content_Types].xml"] = types.encode()
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
        for name, data in parts.items():
            z.writestr(name, data)


def main() -> None:
    CONTAINER.mkdir(parents=True, exist_ok=True)
    src = CONTAINER / "xlrecalc_space.xlsx"
    copy = CONTAINER / "xlrecalc_space_saved.xlsx"
    for p in (src, copy):
        p.unlink(missing_ok=True)
    build(src)
    version = excel_version()
    result = subprocess.run(
        ["osascript", str(HERE / "save_copy.applescript"), str(src), str(copy)],
        capture_output=True,
        text=True,
    )
    if result.returncode:
        sys.exit(f"Excel did not open the workbook: {result.stderr.strip()}")
    with Book(copy) as book:
        got = {(c.row, c.column): c.value for c in book.cells("probe")}
    header = [
        "# suite: whitespace",
        f"# excel: {version} for Mac, locale en_US",
        f"# captured: {dt.date.today().isoformat()}",
        "# generator: capture/space_probe.py",
        "# method: each case written without xml:space as an inline string (column A) and a",
        "#   shared string (column B), LEN of each in C and D; Excel opens it, calculates and",
        "#   saves a copy; text and lengths are read from the copy",
        "# columns: row tag written inline_text inline_len shared_text shared_len (Python repr)",
    ]
    body = []
    for r, (tag, xml) in enumerate(CASES, 1):
        a, b, c, d = (got.get((r, k)) for k in (1, 2, 3, 4))
        body.append("\t".join([str(r), tag, repr(xml), repr(a), repr(c), repr(b), repr(d)]))
        print(f"{r:3} {tag:17} {xml!r:38} {a!r:14} {c!r:6} {b!r:14} {d!r}")
    (CAPTURES / "whitespace.tsv").write_text("\n".join([*header, *body]) + "\n", encoding="utf-8")
    for p in (src, copy):
        p.unlink(missing_ok=True)


if __name__ == "__main__":
    main()
