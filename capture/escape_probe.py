"""Probe how Excel decodes OOXML's ``_xHHHH_`` escapes in text, and how it writes them.

    uv run python -m capture.escape_probe

Text cells hold raw escapes written into the XML (openpyxl's own text is replaced after it
writes). Excel computes ``LEN``, the code of the second character, and the text through ``&``;
Excel saves a copy, read with the package's reader and as raw XML. Prints the results.
"""

from __future__ import annotations

import re
import subprocess
import zipfile

import openpyxl

from capture.pipeline import CONTAINER, HERE
from xlrecalc._xlsx import Book

RAW = ["a_x000D_b", "_x005F_x0041_", "x_x0041_y", "a_x000d_b", "a_x000Db", "a_x00410_b"]


def main() -> None:
    CONTAINER.mkdir(parents=True, exist_ok=True)
    src = CONTAINER / "xlrecalc_escape.xlsx"
    copy = CONTAINER / "xlrecalc_escape_saved.xlsx"
    for p in (src, copy):
        p.unlink(missing_ok=True)
    wb = openpyxl.Workbook()
    ws = wb.active
    for k in range(len(RAW)):
        r = k + 1
        ws.cell(r, 1, f"MARKER{k}")
        ws.cell(r, 2, f"=LEN(A{r})")
        ws.cell(r, 3, f"=CODE(MID(A{r},2,1))")
        ws.cell(r, 4, f'=A{r}&""')
    n = len(RAW) + 1
    ws.cell(n, 4, '=CHAR(13)&"x"')
    ws.cell(n + 1, 4, '="_x0041_"')
    wb.save(src)
    with zipfile.ZipFile(src) as z:
        parts = {name: z.read(name) for name in z.namelist()}
    sheet = parts["xl/worksheets/sheet1.xml"].decode()
    for k, raw in enumerate(RAW):
        sheet = sheet.replace(f">MARKER{k}<", f">{raw}<")
    parts["xl/worksheets/sheet1.xml"] = sheet.encode()
    with zipfile.ZipFile(src, "w", zipfile.ZIP_DEFLATED) as z:
        for name, data in parts.items():
            z.writestr(name, data)
    result = subprocess.run(
        ["osascript", str(HERE / "save_copy.applescript"), str(src), str(copy)],
        capture_output=True,
        text=True,
    )
    if result.returncode:
        raise SystemExit(result.stderr)
    with Book(copy) as book:
        cells = {(c.row, c.column): c for c in book.cells(book.sheets[0])}
    with zipfile.ZipFile(copy) as z:
        names = z.namelist()
        sst = z.read("xl/sharedStrings.xml").decode() if "xl/sharedStrings.xml" in names else ""
        saved_sheet = z.read("xl/worksheets/sheet1.xml").decode()
    print('raw written | LEN | CODE(2nd) | A&"" as the reader decodes it')
    for k, raw in enumerate(RAW):
        r = k + 1
        row = [cells.get((r, c)) for c in (2, 3, 4)]
        print(f"{raw!r:18} {row[0].value!r:6} {row[1].value!r:8} {row[2].value!r}")
    for r in (n, n + 1):
        c = cells.get((r, 4))
        print(
            f"D{r}: {c.formula.text if c and c.formula else None!r} -> {c.value if c else None!r}"
        )
    print("shared strings Excel wrote:", re.findall(r"<t[^>]*>([^<]*)</t>", sst))
    print(
        "str results Excel wrote:",
        re.findall(r'<c r="D\d+"[^>]*t="str"[^>]*>.*?<v[^>]*>([^<]*)</v>', saved_sheet),
    )
    src.unlink()
    copy.unlink()


if __name__ == "__main__":
    main()
