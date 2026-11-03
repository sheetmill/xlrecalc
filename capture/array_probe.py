"""Arrays, spill and implicit intersection, as Excel computes them (ARCHITECTURE.md §7.7).

    uv run python -m capture.array_probe

Non-interactive. One sheet holds a data block (Z1:Z5 = 1..5, AA1:AA5 = 10..50, AB1:AB5 = a..e)
and three kinds of formula over it:

- legacy formulas (as openpyxl writes them: no dynamic-array marker), each in column A at row 3,
  inside the data block's rows, and again at row 20 or below, outside them;
- legacy array formulas (``t="array"``, entered with Ctrl+Shift+Enter in Excel), written by
  openpyxl's ``ArrayFormula`` over a declared range;
- dynamic-array formulas: ``<c cm="1"><f t="array" ref="anchor">``, with the dynamic-array
  metadata part, as Excel writes a formula entered in Excel 365, so Excel spills them; one has
  a constant in its way.

Excel opens the workbook (``fullCalcOnLoad``), calculates and saves a copy. Every cell of the
copy is recorded with its value, formula and the formula's ``t`` and ``ref``, and its ``cm``;
the copy's metadata part and rich-value parts are kept too. Writes ``tests/captures/arrays.tsv``,
``tests/captures/arrays_metadata.xml`` and ``tests/captures/arrays_rd*.xml``.
"""

from __future__ import annotations

import datetime as dt
import re
import subprocess
import sys
import zipfile
from pathlib import Path

import openpyxl
from openpyxl.worksheet.formula import ArrayFormula

from capture.pipeline import CAPTURES, CONTAINER, excel_version
from xlrecalc._xlsx import Book
from xlrecalc._xlsx._write import number_text, text_escape

HERE = Path(__file__).resolve().parent
NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
METADATA = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
    f'<metadata xmlns="{NS}" '
    'xmlns:xda="http://schemas.microsoft.com/office/spreadsheetml/2017/dynamicarray">'
    '<metadataTypes count="1"><metadataType name="XLDAPR" minSupportedVersion="120000" copy="1" '
    'pasteAll="1" pasteValues="1" merge="1" splitFirst="1" rowColShift="1" clearFormats="1" '
    'clearComments="1" assign="1" coerce="1" cellMeta="1"/></metadataTypes>'
    '<futureMetadata name="XLDAPR" count="1"><bk><extLst>'
    '<ext uri="{bdbb8cdc-fa1e-496e-a857-3c3f30c029c3}">'
    '<xda:dynamicArrayProperties fDynamic="1" fCollapsed="0"/></ext></extLst></bk>'
    '</futureMetadata><cellMetadata count="1"><bk><rc t="1" v="0"/></bk></cellMetadata>'
    "</metadata>"
)

LEGACY = [
    "Z1:Z5",
    "Z1:Z5*2",
    "-Z1:Z5",
    'Z1:Z5&"x"',
    "Z1:Z5>2",
    "Z1:Z5+AA1:AA5",
    "Z:Z",
    "Z1:AA1",
    "ABS(Z1:Z5)",
    "ABS(-Z1:Z5)",
    "N(Z1:Z5)",
    "ISNUMBER(Z1:Z5)",
    "SUM(Z1:Z5)",
    "SUM(Z1:Z5*2)",
    "SUM(Z1:Z5>2)",
    "SUM(--(Z1:Z5>2))",
    "SUM(IF(Z1:Z5>2,1,0))",
    "SUM(Z1:Z5*{1;2;3;4;5})",
    "MAX(Z1:Z5*2)",
    "AVERAGE(Z1:Z5*2)",
    "COUNT(Z1:Z5*2)",
    "SUMPRODUCT(Z1:Z5*2)",
    "SUMPRODUCT(--(Z1:Z5>2))",
    "SUMPRODUCT(Z1:Z5,AA1:AA5)",
    "SUMPRODUCT((Z1:Z5>2)*AA1:AA5)",
    "INDEX(Z1:Z5*2,2)",
    "INDEX(Z1:AA5,0,1)",
    "SUM(INDEX(Z1:AA5,0,1))",
    "MATCH(6,Z1:Z5*2,0)",
    "LOOKUP(6,Z1:Z5*2)",
    "VLOOKUP(6,Z1:Z5*2,1,FALSE)",
    "ROWS(Z1:Z5*2)",
    "AND(Z1:Z5>0)",
    "OR(Z1:Z5>4)",
    'COUNTIF(Z1:Z5,">2")',
    'SUMIF(Z1:Z5,">2",AA1:AA5)',
    'IF(Z1:Z5>2,"big","small")',
    "CONCATENATE(AB1:AB5)",
    "{1,2,3}",
    "{1,2,3}*2",
    "SUM({1,2,3}*2)",
    "ABS({-1,-2})",
    "SUM(ABS({-1,-2}))",
    "{1,2,3}+{1;2}",
    "INDEX({1,2;3,4},2,1)",
]
# (formula, declared range)
CSE = [
    ("Z1:Z5*2", "C1:C5"),
    ("Z1:Z5*2", "C7:C9"),
    ("Z1:Z5*2", "C11:C17"),
    ("Z1:AA5", "E1:F5"),
    ("Z1:Z5", "E7:F11"),
    ("SUM(Z1:Z5*AA1:AA5)", "C19:C19"),
    ("{1,2,3}", "H1:J1"),
    ("{1,2,3}", "H3:H5"),
    ("IF(Z1:Z5>2,Z1:Z5,0)", "C21:C25"),
    ("Z1:Z5+{10,20}", "E13:F17"),
    ("{1,2,3}+{1;2}", "H7:J8"),
    ("{1,2,3}+{1,2}", "H10:J10"),
    ("ABS(-Z1:Z5)", "C27:C31"),
    ('Z1:Z5&"x"', "E19:E23"),
    ("MAX(IF(Z1:Z5>2,AA1:AA5))", "C33:C33"),
    ("INDEX(Z1:Z5*2,2)", "C35:C35"),
    ("MATCH(6,Z1:Z5*2,0)", "C37:C37"),
    ("Z1:Z5", "H12:H12"),
    ("Z3:Z5*2", "H14:H18"),
]
# (anchor, formula in file form); a constant sits at S3, in the way of S1's spill
DYNAMIC = [
    ("L1", "Z1:Z5*2"),
    ("N1", "Z1:AA5"),
    ("Q1", "{1,2;3,4}"),
    ("L10", "Z1:Z5+{10,20}"),
    ("S1", "Z1:Z5*2"),
    ("U3", "_xlfn.SINGLE(Z1:Z5)"),
    ("U5", "Z3"),
    ("U7", "SUM(Z1:Z5*2)"),
    ("W1", 'IF(Z1:Z5>2,"big","small")'),
    ("X1", "_xlfn.ANCHORARRAY(L1)*2"),
    ("U9", "SUM(_xlfn.ANCHORARRAY(L1))"),
    ("U11", "ROWS(_xlfn.ANCHORARRAY(N1))"),
    ("U13", "MATCH(6,Z1:Z5*2,0)"),
    ("U15", "SUMPRODUCT(--(Z1:Z5>2))"),
    ("L17", "ABS(-Z1:Z5)"),
    ("N17", "Z1:Z5>2"),
    ("Q10", "{1,2,3}+{1;2}"),
    ("Q13", "{1,2,3}+{1,2}"),
    ("U17", "Z1:Z5"),
]
# Plain (legacy) formulas reading spilled cells
READERS = [("V19", "L2+1"), ("V20", "SUM(L1:L5)"), ("V21", "S2")]


def layout() -> tuple[list[tuple[str, str, str]], list[tuple[str, str, str]]]:
    """(kind, cell or range, formula) for openpyxl, and the dynamic-array cells patched in."""
    cells: list[tuple[str, str, str]] = []
    for k, formula in enumerate(LEGACY):
        cells.append(("legacy", f"A{20 + k}", formula))
        cells.append(("legacy", f"AD{20 + k}", formula))  # a column outside the data's
    for k, formula in enumerate(LEGACY):  # at row 3 too, inside the data's rows, one per column
        cells.append(("legacy_row3", f"{_col(31 + k)}3", formula))
    for formula, ref in CSE:
        cells.append(("cse", ref, formula))
    for cell, formula in READERS:
        cells.append(("reader", cell, formula))
    dynamic = [("dynamic", anchor, formula) for anchor, formula in DYNAMIC]
    return cells, dynamic


def _col(n: int) -> str:
    s = ""
    while n:
        n, r = divmod(n - 1, 26)
        s = chr(65 + r) + s
    return s


def build(path: Path) -> None:
    cells, dynamic = layout()
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "probe"
    for r in range(1, 6):
        ws.cell(r, 26, float(r))
        ws.cell(r, 27, float(10 * r))
        ws.cell(r, 28, "abcde"[r - 1])
    ws["S3"] = "in the way"
    for kind, where, formula in cells:
        if kind == "cse":
            ws[where.split(":")[0]] = ArrayFormula(where, "=" + formula)
        else:
            ws[where] = "=" + formula
    for _, anchor, _ in dynamic:
        ws[anchor] = "=0"  # a placeholder cell, replaced below
    wb.save(path)
    with zipfile.ZipFile(path) as z:
        parts = {name: z.read(name) for name in z.namelist()}
    sheet = parts["xl/worksheets/sheet1.xml"].decode()
    for _, anchor, formula in dynamic:
        new = f'<c r="{anchor}" cm="1"><f t="array" ref="{anchor}">{text_escape(formula)}</f></c>'
        sheet, n = re.subn(rf'<c r="{anchor}"[^>]*>.*?</c>', new, sheet, count=1, flags=re.S)
        assert n == 1, anchor
    parts["xl/worksheets/sheet1.xml"] = sheet.encode()
    parts["xl/metadata.xml"] = METADATA.encode()
    rels = (
        parts["xl/_rels/workbook.xml.rels"]
        .decode()
        .replace(
            "</Relationships>",
            '<Relationship Id="rIdMeta" Type="http://schemas.openxmlformats.org/officeDocument/2006/'
            'relationships/sheetMetadata" Target="metadata.xml"/></Relationships>',
        )
    )
    parts["xl/_rels/workbook.xml.rels"] = rels.encode()
    types = (
        parts["[Content_Types].xml"]
        .decode()
        .replace(
            "</Types>",
            '<Override PartName="/xl/metadata.xml" ContentType="application/'
            'vnd.openxmlformats-officedocument.spreadsheetml.sheetMetadata+xml"/></Types>',
        )
    )
    parts["[Content_Types].xml"] = types.encode()
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
        for name, data in parts.items():
            z.writestr(name, data)


def main() -> None:
    CONTAINER.mkdir(parents=True, exist_ok=True)
    src = CONTAINER / "xlrecalc_arrays.xlsx"
    copy = CONTAINER / "xlrecalc_arrays_saved.xlsx"
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
    cells, dynamic = layout()
    header = [
        "# suite: arrays",
        f"# excel: {version} for Mac, locale en_US",
        f"# captured: {dt.date.today().isoformat()}",
        "# generator: capture/array_probe.py",
        "# method: legacy formulas, legacy array formulas (t=array) and dynamic-array formulas",
        "#   (cm=1 with the XLDAPR metadata) over Z1:AB5; Excel calculates and saves a copy;",
        "#   every cell of the copy is recorded",
        "# input rows: input, kind, cell or range, formula (file form)",
        "# cell rows: cell, address, type, value (repr), formula (repr), f_t, f_ref, cm, vm",
    ]
    body = [
        "\t".join(["input", kind, where, repr(formula)]) for kind, where, formula in cells + dynamic
    ]
    with Book(copy) as book:
        for c in book.cells("probe"):
            f = c.formula
            value = number_text(c.value) if isinstance(c.value, float) else c.value
            body.append(
                "\t".join(
                    [
                        "cell",
                        f"{_col(c.column)}{c.row}",
                        c.type,
                        repr(value),
                        repr(f.text if f else None),
                        f.kind if f else "",
                        (f.ref or "") if f else "",
                        c.cm or "",
                        c.vm or "",
                    ]
                )
            )
    (CAPTURES / "arrays.tsv").write_text("\n".join([*header, *body]) + "\n", encoding="utf-8")
    with zipfile.ZipFile(copy) as z:
        names = z.namelist()
        meta = next((n for n in names if n.endswith("metadata.xml")), None)
        if meta:
            (CAPTURES / "arrays_metadata.xml").write_bytes(z.read(meta))
        for name in names:  # the rich values behind errors such as #SPILL!
            if name.startswith("xl/richData/") and name.endswith(".xml"):
                (CAPTURES / f"arrays_{Path(name).name}").write_bytes(z.read(name))
        print("parts:", [n for n in names if "metadata" in n or "rich" in n.lower()])
    print(f"{len(body)} rows -> arrays.tsv")
    for p in (src, copy):
        p.unlink(missing_ok=True)


if __name__ == "__main__":
    main()
