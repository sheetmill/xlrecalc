"""What does Excel recalculate on opening a workbook it saved, with volatile functions in it?

    uv run python -m capture.volatile_name_probe

Each variant is its own workbook: openpyxl writes it (``probe!A1`` = 2, ``A2 = A1*3``, ``A3`` the
formula under test, ``A4 = A3+1``, the variant's defined names) and Excel saves a copy, so the
file has Excel's own ``calcId``, ``<fileVersion>``, cell forms and ``ca="1"`` marks. The copy's
cached values of A2, A3 and A4 are then planted wrong (999, 777 and 555, only the ``<v>`` text,
each cell otherwise as Excel wrote it), and it is opened with alerts off and read
(``open_read.applescript``), closed without saving, then saved as a copy whose cached values are
read. A cell showing its sentinel was taken from the file; one showing its right value was
recalculated. The variants: a volatile function in the cell, through a defined name (``OFFSET``,
``INDIRECT``, ``RAND``, ``NOW``, a chain of names, a ``LAMBDA`` name calling ``NOW()``); controls
(a plain name or ``LAMBDA``, an unused volatile name, a function name in a string literal); the
``ca`` mark removed from a volatile cell (as ``open_calc``'s ``volatile`` variant planted its
cell) or added to a plain one; manual calculation, after which the application's mode is set back
to automatic (as ``open_calc_probe`` does).

Workbooks live in Excel's container and are removed after; the record is written to
``capture/build/volatile_name_probe.out`` and the capture to ``tests/captures/volatile_names.tsv``.
"""

from __future__ import annotations

import re
import subprocess
import sys
import zipfile
from pathlib import Path

import openpyxl
from openpyxl.workbook.defined_name import DefinedName

from capture.edition import edition_line
from capture.excel_watch import preflight
from capture.open_calc_probe import calc_attributes, edit_calc, saved_values
from capture.writeback_probe import CONTAINER, osascript

HERE = Path(__file__).resolve().parent
OUT = HERE / "build" / "volatile_name_probe.out"
CAPTURE = HERE.parent / "tests" / "captures" / "volatile_names.tsv"
US = "\x1f"
ROWS = 4
SENTINELS = {"A2": "999", "A3": "777", "A4": "555"}
LAMBDA_NOW = "_xlfn.LAMBDA(_xlpm.x,_xlpm.x+NOW()*0)"
LAMBDA_PLAIN = "_xlfn.LAMBDA(_xlpm.x,_xlpm.x+0)"

# variant: (formula of A3, the defined names {name: text}); every right value of A3 is 6, but
# the literal's (8). A name ending in ``_no_ca`` drops Excel's ``ca="1"`` from A3's formula, as
# ``open_calc``'s ``volatile`` variant wrote its cell; one ending in ``_add_ca`` adds it; one
# starting ``manual`` sets manual calculation (``calcMode="manual"``), and the application's mode
# is set back to automatic after it.
VARIANTS: dict[str, tuple[str, dict[str, str]]] = {
    "plain_name": ("=N*3", {"N": "probe!$A$1"}),
    "offset_in_cell": ("=OFFSET(A1,0,0)*3", {}),
    "offset_name": ("=N*3", {"N": "OFFSET(probe!$A$1,0,0)"}),
    "indirect_name": ("=N*3", {"N": 'INDIRECT("probe!A1")'}),
    "rand_name": ("=N*3", {"N": "RAND()*0+probe!$A$1"}),
    "now_name": ("=N*3", {"N": "NOW()*0+probe!$A$1"}),
    "chained_name": ("=N*3", {"M": "OFFSET(probe!$A$1,0,0)", "N": "M+0"}),
    "unused_offset_name": ("=A1*3", {"N": "OFFSET(probe!$A$1,0,0)"}),
    "lambda_plain_name": ("=F(A1)*3", {"F": LAMBDA_PLAIN}),
    "lambda_now_name": ("=F(A1)*3", {"F": LAMBDA_NOW}),
    "literal": ('=LEN("RAND()")+A1', {}),
    "rand_in_cell": ("=RAND()*0+A1*3", {}),
    "rand_in_cell_no_ca": ("=RAND()*0+A1*3", {}),
    "offset_in_cell_no_ca": ("=OFFSET(A1,0,0)*3", {}),
    "offset_name_no_ca": ("=N*3", {"N": "OFFSET(probe!$A$1,0,0)"}),
    "plain_add_ca": ("=A1*3", {}),
    "manual_rand_in_cell": ("=RAND()*0+A1*3", {}),
    "manual_rand_in_cell_no_ca": ("=RAND()*0+A1*3", {}),
}


def build(path: Path, formula: str, names: dict[str, str]) -> None:
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "probe"
    ws["A1"] = 2
    ws["A2"] = "=A1*3"
    ws["A3"] = formula
    ws["A4"] = "=A3+1"
    for name, text in names.items():
        wb.defined_names[name] = DefinedName(name, attr_text=text)
    wb.save(path)


def _same(text: str) -> str:
    return text


def _manual(text: str) -> str:
    """Manual calculation in workbook.xml's ``<calcPr>``."""
    return re.sub(r"<calcPr\b", '<calcPr calcMode="manual"', text, count=1)


def plant(xml: str, ca: str | None = None) -> str:
    """Each sentinel cell's ``<v>`` text replaced, the rest of the element as Excel wrote it."""
    if ca == "drop":
        xml, n = re.subn(r'(<c r="A3"[^>]*><f)\s+ca="1"', r"\1", xml, count=1)
        if n != 1:
            raise RuntimeError("no ca on A3: cannot drop it")
    elif ca == "add":
        xml, n = re.subn(r'(<c r="A3"[^>]*><f)>', r'\1 ca="1">', xml, count=1)
        if n != 1:
            raise RuntimeError("A3's formula has attributes: cannot add ca")
    for ref, value in SENTINELS.items():
        cell = re.compile(rf'(<c r="{ref}"[^>]*>.*?<v>)(.*?)(</v>)', re.S)
        xml, n = cell.subn(lambda m, v=value: m.group(1) + v + m.group(3), xml, count=1)
        if n != 1:
            raise RuntimeError(f"no cached value in {ref}: cannot plant")
    return xml


def defined_names(path: Path) -> str:
    with zipfile.ZipFile(path) as z:
        workbook = z.read("xl/workbook.xml").decode()
    return " ".join(re.findall(r"<definedName\b[^>]*>.*?</definedName>", workbook)) or "-"


def cell_xml(path: Path, ref: str) -> str:
    with zipfile.ZipFile(path) as z:
        sheet = z.read("xl/worksheets/sheet1.xml").decode()
    m = re.search(rf'<c r="{ref}"[^>]*?(?:/>|>.*?</c>)', sheet, re.S)
    return m.group(0) if m else "-"


def main() -> None:
    preflight()
    CONTAINER.mkdir(parents=True, exist_ok=True)
    OUT.parent.mkdir(exist_ok=True)
    lines: list[str] = []

    def log(text: str) -> None:
        print(text, flush=True)
        lines.append(text)

    for p in CONTAINER.glob("xlrecalc_vn_*.xlsx"):
        p.unlink()
    version = subprocess.run(
        ["osascript", "-e", 'tell application "Microsoft Excel" to version'],
        capture_output=True,
        text=True,
    ).stdout.strip()
    log(f"excel {version}; openpyxl {openpyxl.__version__}")
    rows: list[str] = []
    restore: Path | None = None  # a workbook to set the mode back with, once manual ran
    try:
        for name, (formula, names) in VARIANTS.items():
            src = CONTAINER / f"xlrecalc_vn_{name}_src.xlsx"
            base = CONTAINER / f"xlrecalc_vn_{name}_base.xlsx"
            path = CONTAINER / f"xlrecalc_vn_{name}.xlsx"
            copy = CONTAINER / f"xlrecalc_vn_{name}_saved.xlsx"
            build(src, formula, names)
            log(f"\n{name}: {formula} {names}")
            log("  base: " + osascript("save_copy.applescript", str(src), str(base)))
            log(f"  base calcPr {calc_attributes(base)}; names {defined_names(base)}")
            log(f"  base A3 {cell_xml(base, 'A3')}; file {saved_values(base)}")
            ca = "drop" if name.endswith("_no_ca") else "add" if name.endswith("_add_ca") else None
            manual = name.startswith("manual")
            restore = base if manual else restore
            edit_calc(base, path, lambda x, ca=ca: plant(x, ca), _manual if manual else _same)
            log(f"  planted A3 {cell_xml(path, 'A3')}")
            log(f"  planted: {saved_values(path)}; calcPr {calc_attributes(path)}")
            fields = osascript("open_read.applescript", str(path), str(ROWS)).split(US)
            if fields[0] != "open":
                log(f"  REFUSED {fields[1:]}")
                continue
            shown = [fields[2 + 3 * i] for i in range(ROWS)]  # displayed text of A1:A4
            log(f"  opened: {fields[-2]} {fields[-1]}; shown {shown}")
            osascript("save_copy.applescript", str(path), str(copy))
            log(f"  saved copy: {saved_values(copy)}")
            if manual:  # the application's mode back to automatic (open_calc does the same)
                log("  restore: " + osascript("calc_automatic.applescript", str(base)))
            fields_out = (
                name,
                formula,
                " ".join(f"{k}={v}" for k, v in names.items()) or "-",
                defined_names(base),
                calc_attributes(path),
                " ".join(cell_xml(path, ref) for ref in ("A2", "A3", "A4")),
                saved_values(base),
                " ".join(map(repr, shown[1:])),
                saved_values(copy),
            )
            rows.append("\t".join(fields_out))
    finally:
        if restore is not None:
            log("\nrestore: " + osascript("calc_automatic.applescript", str(restore)))
        OUT.write_text("\n".join(lines) + "\n", encoding="utf-8")
        if len(rows) == len(VARIANTS):
            header = [
                "# suite: volatile_names (volatile functions, ca marks and names, on opening)",
                f"# excel: {version} for Mac, locale en_US",
                edition_line(),
                "# generator: capture/volatile_name_probe.py (a copy Excel saved, A1 = 2, A2 =",
                "#   A1*3 planted 999, A3 the formula under test 777, A4 = A3+1 555)",
                "# columns: variant, A3's formula as openpyxl wrote it, its names, the names",
                "#   as Excel saved them, <fileVersion> and <calcPr> as planted, A2:A4 as",
                "#   planted, the cells Excel saved, displayed text of A2:A4 on opening the",
                "#   planted file, the cells of the copy Excel saved after opening it",
            ]
            CAPTURE.write_text("\n".join([*header, *rows]) + "\n", encoding="utf-8")
        for p in CONTAINER.glob("xlrecalc_vn_*.xlsx"):
            p.unlink()


if __name__ == "__main__":
    sys.exit(main())
