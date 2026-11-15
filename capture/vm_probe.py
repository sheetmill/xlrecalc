"""What Excel shows for a cell whose new cached value keeps the old value's ``vm``.

    uv run python -m capture.vm_probe

Non-interactive. Excel caches a newer error (``#SPILL!``, ``#CALC!``) as ``#VALUE!`` with a rich
value behind the cell's ``vm`` attribute. Column A holds formulas that Excel computes as such
errors; Excel opens the source and saves a copy (``save_copy.applescript``), so the copy holds
the rich values. The package then edits the inputs so each formula gives another value (a number,
a plain ``#VALUE!``, another newer error) and saves without asking Excel to recalculate on opening:
once as the writer writes it (``written``: a cell given a new value loses its ``vm``) and once with
each cell's ``vm`` from Excel's copy put back (``stale_vm``). Excel opens each with alerts off
(``open_read.applescript``: a refused file fails with -50) and saves a copy.
Writes ``capture/build/vm_probe.out``: each file's cells as XML, what Excel shows, and its copy.
"""

from __future__ import annotations

import re
import zipfile
from pathlib import Path

import openpyxl

import xlrecalc
from capture.excel_watch import preflight
from capture.writeback_probe import CONTAINER, osascript

HERE = Path(__file__).resolve().parent
OUT = HERE / "build" / "vm_probe.out"
US = "\x1f"
ROWS = 9
FORMULAS = {  # row: formula; Z1:Z2 = 1, 2, Z5 = 5, Z6 blank, blockers in B1, B3 and B4
    1: "=TRANSPOSE(Z1:Z2)",  # #SPILL!, then 1 once B1 is cleared
    2: "=FILTER(Z1:Z2,Z1:Z2>Z5)",  # #CALC!, then 2 once Z5 is 1.5
    3: '=IF(Z6=1,"a"+1,TRANSPOSE(Z1:Z2))',  # #SPILL!, then a plain #VALUE! once Z6 is 1
    4: "=IF(Z6=1,FILTER(Z1:Z2,Z1:Z2>9),TRANSPOSE(Z1:Z2))",  # #SPILL!, then #CALC!
    5: "=Z1+Z2",  # 3 throughout: the control
}
LEGACY = {  # legacy formulas (openpyxl's), caching A2's #CALC! with a rich value
    6: '=IF(Z6=1,"a"+1,A2)',  # then a plain #VALUE!
    7: "=IF(Z6=1,A2*0+5,A2)",  # then 5
    8: "=IF(Z6=1,1/0,A2)",  # then #DIV/0!
    9: "=IF(Z6=1,TRUE,A2)",  # then TRUE
}
_CELL = re.compile(r'<c r="A(\d)"[^>]*?(?:/>|>.*?</c>)', re.S)


def cells(path: Path) -> str:
    """Column A's cell elements in the file, one a line."""
    with zipfile.ZipFile(path) as z:
        sheet = z.read("xl/worksheets/sheet1.xml").decode()
    return "\n    ".join(m[0] for m in _CELL.finditer(sheet))


def stale_vm(excel: Path, src: Path, dst: Path) -> None:
    """``src`` with each column-A cell given the ``vm`` it has in ``excel``, if it lost it."""
    with zipfile.ZipFile(excel) as z:
        old = z.read("xl/worksheets/sheet1.xml").decode()
    vms = {m[1]: m[2] for m in re.finditer(r'<c r="(A\d)"[^>]*? vm="(\d+)"', old)}
    with zipfile.ZipFile(src) as z:
        parts = [(i, z.read(i.filename)) for i in z.infolist()]

    def put_back(m: re.Match[str]) -> str:
        tag = m[0]
        if m[1] in vms and " vm=" not in tag:
            tag = tag[:-1] + f' vm="{vms[m[1]]}">'
        return tag

    with zipfile.ZipFile(dst, "w", zipfile.ZIP_DEFLATED) as z:
        for info, data in parts:
            if info.filename == "xl/worksheets/sheet1.xml":
                data = re.sub(r'<c r="(A\d)"[^>/]*>', put_back, data.decode()).encode()
            z.writestr(info.filename, data)


def main() -> None:
    preflight()
    CONTAINER.mkdir(parents=True, exist_ok=True)
    OUT.parent.mkdir(parents=True, exist_ok=True)
    names = ("src", "excel", "written", "stale_vm")
    paths = {n: CONTAINER / f"xlrecalc_vm_{n}.xlsx" for n in names}
    for p in CONTAINER.glob("xlrecalc_vm_*.xlsx"):
        p.unlink()
    lines: list[str] = []

    def log(text: str) -> None:
        print(text)
        lines.append(text)

    wb = openpyxl.Workbook()
    ws = wb.active
    ws["Z1"], ws["Z2"], ws["Z5"] = 1.0, 2.0, 5.0
    for ref in ("B1", "B3", "B4"):
        ws[ref] = "x"
    for row, formula in LEGACY.items():
        ws[f"A{row}"] = formula
    wb.save(paths["src"])
    book = xlrecalc.load(paths["src"])
    for row, formula in FORMULAS.items():
        book.set_formula(f"A{row}", formula)  # dynamic-array formulas, as typed
    book.save()
    try:
        osascript("save_copy.applescript", str(paths["src"]), str(paths["excel"]))
        log(f"excel: Excel's copy\n    {cells(paths['excel'])}")
        book = xlrecalc.load(paths["excel"])
        log("  read: " + " ".join(repr(book.value(f"A{r}")) for r in range(1, ROWS + 1)))
        book.set("B1", None)
        book.set("Z5", 1.5)
        book.set("Z6", 1.0)
        log("  now: " + " ".join(repr(book.value(f"A{r}")) for r in range(1, ROWS + 1)))
        book.save(paths["written"], full_calc_on_load=False)
        stale_vm(paths["excel"], paths["written"], paths["stale_vm"])
        for name in ("written", "stale_vm"):
            path = paths[name]
            log(f"\n{name}:\n    {cells(path)}")
            fields = osascript("open_read.applescript", str(path), str(ROWS)).split(US)
            if fields[0] != "open":
                log(f"  REFUSED {fields[1:]}")
                continue
            shown = [" | ".join(fields[1 + 3 * i : 4 + 3 * i]) for i in range(ROWS)]
            log("  shown:\n    " + "\n    ".join(shown))
            copy = CONTAINER / f"xlrecalc_vm_{name}_saved.xlsx"
            osascript("save_copy.applescript", str(path), str(copy))
            log(f"  Excel's copy:\n    {cells(copy)}")
    finally:
        OUT.write_text("\n".join(lines) + "\n", encoding="utf-8")


if __name__ == "__main__":
    main()
