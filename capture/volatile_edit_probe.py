"""Does an edit make Excel recompute a volatile cell it kept on opening?

    uv run python -m capture.volatile_edit_probe

One workbook: openpyxl writes it (``probe!A1`` = 0.1, ``B1 = vn`` with the defined name ``vn`` =
``OFFSET(probe!$A$1,0,0)``, ``C1`` = 1, ``E1 = A1*3``) and Excel saves a copy, so the file has
Excel's own ``calcId``, ``<fileVersion>`` and cell forms. In that copy B1's ``ca="1"`` mark is
removed and its cached value planted as 7, and E1's as 999 (only the ``<v>`` text): opening keeps
both (capture ``volatile_names``, its ``_no_ca`` variants). Excel opens it with alerts off and
shows the cells, C1 is set to 2, which neither formula reads, and the cells are shown again
(``iterate_edit.applescript``); the workbook is closed without saving. Predicted before the run:
B1 7 then 0.1 (a volatile formula is computed at every recalculation, an edit's too), E1 999
both times (nothing it reads changed).

The workbooks live in Excel's container and are removed after; the capture is written to
``tests/captures/volatile_edit.tsv``.
"""

from __future__ import annotations

import datetime as dt
import re
import subprocess

import openpyxl
from openpyxl.workbook.defined_name import DefinedName

from capture.edition import edition_line
from capture.excel_watch import preflight
from capture.open_calc_probe import calc_attributes, edit_calc
from capture.pipeline import CAPTURES
from capture.volatile_name_probe import cell_xml
from capture.writeback_probe import CONTAINER, osascript

US = "\x1f"
CELLS = ["probe!A1", "probe!B1", "probe!C1", "probe!E1"]
PLANTED = {"B1": "7", "E1": "999"}


def build(path: object) -> None:
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "probe"
    ws["A1"] = 0.1
    ws["B1"] = "=vn"
    ws["C1"] = 1
    ws["E1"] = "=A1*3"
    wb.defined_names["vn"] = DefinedName("vn", attr_text="OFFSET(probe!$A$1,0,0)")
    wb.save(path)


def plant(xml: str) -> str:
    """B1 without ``ca="1"``; B1's and E1's ``<v>`` text replaced, the rest as Excel wrote it."""
    xml = re.sub(r'(<c r="B1"[^>]*><f)\s+ca="1"', r"\1", xml, count=1)
    for ref, value in PLANTED.items():
        cell = re.compile(rf'(<c r="{ref}"[^>]*>.*?<v>)(.*?)(</v>)', re.S)
        xml, n = cell.subn(lambda m, v=value: m.group(1) + v + m.group(3), xml, count=1)
        if n != 1:
            raise RuntimeError(f"no cached value in {ref}: cannot plant")
    return xml


def main() -> None:
    preflight()
    CONTAINER.mkdir(parents=True, exist_ok=True)
    src = CONTAINER / "xlrecalc_ve_src.xlsx"
    base = CONTAINER / "xlrecalc_ve_base.xlsx"
    path = CONTAINER / "xlrecalc_ve.xlsx"
    try:
        build(src)
        print("base:", osascript("save_copy.applescript", str(src), str(base)), flush=True)
        edit_calc(base, path, plant, lambda text: text)
        planted = {ref: cell_xml(path, ref) for ref in ("B1", "E1")}
        calc = calc_attributes(path)
        print("planted:", planted, "calcPr", calc, flush=True)
        fields = osascript("iterate_edit.applescript", str(path), "probe!C1", "2", *CELLS).split(US)
        if fields[0] != "open":
            raise SystemExit(f"did not open: {fields}")
        k = len(CELLS)
        opened, edited = fields[1 : 1 + k], fields[2 + k : 2 + 2 * k]
    finally:
        for p in (src, base, path):
            p.unlink(missing_ok=True)
    version = subprocess.run(
        ["osascript", "-e", 'tell application "Microsoft Excel" to version'],
        capture_output=True,
        text=True,
    ).stdout.strip()
    header = [
        "# suite: volatile_edit",
        f"# excel: {version} for Mac, locale en_US",
        edition_line(),
        f"# captured: {dt.date.today().isoformat()}",
        "# generator: capture/volatile_edit_probe.py",
        f"# workbook: {calc}",
        "# method: Excel's saved copy, B1's ca removed, B1 and E1 planted; shown after opening,",
        "#   then after C1 is set to 2",
        "# columns: cell planted_xml shown_after_opening shown_after_edit",
    ]
    rows = [
        "\t".join([ref, planted.get(ref.split("!")[1], "-"), opened[i], edited[i]])
        for i, ref in enumerate(CELLS)
    ]
    out = CAPTURES / "volatile_edit.tsv"
    out.write_text("\n".join([*header, *rows]) + "\n", encoding="utf-8")
    print("\n".join(rows))


if __name__ == "__main__":
    main()
