"""What ``[0]!Name`` means to Excel: a name of the workbook itself (ARCHITECTURE.md §5.2).

    uv run python -m capture.workbook_name_probe

Non-interactive. One workbook: a workbook-level ``Rate`` and ``Range`` on sheet ``s``, a sheet-level
``Rate`` on sheet ``t``, and formulas reading them as ``[0]!Rate`` from both sheets, beside ``Rate``
alone and a name that does not exist (``[1]!Rate`` with no external link makes Excel refuse the
file). Excel opens it and saves a copy (``save_copy.applescript``); each formula's cached value
and its text in the copy are read.
Writes ``tests/captures/workbook_name.tsv``.
"""

from __future__ import annotations

import datetime as dt
import sys
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET

import openpyxl
from openpyxl.workbook.defined_name import DefinedName

from capture.edition import edition_line
from capture.excel_watch import osascript_or_exit
from capture.pipeline import CAPTURES, CONTAINER, excel_version

HERE = Path(__file__).resolve().parent
NS = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
FORMULAS = {
    "s": ["[0]!Rate*2", "Rate*2", "SUM([0]!Range)", "[0]!Nope", "[0]!rate+0"],
    "t": ["[0]!Rate*2", "Rate*2"],
}


def build(path: Path) -> None:
    wb = openpyxl.Workbook()
    s = wb.active
    s.title = "s"
    t = wb.create_sheet("t")
    for r, v in enumerate((3, 4, 5), 1):
        s.cell(r, 2, v)
    t["B1"] = 50
    wb.defined_names["Rate"] = DefinedName("Rate", attr_text="s!$B$1")
    wb.defined_names["Range"] = DefinedName("Range", attr_text="s!$B$1:$B$3")
    t.defined_names["Rate"] = DefinedName("Rate", attr_text="t!$B$1")
    for name, formulas in FORMULAS.items():
        for r, f in enumerate(formulas, 1):
            wb[name].cell(r, 1, "=" + f)
    wb.save(path)


def read(path: Path) -> dict[tuple[str, int], tuple[str, str, str]]:
    """Each column-A formula's (formula text, type, cached value) in the copy Excel saved."""
    out = {}
    with zipfile.ZipFile(path) as z:
        for i, name in enumerate(FORMULAS, 1):
            for c in ET.fromstring(z.read(f"xl/worksheets/sheet{i}.xml")).iter(f"{NS}c"):
                ref = c.get("r", "")
                f, v = c.find(f"{NS}f"), c.find(f"{NS}v")
                if ref.startswith("A") and f is not None:
                    value = "" if v is None or v.text is None else v.text
                    out[(name, int(ref[1:]))] = (f.text or "", c.get("t", "n"), value)
    return out


def main() -> None:
    CONTAINER.mkdir(parents=True, exist_ok=True)
    path = CONTAINER / "xlrecalc_workbook_name.xlsx"
    copy = CONTAINER / "xlrecalc_workbook_name_saved.xlsx"
    try:
        build(path)
        copy.unlink(missing_ok=True)
        done = osascript_or_exit(
            str(HERE / "save_copy.applescript"), str(path), str(copy), seconds=90
        )
        if done.returncode:
            sys.exit(f"Excel did not save a copy: {done.stderr.strip()}")
        saved = read(copy)
    finally:
        path.unlink(missing_ok=True)
        copy.unlink(missing_ok=True)
    lines = [
        "# suite: workbook_name",
        f"# excel: {excel_version()} for Mac, locale en_US",
        edition_line(),
        f"# captured: {dt.date.today().isoformat()}",
        "# generator: capture/workbook_name_probe.py",
        "# method: workbook-level Rate (s!$B$1 = 3) and Range (s!$B$1:$B$3 = 3, 4, 5),",
        "#   sheet-level Rate on t (t!$B$1 = 50); openpyxl writes the formulas, Excel saves a",
        "#   copy; saved_formula is the formula's text in that copy. With [1]!Rate*2 added (no",
        "#   external link in the file) Excel refuses the whole file (-50)",
        "# columns: sheet row formula saved_formula type value",
    ]
    for (sheet, row), (text, kind, value) in sorted(saved.items()):
        lines.append("\t".join([sheet, str(row), FORMULAS[sheet][row - 1], text, kind, value]))
        print(lines[-1], flush=True)
    (CAPTURES / "workbook_name.tsv").write_text("\n".join(lines) + "\n", encoding="utf-8")


if __name__ == "__main__":
    main()
