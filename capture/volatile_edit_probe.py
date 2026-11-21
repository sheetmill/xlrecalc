"""Which formulas an edit makes Excel recompute after it kept their values on opening.

    uv run python -m capture.volatile_edit_probe

One workbook a variant: openpyxl writes it (``probe!A1`` = 0.1, ``B1`` the formula under test,
``C1`` = 1, ``D1 = C1*10``, ``E1 = A1*3``, and the variant's defined name and cells) and Excel
saves a copy, so the file has Excel's own ``calcId``, ``<fileVersion>`` and cell forms. In that
copy B1's ``ca="1"`` mark is removed or written, as the variant asks, and the cached values of
B1, D1 and E1 (and F1) are planted (7, 999, 999, 999: only the ``<v>`` text). Excel opens it
with alerts off and shows the cells, C1 is set to 2, the cells are shown again, Excel calculates
in full (``calculate full``), and they are shown a third time (``edit_full_calc.applescript``);
the workbook is closed without saving.

B1 reads C1 only through ``INDIRECT``'s text, so it changes after the edit only if Excel computes
it as volatile; ``name_dependent_no_ca`` names C1 (``OFFSET(probe!$C$1,0,0)``), so there B1 is a
dependent of the edit too. D1 is a plain dependent of the edit; E1 is reached by nothing the edit
changes. ``PREDICTED`` holds what was expected before the run, per cell: after opening, after the
edit, after the full calculation.

The workbooks live in Excel's container and are removed after; the capture is written to
``tests/captures/volatile_edit.tsv``.
"""

from __future__ import annotations

import datetime as dt
import re
import subprocess
import sys

from capture.edition import edition_line
from capture.excel_watch import preflight
from capture.open_calc_probe import calc_attributes, edit_calc
from capture.pipeline import CAPTURES
from capture.volatile_edit_cases import EDITED, FULL_FIRST, FULL_ON_LOAD, VARIANTS
from capture.volatile_name_probe import cell_xml
from capture.writeback_probe import CONTAINER, osascript

US = "\x1f"
PLANTED = {"B1": "7", "D1": "999", "E1": "999", "F1": "999"}

# Expected before the run: (after opening, after the edit, after the full calculation).
PREDICTED = {
    "name_no_ca": {"B1": (7, 7, 2), "D1": (999, 20, 20), "E1": (999, 999, 0.3)},
    "name_ca": {"B1": (1, 2, 2), "D1": (999, 20, 20), "E1": (999, 999, 0.3)},
    "own_no_ca": {"B1": (1, 2, 2), "D1": (10, 20, 20), "E1": (0.3, 0.3, 0.3)},
    "now_no_ca": {
        "B1": (0.1, 0.1, 0.1),
        "D1": (10, 20, 20),
        "E1": (0.3, 0.3, 0.3),
        "F1": (3, 3, 3),
    },
    "name_dependent_no_ca": {"B1": (7, 2, 2), "D1": (999, 20, 20), "E1": (999, 999, 0.3)},
    # plain_name, name_offset_plus, name_static_dependent, name_no_ca_full_first and
    # name_no_ca_full_on_load, expected before they ran (after opening, [after a first full
    # calculation,] after the edit, after the full calculation): a name's references followed as
    # any others but OFFSET's reference argument, and a cell Excel computed tracked as volatile
    # from then on.
    "plain_name": {"B1": (7, 2, 2)},
    "name_offset_plus": {"B1": (7, 2, 2)},
    "name_static_dependent": {"B1": (7, 3, 3), "D1": (999, 999, 10), "E1": (999, 6, 6)},
    "name_no_ca_full_first": {"B1": (7, 1, 2, 2), "D1": (999, 10, 20, 20)},
    "name_no_ca_full_on_load": {"B1": (1, 2, 2), "D1": (10, 20, 20), "E1": (0.3, 0.3, 0.3)},
}


def build(path: object, formula: str, name: str | None, extra: dict[str, str]) -> None:
    import openpyxl  # imported here: the tests read VARIANTS without it
    from openpyxl.workbook.defined_name import DefinedName

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "probe"
    ws["A1"] = 0.1
    ws["B1"] = formula
    ws["C1"] = 1
    ws["D1"] = "=C1*10"
    ws["E1"] = "=A1*3"
    for ref, text in extra.items():
        ws[ref] = text
    if name is not None:
        wb.defined_names["vn"] = DefinedName("vn", attr_text=name)
    wb.save(path)


def plant(xml: str, marked: bool) -> str:
    """B1 marked or not; the planted cells' ``<v>`` text replaced, the rest as Excel wrote it.
    F1 keeps the mark Excel wrote."""
    xml = re.sub(r'(<c r="B1"[^>]*><f)\s+ca="1"', r"\1", xml, count=1)
    if marked:
        xml, n = re.subn(r'(<c r="B1"[^>]*><f)>', r'\1 ca="1">', xml, count=1)
        if n != 1:
            raise RuntimeError("B1's formula has attributes: cannot mark it")
    for ref, value in PLANTED.items():
        cell = re.compile(rf'(<c r="{ref}"[^>]*>.*?<v>)(.*?)(</v>)', re.S)
        xml = cell.sub(lambda m, v=value: m.group(1) + v + m.group(3), xml, count=1)
    return xml


def _same(text: str) -> str:
    return text


def _flag(text: str) -> str:
    """``fullCalcOnLoad="1"`` in workbook.xml's ``<calcPr>``."""
    return re.sub(r"<calcPr\b", '<calcPr fullCalcOnLoad="1"', text, count=1)


def main() -> None:
    preflight()
    CONTAINER.mkdir(parents=True, exist_ok=True)
    rows: list[str] = []
    calc = "-"
    for variant, (formula, name, marked, extra) in VARIANTS.items():
        src = CONTAINER / f"xlrecalc_ve_{variant}_src.xlsx"
        base = CONTAINER / f"xlrecalc_ve_{variant}_base.xlsx"
        path = CONTAINER / f"xlrecalc_ve_{variant}.xlsx"
        cells = ["A1", "B1", "C1", "D1", "E1", *extra]
        try:
            build(src, formula, name, extra)
            osascript("save_copy.applescript", str(src), str(base))
            full_on_load = variant in FULL_ON_LOAD
            edit_calc(base, path, lambda x, m=marked: plant(x, m), _flag if full_on_load else _same)
            if not full_on_load:
                calc = calc_attributes(path)
            planted = {ref: cell_xml(path, ref) for ref in cells}
            refs = [f"probe!{ref}" for ref in cells]
            target = ("first:" if variant in FULL_FIRST else "") + EDITED.get(variant, "probe!C1")
            fields = osascript("edit_full_calc.applescript", str(path), target, "2", *refs)
        finally:
            for p in (src, base, path):
                p.unlink(missing_ok=True)
        fields_list = fields.split(US)
        if fields_list[0] != "open":
            sys.exit(f"{variant} did not open: {fields_list}")
        k = len(cells)
        steps = [fields_list[1 : 1 + k]]
        for at in range(1 + k, len(fields_list), 1 + k):  # each step: its label, then the cells
            steps.append(fields_list[at + 1 : at + 1 + k])
        if variant not in FULL_FIRST:
            steps.insert(1, ["-"] * k)
        for i, ref in enumerate(cells):
            row = [variant, f"probe!{ref}", planted[ref], *(step[i] for step in steps)]
            rows.append("\t".join(row))
            print("  ".join(row), flush=True)
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
        "# method: Excel's saved copy, B1 marked or not, B1 D1 E1 F1 planted, fullCalcOnLoad",
        "#   written where the variant asks; shown after opening, after a calculate full where",
        "#   the variant asks ('-' else), after C1 (A1 where the variant edits it) is set to 2,",
        "#   and after calculate full; the workbook line is that of the files without the flag",
        "# columns: variant cell planted_xml shown_after_opening shown_after_full_first",
        "#   shown_after_edit shown_after_full",
    ]
    out = CAPTURES / "volatile_edit.tsv"
    out.write_text("\n".join([*header, *rows]) + "\n", encoding="utf-8")
    print(f"-> {out.name}, {len(rows)} rows")


if __name__ == "__main__":
    main()
