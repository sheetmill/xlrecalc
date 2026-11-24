"""What Excel shows after an edit makes a dynamic array spill further or less (§7.7).

    uv run python -m capture.spill_edit_probe

Non-interactive. For each case of ``capture/spill_edit_cases.py``, openpyxl writes the workbook
(sheet ``probe``, and ``other`` where a case names it), the dynamic-array formulas are put in as
Excel writes them (``add_dynamic``: ``<c cm="1"><f t="array">`` with the dynamic-array
metadata), and Excel opens it, calculates and saves a copy (``save_copy.applescript``). Excel then
opens the copy with alerts off and shows the case's cells, takes each step (a number entered in a
cell, in automatic calculation) and shows the cells after each (``edit_steps.applescript``); the
copy is closed without saving. What Excel shows is all the capture holds: it records no order in
which Excel computes the cells.

The workbooks live in Excel's container and are removed after; the capture is written to
``tests/captures/spill_edit.tsv``.
"""

from __future__ import annotations

import datetime as dt
import sys
from pathlib import Path

from capture.array_probe import add_dynamic
from capture.edition import edition_line
from capture.excel_watch import preflight
from capture.pipeline import CAPTURES, CONTAINER, excel_version
from capture.spill_edit_cases import CASES
from capture.writeback_probe import osascript

US = "\x1f"


def build(path: Path, case: str) -> None:
    import openpyxl  # imported here: the tests read the cases without it

    cells, _, _ = CASES[case]
    wb = openpyxl.Workbook()
    wb.active.title = "probe"
    if any(ref.startswith("other!") for ref in cells):
        wb.create_sheet("other")
    dynamic: list[tuple[str, str, str]] = []
    for ref, content in cells.items():
        sheet, where = ref.split("!")
        ws = wb[sheet]
        if isinstance(content, float):
            ws[where] = content
        elif content[0] == "#":
            assert sheet == "probe", "add_dynamic writes the first sheet only"
            ws[where] = "=0"  # a placeholder, replaced by the dynamic-array formula
            dynamic.append(("dynamic", where, content[1]))
        else:
            ws[where] = "=" + content[1]
    wb.save(path)
    add_dynamic(path, dynamic)


def main() -> None:
    preflight()
    CONTAINER.mkdir(parents=True, exist_ok=True)
    rows: list[str] = []
    for case, (_, steps, refs) in CASES.items():
        src = CONTAINER / f"xlrecalc_se_{case}_src.xlsx"
        path = CONTAINER / f"xlrecalc_se_{case}.xlsx"
        try:
            build(src, case)
            osascript("save_copy.applescript", str(src), str(path))
            fields = osascript("edit_steps.applescript", str(path), ";".join(steps), *refs)
        finally:
            for p in (src, path):
                p.unlink(missing_ok=True)
        got = fields.split(US)
        if got[0] != "open":
            sys.exit(f"{case} did not open: {got}")
        k = len(refs)
        shown = [got[1 : 1 + k]]
        for at in range(1 + k, len(got), 1 + k):  # each step: its label, then the cells
            if got[at] != steps[len(shown) - 1]:
                sys.exit(f"{case}: step {got[at]!r} where {steps[len(shown) - 1]!r} was taken")
            shown.append(got[at + 1 : at + 1 + k])
        if len(shown) != 1 + len(steps):
            sys.exit(f"{case}: {len(shown)} readings for {1 + len(steps)}")
        for i, ref in enumerate(refs):
            row = [case, ref, *(step[i] for step in shown)]
            rows.append("\t".join(row))
            print("  ".join(row), flush=True)
    header = [
        "# suite: spill_edit",
        f"# excel: {excel_version()} for Mac, locale en_US",
        edition_line(),
        f"# captured: {dt.date.today().isoformat()}",
        "# generator: capture/spill_edit_probe.py",
        "# method: each case of capture/spill_edit_cases.py written with its dynamic arrays as",
        "#   Excel writes them, saved by Excel, opened (automatic calculation), then each step",
        "#   taken (a number entered in a cell); each cell's value as Excel shows it, as text",
        "#   ('' blank), after opening and after each step",
        "# columns: case cell shown_after_opening then one column per step",
    ]
    out = CAPTURES / "spill_edit.tsv"
    out.write_text("\n".join([*header, *rows]) + "\n", encoding="utf-8")
    print(f"-> {out.name}, {len(rows)} rows")


if __name__ == "__main__":
    main()
