"""Loops of references that evaluation does not follow: does Excel call them circular?

    uv run python -m capture.static_cycle_probe

Non-interactive. Each group is a few cells in a column of its own, whose references form a loop
through a branch of ``IF`` that is not taken, or one that is, as a control. Excel saves a circular
formula as the empty text, so each cell's cached value in the copy shows what it did. Writes
``tests/captures/static_cycle.tsv``.
"""

from __future__ import annotations

import datetime as dt
import sys
from pathlib import Path

import openpyxl
from openpyxl.utils import get_column_letter

from capture.edition import edition_line
from capture.excel_watch import osascript_or_exit
from capture.pipeline import CAPTURES, CONTAINER, excel_version
from xlrecalc._xlsx import Book

HERE = Path(__file__).resolve().parent
# Each group: cells a, b, c... top to bottom in its column; {a} is the address of cell a.
GROUPS = [
    ("pair, not taken", ["=IF(TRUE,1,{b})", "={a}+1"]),
    ("pair, not taken, reader first", ["={b}+1", "=IF(TRUE,1,{a})"]),
    ("pair, on a constant", ["=IF({c}=0,1,{b})", "={a}+1", 0.0]),
    ("three cells", ["=IF({c}=0,1,{d})", "={a}+1", 0.0, "={b}*2"]),
    ("a dependent after", ["={b}+1", "={c}*2", "={d}+1", "=IF({e}=0,1,{b})", 0.0]),
    (
        "two readers, as a real workbook",
        [
            "=IF(ISNUMBER({d}),{e}+{d},{c})",
            "=IF(ISNUMBER({f}),{g}+{f},{c})",
            "={b}+{a}",
            0.0,
            2.0,
            0.0,
            3.0,
        ],
    ),
    ("CHOOSE, not taken", ["=CHOOSE(1,1,{b})", "={a}+1"]),
    ("IFERROR, not taken", ["=IFERROR(1,{b})", "={a}+1"]),
    ("control: taken", ["=IF(TRUE,{b},1)", "={a}+1"]),
    ("control: taken after a constant", ["=IF({c}=0,{b},1)", "={a}+1", 0.0]),
    ("control: a plain loop", ["={b}+1", "={a}+1"]),
    ("control: three in a loop", ["={b}", "={c}", "={a}"]),
]


def addresses(k: int, cells: list[object]) -> dict[str, str]:
    column = get_column_letter(1 + k)
    return {"abcdefghij"[i]: f"{column}{i + 1}" for i in range(len(cells))}


def build(path: Path) -> None:
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "probe"
    for k, (_, cells) in enumerate(GROUPS):
        names = addresses(k, cells)
        for i, cell in enumerate(cells):
            value = cell.format(**names) if isinstance(cell, str) else cell
            ws[f"{get_column_letter(1 + k)}{i + 1}"] = value
    wb.save(path)


def main() -> None:
    CONTAINER.mkdir(parents=True, exist_ok=True)
    src = CONTAINER / "xlrecalc_static_cycle.xlsx"
    copy = CONTAINER / "xlrecalc_static_cycle_saved.xlsx"
    for p in (src, copy):
        p.unlink(missing_ok=True)
    build(src)
    version = excel_version()
    result = osascript_or_exit(
        str(HERE / "save_copy.applescript"), str(src), str(copy), seconds=180
    )
    if result.returncode:
        sys.exit(f"Excel did not open the workbook: {result.stderr.strip()}")
    rows = [
        "# suite: static_cycle",
        f"# excel: {version} for Mac, locale en_US",
        edition_line(),
        f"# captured: {dt.date.today().isoformat()}",
        "# generator: capture/static_cycle_probe.py",
        "# columns: group, cell, formula (repr), cached value (repr) in the copy Excel saved",
    ]
    with Book(copy) as book:
        cells = {(c.row, c.column): c for c in book.cells("probe")}
    for k, (name, group) in enumerate(GROUPS):
        names = addresses(k, group)
        for i, item in enumerate(group):
            text = item.format(**names) if isinstance(item, str) else item
            cell = cells[i + 1, k + 1]
            rows.append(f"{name}\t{names['abcdefghij'[i]]}\t{text!r}\t{cell.value!r}")
    (CAPTURES / "static_cycle.tsv").write_text("\n".join(rows) + "\n", encoding="utf-8")
    for p in (src, copy):
        p.unlink(missing_ok=True)
    print(f"{len(rows) - 5} rows -> static_cycle.tsv")


if __name__ == "__main__":
    main()
