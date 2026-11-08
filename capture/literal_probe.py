"""Which number literals Excel accepts in a file formula (``CAPTURING.md`` §1).

    uv run python -m capture.literal_probe

Non-interactive. Each formula is written alone into a workbook openpyxl builds, and Excel opens
it and saves a copy (``save_copy.applescript``); a workbook Excel refuses (error -50, a repair
prompt with alerts off) has no copy. Writes ``tests/captures/number_literals.tsv``: the formula,
whether Excel opened the file, and the formula text and cached value in the copy.
"""

from __future__ import annotations

import datetime as dt
import sys
from pathlib import Path

import openpyxl

from capture.excel_watch import osascript_or_exit
from capture.pipeline import CAPTURES, CONTAINER, excel_version
from xlrecalc._xlsx import Book
from xlrecalc._xlsx._write import number_text

HERE = Path(__file__).resolve().parent

FORMULAS = [
    "9.99999999999999E+307",
    "9.99999999999999E+307*1",
    "9.999999999999999E+307",
    "9.999999999999995E+307",
    "1E+308",
    "1E308",
    "-1E+308",
    "1.7976931348623157E+308",
    "1E+309",
    "2*1E+308",
    '"1E+308"*1',
    "1E-307",
    "2.2250738585072014E-308",
    "1E-308",
    "1E-320",
    "1E-400",
    # a second set: other mantissas past the limit, and inside it
    "9.9999999999999999E+307",
    "1.5E+308",
    "99999999999999E+294",
    "999999999999999E+293",
    "0.1E+309",
    # the small end, between the last accepted and the first refused
    "9E-309",
    "1E-309",
    "1E-310",
    "1E-315",
]


def run_one(formula: str) -> tuple[str, str, str]:
    src = CONTAINER / "xlrecalc_literal.xlsx"
    copy = CONTAINER / "xlrecalc_literal_saved.xlsx"
    for p in (src, copy):
        p.unlink(missing_ok=True)
    wb = openpyxl.Workbook()
    wb.active["A1"] = "=" + formula
    wb.save(src)
    result = osascript_or_exit(
        str(HERE / "save_copy.applescript"), str(src), str(copy), seconds=120.0
    )
    if result.returncode:
        src.unlink(missing_ok=True)
        return "refused", "", ""
    with Book(copy) as book:
        cell = next(c for c in book.cells(book.sheets[0]) if c.row == 1 and c.column == 1)
        text = cell.formula.text if cell.formula else ""
        value = number_text(cell.value) if isinstance(cell.value, float) else str(cell.value)
    for p in (src, copy):
        p.unlink(missing_ok=True)
    return "opens", text or "", f"{cell.type}:{value}"


def main() -> None:
    CONTAINER.mkdir(parents=True, exist_ok=True)
    version = excel_version()
    header = [
        "# suite: number_literals",
        f"# excel: {version} for Mac, locale en_US",
        f"# captured: {dt.date.today().isoformat()}",
        "# generator: capture/literal_probe.py (each formula alone in a workbook)",
        "# columns: formula as written, opens or refused, formula Excel saved, type:value",
    ]
    rows = []
    for formula in FORMULAS:
        outcome, saved, value = run_one(formula)
        rows.append("\t".join([formula, outcome, saved, value]))
        print(formula, outcome, saved, value, flush=True)
    (CAPTURES / "number_literals.tsv").write_text(
        "\n".join([*header, *rows]) + "\n", encoding="utf-8"
    )
    sys.exit(0)


if __name__ == "__main__":
    main()
