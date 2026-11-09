"""Which formulas make Excel refuse a file, each alone in its own workbook.

    uv run python -m capture.refusal_probe FORMULA [FORMULA ...]
    uv run python -m capture.refusal_probe --suite NAME

Non-interactive. Each formula (file form, no ``=``) is written as a dynamic-array formula at B1
of a workbook holding the data of ``capture/dynamic_suites.py``; Excel opens it and saves a copy
(``save_copy.applescript``). With alerts off, a file Excel would offer to repair fails to open
with error -50, so a formula that does is one Excel refuses. ``--suite`` tries every case of a
suite one by one, to find the case that makes a suite's workbook fail. Prints one line per
formula, ``opens`` or ``refused``, and writes nothing else; both files are removed after each.
"""

from __future__ import annotations

import sys
from pathlib import Path

import openpyxl

from capture.array_probe import add_dynamic
from capture.dynamic_probe import eval_constant
from capture.dynamic_suites import DATA, SUITES
from capture.excel_watch import osascript_or_exit
from capture.pipeline import CONTAINER
from xlrecalc._values import ErrorValue

HERE = Path(__file__).resolve().parent


def build(path: Path, formula: str) -> None:
    """A workbook with the suites' data and one dynamic-array formula at B1."""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "probe"
    for letters, values in DATA.items():
        for r, v in enumerate(values, 1):
            if v is None:
                continue
            cell = ws[f"{letters}{r}"]
            if isinstance(v, str) and v.startswith("="):
                cell.value = v
                continue
            value = eval_constant(repr(v))
            cell.value = str(value) if isinstance(value, ErrorValue) else value
            if isinstance(value, ErrorValue):
                cell.data_type = "e"
    ws["B1"] = "=0"
    wb.save(path)
    add_dynamic(path, [("dynamic", "B1", formula)])


def opens(formula: str) -> bool:
    """Whether Excel opens a workbook holding the formula."""
    CONTAINER.mkdir(parents=True, exist_ok=True)
    src = CONTAINER / "xlrecalc_refusal.xlsx"
    copy = CONTAINER / "xlrecalc_refusal_saved.xlsx"
    for p in (src, copy):
        p.unlink(missing_ok=True)
    build(src, formula)
    result = osascript_or_exit(
        str(HERE / "save_copy.applescript"), str(src), str(copy), seconds=120.0
    )
    for p in (src, copy):
        p.unlink(missing_ok=True)
    if result.returncode and "-50" not in result.stderr:
        sys.exit(f"Excel failed otherwise on {formula}: {result.stderr.strip()}")
    return result.returncode == 0


def main() -> None:
    args = sys.argv[1:]
    if args[:1] == ["--suite"] and len(args) == 2 and args[1] in SUITES:
        formulas = [case.formula for case in SUITES[args[1]]]
    elif args and not args[0].startswith("--"):
        formulas = args
    else:
        sys.exit("usage: python -m capture.refusal_probe FORMULA ... | --suite NAME")
    for formula in formulas:
        print(f"{'opens' if opens(formula) else 'refused'}\t{formula}", flush=True)


if __name__ == "__main__":
    main()
