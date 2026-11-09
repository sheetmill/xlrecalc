"""Functions that make arrays, as Excel spills them (ARCHITECTURE.md §7.7).

    uv run python -m capture.dynamic_probe SUITE [SUITE ...]

Non-interactive. Suites are in ``capture/dynamic_suites.py``. One sheet holds the suite's data
(columns AA on) and, for each case, a dynamic-array formula (``<c cm="1"><f t="array">`` with
the dynamic-array metadata, as Excel writes a formula entered in Excel 365) at the anchor of its
own block of rows in column B, with any cells the case puts in its way and any plain formulas
reading it (``{a}`` in their text is the anchor). Excel opens the workbook, calculates and saves
a copy (``save_copy.applescript``); every cell of the copy is recorded with its value, formula,
the formula's ``t`` and ``ref``, ``cm`` and ``vm``, and the copy's rich-value parts are kept.
Writes ``tests/captures/dyn_<suite>.tsv`` and ``tests/captures/dyn_<suite>_rd*.xml``.
"""

from __future__ import annotations

import ast
import datetime as dt
import sys
import zipfile
from pathlib import Path

import openpyxl

from capture.array_probe import add_dynamic
from capture.dynamic_suites import BLOCK, SUITES, anchor_of, layout
from capture.edition import edition_line
from capture.excel_watch import osascript_or_exit
from capture.pipeline import CAPTURES, CONTAINER, excel_version
from xlrecalc._values import ErrorValue
from xlrecalc._xlsx import Book
from xlrecalc._xlsx._write import number_text

HERE = Path(__file__).resolve().parent


def _col(n: int) -> str:
    s = ""
    while n:
        n, r = divmod(n - 1, 26)
        s = chr(65 + r) + s
    return s


def build(path: Path, suite: str) -> list[tuple[str, str, str]]:
    """The workbook for a suite; returns its (kind, cell, text) inputs."""
    inputs = layout(suite)
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "probe"
    dynamic = []
    for kind, where, text in inputs:
        if kind == "name":
            from openpyxl.workbook.defined_name import DefinedName

            wb.defined_names[where] = DefinedName(where, attr_text=text)
        elif kind == "dynamic":
            ws[where] = "=0"  # a placeholder, replaced by the dynamic-array formula
            dynamic.append((kind, where, text))
        elif kind in ("legacy", "reader"):
            ws[where] = "=" + text
        else:  # a constant: data, or a cell in a result's way
            value = eval_constant(text)
            cell = ws[where]
            cell.value = str(value) if isinstance(value, ErrorValue) else value
            if isinstance(value, ErrorValue):
                cell.data_type = "e"
    wb.save(path)
    add_dynamic(path, dynamic)
    return inputs


def eval_constant(text: str) -> object:
    """A constant from its repr: a float, a str, a bool or an ErrorValue."""
    if text.startswith("ErrorValue("):
        return ErrorValue(ast.literal_eval(text[len("ErrorValue(") : -1]))
    return ast.literal_eval(text)


def run(suite: str) -> None:
    CONTAINER.mkdir(parents=True, exist_ok=True)
    src = CONTAINER / f"xlrecalc_dyn_{suite}.xlsx"
    copy = CONTAINER / f"xlrecalc_dyn_{suite}_saved.xlsx"
    for p in (src, copy):
        p.unlink(missing_ok=True)
    inputs = build(src, suite)
    version = excel_version()
    result = osascript_or_exit(
        str(HERE / "save_copy.applescript"), str(src), str(copy), seconds=300.0
    )
    if result.returncode:
        src.unlink(missing_ok=True)
        sys.exit(f"Excel did not open the workbook: {result.stderr.strip()}")
    header = [
        f"# suite: dyn_{suite}",
        f"# excel: {version} for Mac, locale en_US",
        edition_line(),
        f"# captured: {dt.date.today().isoformat()}",
        "# generator: capture/dynamic_probe.py",
        "# method: dynamic-array formulas (cm=1 with the XLDAPR metadata), one per block of "
        f"{BLOCK} rows in column B, over data in columns AA on; Excel calculates and saves a copy;",
        "#   every cell of the copy is recorded",
        "# input rows: input, kind, cell, text (formula in file form, or a constant's repr)",
        "# cell rows: cell, address, type, value (repr), formula (repr), f_t, f_ref, cm, vm",
    ]
    body = ["\t".join(["input", kind, where, repr(text)]) for kind, where, text in inputs]
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
    out = CAPTURES / f"dyn_{suite}.tsv"
    out.write_text("\n".join([*header, *body]) + "\n", encoding="utf-8")
    with zipfile.ZipFile(copy) as z:
        for name in z.namelist():  # the rich values behind errors such as #SPILL! and #CALC!
            if name.startswith("xl/richData/") and name.endswith(".xml"):
                (CAPTURES / f"dyn_{suite}_{Path(name).name}").write_bytes(z.read(name))
    print(f"dyn_{suite}: {len(inputs)} inputs, {len(body) - len(inputs)} cells -> {out.name}")
    for p in (src, copy):
        p.unlink(missing_ok=True)


def main() -> None:
    names = sys.argv[1:]
    unknown = [n for n in names if n not in SUITES]
    if not names or unknown:
        sys.exit(f"usage: python -m capture.dynamic_probe SUITE ... (one of {sorted(SUITES)})")
    for name in names:
        run(name)


__all__ = ["anchor_of", "build", "main", "run"]

if __name__ == "__main__":
    main()
