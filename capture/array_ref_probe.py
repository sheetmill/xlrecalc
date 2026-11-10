"""Functions that take a reference and may return an array, as Excel computes them (§7.7).

    uv run python -m capture.array_ref_probe

Non-interactive. Sheet ``probe`` holds a data block (Z1:Z5 = 1, 2, 2, 3, 5; AA1:AA5 = 10..50;
AB1:AB5 = a..e), sheet ``other`` the same block, and the name ``Block`` is ``probe!$Z$1:$Z$4``.
Each formula goes in as:

- a plain legacy formula: in column A (outside the block's rows and columns), in row 3 (inside its
  rows) and in column AA (inside its columns);
- a legacy array formula over a 4 by 3 range, and ``ROW(Z1:Z4)`` over other shapes too;
- a dynamic-array formula (the marker as Excel saves it, ``capture/array_probe.py``), unless it
  would spill over a whole row or column.

The formulas: ``ROW`` and ``COLUMN`` of ranges of every form, alone and inside other functions,
and of whole rows, columns and the whole sheet, inside aggregates only; ``CELL`` of a range;
reference functions not implemented yet, recorded for when they are; arrays of references
(``N(OFFSET(Z1,{0;1},0))``); ranges and arrays as a criterion or a lookup key; and functions
already taken element by element, as controls.

Excel opens the workbook, calculates and saves a copy; every cell of the copy is recorded, as in
``capture/array_probe.py``. Writes ``tests/captures/array_refs.tsv``.
"""

from __future__ import annotations

import datetime as dt
import subprocess
import sys
from pathlib import Path

import openpyxl
from openpyxl.utils import get_column_letter
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.worksheet.formula import ArrayFormula

from capture.array_probe import add_dynamic
from capture.edition import edition_line
from capture.pipeline import CAPTURES, CONTAINER, excel_version
from xlrecalc._xlsx import Book
from xlrecalc._xlsx._write import number_text

HERE = Path(__file__).resolve().parent
DATA = {"Z": [1, 2, 2, 3, 5], "AA": [10, 20, 30, 40, 50], "AB": list("abcde")}

SHAPES = ["Z1:Z4", "Z1:AB1", "Z1:AA3"]
OTHER_FORMS = [
    "Z2",
    "other!Z1:Z4",
    "OFFSET(Z1,1,0,3)",
    "INDEX(Z1:Z4,0)",
    "Z1:INDEX(Z1:Z4,3)",
    "Block",
    "(Z1:Z2,AA1:AA2)",
    "Z1:Z4 Z2:AA2",
]
WRAPPERS = [
    "{x}",
    "SUM({x})",
    "SUMPRODUCT({x})",
    "MAX({x})",
    "COUNT({x})",
    "AVERAGE({x})",
    "{x}*2",
    "SUM({x}*2)",
    "ABS({x})",
    "IF(TRUE,{x})",
    "INDEX({x},2)",
    "ROWS({x})",
    "COLUMNS({x})",
    "MATCH(2,{x},0)",
]
AGGREGATES = ["SUM({x})", "SUMPRODUCT({x})", "MAX({x})", "COUNT({x})", "ROWS({x})", "COLUMNS({x})"]
ARRAY_REFS = "{0;1;2;3}"
KEYS = ["Z1:Z4", "{1;2}", "{1,2}", "Z2"]
KEYED = [
    "COUNTIF(Z1:Z4,{k})",
    "SUMIF(Z1:Z4,{k})",
    "SUMIF(Z1:Z4,{k},AA1:AA4)",
    "COUNTIFS(Z1:Z4,{k})",
    'COUNTIFS(Z1:Z4,{k},AA1:AA4,">0")',
    "SUMIFS(AA1:AA4,Z1:Z4,{k})",
    "MATCH({k},Z1:Z4,0)",
    "MATCH({k},{{1;2;3}},0)",
    "VLOOKUP({k},Z1:AA4,2,FALSE)",
    "HLOOKUP({k},Z1:AA4,1,FALSE)",
    "LOOKUP({k},Z1:Z4)",
]


def formulas() -> list[tuple[str, bool]]:
    """(formula, whether it may spill over a whole row or column), in a fixed order."""
    out: list[tuple[str, bool]] = []
    for fn in ("ROW", "COLUMN"):
        for ref in SHAPES:
            x = f"{fn}({ref})"
            out += [(w.format(x=x), False) for w in WRAPPERS]
            out.append((f"{x}-{fn}(Z1)+1", False))
        for ref in OTHER_FORMS:
            x = f"{fn}({ref})"
            out += [(w.format(x=x), False) for w in WRAPPERS[:3]]
        for ref in ("Z:Z", "2:3", "1:1048576", "A:XFD", "A:J"):
            x = f"{fn}({ref})"
            out += [(w.format(x=x), False) for w in AGGREGATES]
            spills = (fn == "ROW" and ref in ("Z:Z", "1:1048576", "A:XFD", "A:J")) or (
                fn == "COLUMN" and ref in ("2:3", "1:1048576", "A:XFD")
            )
            out.append((x, spills))
    out += [
        ("SUMPRODUCT(ROW(Z1:Z4)*COLUMN(Z1:AB1))", False),
        ("SUM(ROW(Z1:Z4)*COLUMN(Z1:AB1))", False),
        ("ROW(Z1:Z4)*COLUMN(Z1:AB1)", False),
        ("ADDRESS(ROW(Z1:Z4),COLUMN(Z1:Z4))", False),
        ("SUM(INDEX(Z1:Z4,ROW(Z1:Z2)))", False),
        ("SUMPRODUCT(INDEX(Z1:Z4,ROW(Z1:Z2)))", False),
    ]
    for kind in ("row", "col", "address", "contents", "type"):
        for ref in SHAPES:
            x = f'CELL("{kind}",{ref})'
            out += [(x, False), (f"SUMPRODUCT({x})", False)]
    for name in ("_xlfn.ISFORMULA", "_xlfn.FORMULATEXT", "_xlfn.SHEET", "ISREF", "AREAS"):
        for ref in ("Z1:Z4", "Z1:AB1"):
            x = f"{name}({ref})"
            out += [(x, False), (f"SUMPRODUCT(--{x})", False)]
    offsets = f"OFFSET(Z1,{ARRAY_REFS},0)"
    texts = 'INDIRECT({"Z1";"Z2";"Z3";"Z4"})'
    for wrapped in (
        "N(Z1:Z4)",
        "SUM(N(Z1:Z4))",
        "SUMPRODUCT(N(Z1:Z4))",
        "N(Z1:Z4*1)",
        f"N({offsets})",
        f"SUM(N({offsets}))",
        f"SUMPRODUCT(N({offsets}))",
        f"SUBTOTAL(9,{offsets})",
        f"SUMPRODUCT(SUBTOTAL(9,{offsets}))",
        f'COUNTIF({offsets},">0")',
        f'SUMPRODUCT(COUNTIF({offsets},">0"))',
        f"SUM({offsets})",
        f"ROW({offsets})",
        offsets,
        texts,
        f"N({texts})",
        f"SUMPRODUCT(N({texts}))",
        f'SUMPRODUCT(COUNTIF({texts},">0"))',
    ):
        out.append((wrapped, False))
    for template in KEYED:
        for k in KEYS:
            x = template.format(k=k)
            out += [(x, False), (f"SUM({x})", False), (f"SUMPRODUCT({x})", False)]
    out += [
        ("SUMPRODUCT(1/COUNTIF(Z1:Z4,Z1:Z4))", False),
        ("SUM(1/COUNTIF(Z1:Z4,Z1:Z4))", False),
        ("1/COUNTIF(Z1:Z4,Z1:Z4)", False),
        ("COUNTIFS(Z1:Z4,{1;2},AA1:AA4,{10;20})", False),
        ("COUNTIFS(Z1:Z4,{1;2},AA1:AA4,{10,20})", False),
        ("SUMIFS(AA1:AA4,Z1:Z4,{1;2},AA1:AA4,{10;20})", False),
        ("SUBTOTAL({9;1},Z1:Z4)", False),
    ]
    for control in (
        "ISNUMBER(Z1:Z4)",
        "RANK(Z1:Z4,Z1:Z4)",
        "ADDRESS(Z1:Z4,1)",
        "ROWS(Z1:Z4)",
        "COLUMNS(Z1:Z4)",
        "INDEX(Z1:Z4,{1;2})",
    ):
        out.append((control, False))
    return list(dict.fromkeys(out + _newer_functions()))


def _newer_functions() -> list[tuple[str, bool]]:
    """The functions not in the first set of cases, each with a range or array where it may lift."""
    x = "_xlfn."
    rows = []
    for k in KEYS:
        rows += [
            f"AVERAGEIF(Z1:Z4,{k})",
            f"AVERAGEIF(Z1:Z4,{k},AA1:AA4)",
            f"AVERAGEIFS(AA1:AA4,Z1:Z4,{k})",
            f"{x}MAXIFS(AA1:AA4,Z1:Z4,{k})",
            f"{x}MINIFS(AA1:AA4,Z1:Z4,{k})",
            f"{x}XLOOKUP({k},Z1:Z4,AA1:AA4)",
            f"{x}XMATCH({k},Z1:Z4)",
            f"LARGE(Z1:Z4,{k})",
            f"SMALL(Z1:Z4,{k})",
            f"{x}SWITCH({k},1,10,2,20,0)",
            f"{x}IFS({k}>1,1,TRUE,0)",
            f'{x}TEXTJOIN(",",TRUE,{k})',
            f'{x}TEXTJOIN({k},TRUE,"a","b","c")',
            f"{x}CONCAT({k})",
        ]
    for p in ("0.5", "{0.25;0.75}", "{0.25,0.75}", "Z1:Z2/10"):
        rows += [f"PERCENTILE(Z1:Z4,{p})", f"{x}PERCENTILE.INC(Z1:Z4,{p})"]
    for q in ("1", "{1;3}", "Z1:Z2"):
        rows += [f"QUARTILE(Z1:Z4,{q})", f"{x}QUARTILE.INC(Z1:Z4,{q})"]
    rows += [
        "MODE(Z1:Z4)",
        f"{x}MODE.SNGL(Z1:Z4)",
        "MODE(Z1:Z4*1)",
        "LARGE(Z1:Z4*2,1)",
        "SMALL(Z1:Z4*2,1)",
        "PERCENTILE(Z1:Z4*2,0.5)",
        "QUARTILE(Z1:Z4*2,1)",
        f"{x}XLOOKUP(2,Z1:Z4,AA1:AB4)",
        f"{x}XLOOKUP(2,Z1:Z4,AA1:AA4)*2",
        f"SUM({x}XLOOKUP(2,Z1:Z4,AA1:AB4))",
        f"SUMPRODUCT({x}XLOOKUP(Z1:Z4,Z1:Z4,AA1:AA4))",
        f"{x}XLOOKUP(2,Z1:Z4*1,AA1:AA4)",
        f"{x}XLOOKUP(2,Z1:Z4,AA1:AA4*2)",
        f"{x}XMATCH(2,Z1:Z4*1)",
        f"{x}IFS(Z1:Z4>1,AA1:AA4,TRUE,0)",
        f"{x}SWITCH(Z1:Z4,1,AA1:AA4,0)",
        f"{x}CONCAT(Z1:Z4*2)",
        f'{x}TEXTJOIN(",",TRUE,Z1:Z4*2)',
        "SUM(LARGE(Z1:Z4,{1,2}))",
        "SUMPRODUCT(LARGE(Z1:Z4,ROW(Z1:Z2)))",
        f"SUMPRODUCT({x}XMATCH(Z1:Z4,Z1:Z4))",
        f"SUM({x}XMATCH(Z1:Z4,Z1:Z4))",
        f"SUM({x}XLOOKUP(Z1:Z4,Z1:Z4,AA1:AA4))",
    ]
    return [(r, False) for r in rows]


def layout() -> list[tuple[str, str, str]]:
    """(kind, cell or range, formula) for every cell of the workbook."""
    cells: list[tuple[str, str, str]] = []
    array_row = 20
    dynamic_row = 20
    for k, (formula, spills) in enumerate(formulas()):
        cells.append(("legacy", f"A{20 + k}", formula))
        cells.append(("legacy_row3", f"{get_column_letter(40 + k)}3", formula))
        cells.append(("legacy_column", f"AA{20 + k}", formula))
        cells.append(("cse", f"C{array_row}:E{array_row + 3}", formula))
        array_row += 5
        if not spills:
            cells.append(("dynamic", f"H{dynamic_row}", formula))
            dynamic_row += 8
    x = "ROW(Z1:Z4)"
    for rows, cols in ((2, 1), (6, 1), (4, 2), (1, 1)):
        last = f"{get_column_letter(2 + cols)}{array_row + rows - 1}"
        cells.append(("cse", f"C{array_row}:{last}", x))
        array_row += rows + 1
    return cells


def _data(ws: object) -> None:
    for column, values in DATA.items():
        for r, value in enumerate(values, 1):
            ws[f"{column}{r}"] = value  # type: ignore[index]


def build(path: Path) -> None:
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "probe"
    _data(ws)
    _data(wb.create_sheet("other"))
    wb.defined_names["Block"] = DefinedName("Block", attr_text="probe!$Z$1:$Z$4")
    dynamic = []
    for kind, where, formula in layout():
        if kind == "cse":
            ws[where.split(":")[0]] = ArrayFormula(where, "=" + formula)
        elif kind == "dynamic":
            ws[where] = "=0"  # a placeholder cell, replaced by add_dynamic
            dynamic.append((kind, where, formula))
        else:
            ws[where] = "=" + formula
    wb.save(path)
    add_dynamic(path, dynamic)


def main() -> None:
    CONTAINER.mkdir(parents=True, exist_ok=True)
    src = CONTAINER / "xlrecalc_array_refs.xlsx"
    copy = CONTAINER / "xlrecalc_array_refs_saved.xlsx"
    for p in (src, copy):
        p.unlink(missing_ok=True)
    build(src)
    version = excel_version()
    result = subprocess.run(
        ["osascript", str(HERE / "save_copy.applescript"), str(src), str(copy)],
        capture_output=True,
        text=True,
    )
    if result.returncode:
        sys.exit(f"Excel did not open the workbook: {result.stderr.strip()}")
    header = [
        "# suite: array_refs",
        f"# excel: {version} for Mac, locale en_US",
        edition_line(),
        f"# captured: {dt.date.today().isoformat()}",
        "# generator: capture/array_ref_probe.py",
        "# method: plain, legacy array and dynamic-array formulas over a data block; Excel",
        "#   calculates and saves a copy; every cell of the copy is recorded",
        "# input rows: input, kind, cell or range, formula (file form)",
        "# cell rows: cell, address, type, value (repr), formula (repr), f_t, f_ref, cm, vm",
    ]
    body = ["\t".join(["input", kind, where, repr(f)]) for kind, where, f in layout()]
    with Book(copy) as book:
        for c in book.cells("probe"):
            f = c.formula
            value = number_text(c.value) if isinstance(c.value, float) else c.value
            address = f"{get_column_letter(c.column)}{c.row}"
            fields = [repr(value), repr(f.text if f else None), f.kind if f else ""]
            fields += [(f.ref or "") if f else "", c.cm or "", c.vm or ""]
            body.append("\t".join(["cell", address, c.type, *fields]))
    (CAPTURES / "array_refs.tsv").write_text("\n".join([*header, *body]) + "\n", encoding="utf-8")
    print(f"{len(body)} rows -> array_refs.tsv")
    for p in (src, copy):
        p.unlink(missing_ok=True)


if __name__ == "__main__":
    main()
