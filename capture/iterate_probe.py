"""How Excel iterates a circular workbook: order, count, stopping test (ARCHITECTURE.md §7.2).

    uv run python -m capture.iterate_probe

Non-interactive. Each variant is one small workbook with iterative calculation on, its cells'
cached values written into the XML (the values iteration starts from). Each is opened alone, with
alerts off: what Excel shows after opening (one run of iterations) and after one ``calculate``
(another run) is read (``iterate_read.applescript``), then it is opened again and a copy saved
(``save_copy.applescript``, which runs once more on saving), and the copy's exact cached values are
read. Writes ``tests/captures/iterate_calc.tsv``.

Most values are integers, so what Excel shows is exact; the saved copy gives the exact double.

Iteration is an application setting that Excel takes from the first workbook it opens, so this
needs Excel with no workbook open, and refuses to run otherwise. Each Excel step has a time limit;
when one passes, the probe stops and leaves Excel as it is.
"""

from __future__ import annotations

import datetime as dt
import re
import sys
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET

import openpyxl

from capture.edition import edition_line
from capture.excel_watch import osascript_or_exit
from capture.pipeline import CAPTURES, CONTAINER, excel_version

HERE = Path(__file__).resolve().parent
US = "\x1f"
STEP_LIMIT = 90  # seconds
NS = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
REL = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id"

Cell = tuple[str, str, str | None]  # (cell, formula, cached value or None)
Sheets = list[tuple[str, list[Cell]]]


def one(*cells: Cell) -> Sheets:
    return [("s", list(cells))]


HALF = ("A1", "A1/2+1", "0")  # 1, 1.5, 1.75, ...: the change halves each iteration
COUNTER = ("A1", "A1+1", "0")

# (name, calcPr attributes after iterate="1", sheets)
VARIANTS: list[tuple[str, str, Sheets]] = [
    # How many iterations a run makes.
    *[(f"count_{n}", f'iterateCount="{n}"', one(COUNTER)) for n in (1, 2, 7, 1000)],
    # Order within a sheet: which of two cells is calculated first, each reading the other.
    ("rows", 'iterateCount="1"', one(("A1", "A5+1", "0"), ("A5", "A1+10", "0"))),
    ("rows_reversed", 'iterateCount="1"', one(("A1", "A5+10", "0"), ("A5", "A1+1", "0"))),
    ("columns", 'iterateCount="1"', one(("A2", "B1+1", "0"), ("B1", "A2+10", "0"))),
    # A later cell sees an earlier one's new value in the same iteration, or the old one.
    (
        "three",
        'iterateCount="1"',
        one(("A1", "A3+1", "0"), ("A2", "A1+1", "0"), ("A3", "A2+1", "0")),
    ),
    (
        "three_backward",
        'iterateCount="1"',
        one(("A1", "A2+1", "0"), ("A2", "A3+1", "0"), ("A3", "A1+1", "0")),
    ),
    (
        "three_count_2",
        'iterateCount="2"',
        one(("A1", "A2+1", "0"), ("A2", "A3+1", "0"), ("A3", "A1+1", "0")),
    ),
    # Across sheets, in tab order and reversed.
    (
        "sheets",
        'iterateCount="1"',
        [("one", [("A1", "two!A1+1", "0")]), ("two", [("A1", "one!A1+10", "0")])],
    ),
    (
        "sheets_reversed",
        'iterateCount="1"',
        [("two", [("A1", "one!A1+10", "0")]), ("one", [("A1", "two!A1+1", "0")])],
    ),
    # Other names, in each tab order and with the reading turned around, and three sheets.
    (
        "names_ba",
        'iterateCount="1"',
        [("b", [("A1", "a!A1+1", "0")]), ("a", [("A1", "b!A1+10", "0")])],
    ),
    (
        "names_ab",
        'iterateCount="1"',
        [("a", [("A1", "b!A1+10", "0")]), ("b", [("A1", "a!A1+1", "0")])],
    ),
    (
        "names_ba_turned",
        'iterateCount="1"',
        [("b", [("A1", "a!A1+10", "0")]), ("a", [("A1", "b!A1+1", "0")])],
    ),
    (
        "names_cab",
        'iterateCount="1"',
        [
            ("c", [("A1", "b!A1+1", "0")]),
            ("a", [("A1", "c!A1+10", "0")]),
            ("b", [("A1", "a!A1+100", "0")]),
        ],
    ),
    # A precedent outside the loop, below it: calculated before the loop, or read stale.
    ("precedent_after", 'iterateCount="1"', one(("A1", "A1+A9", "0"), ("A9", "2+3", "0"))),
    ("precedent_chain", 'iterateCount="2"', one(("A1", "A1+A9", "0"), ("A9", "A1*0+7", "0"))),
    # A dependent outside the loop.
    ("dependent", 'iterateCount="3"', one(COUNTER, ("B1", "A1*2", "999"))),
    # The stopping test: maximum change against iterateDelta.
    *[
        (f"delta_{d}", f'iterateCount="100" iterateDelta="{d}"', one(HALF))
        for d in ("0.3", "0.25", "0.2", "0.001", "0")
    ],
    ("delta_default", 'iterateCount="100"', one(HALF)),
    ("delta_0_count_5", 'iterateCount="5" iterateDelta="0"', one(HALF)),
    ("large", 'iterateCount="100" iterateDelta="0.001"', one(("A1", "A1/2+1E6", "0"))),
    ("small", 'iterateCount="100" iterateDelta="0.001"', one(("A1", "A1/2+1E-6", "0"))),
    ("negative", 'iterateCount="100" iterateDelta="0.3"', one(("A1", "A1/2-1", "0"))),
    # Two loops: one stopping test for all, or one each.
    (
        "two_loops",
        'iterateCount="100" iterateDelta="0.3"',
        one(HALF, ("A3", "A3+1", "0")),
    ),
    (
        "two_loops_reversed",
        'iterateCount="100" iterateDelta="0.3"',
        one(("A1", "A1+1", "0"), ("A3", "A3/2+1", "0")),
    ),
    (
        "two_converging",
        'iterateCount="100" iterateDelta="0.3"',
        one(HALF, ("A3", "A3/4+3", "0")),
    ),
    # A loop of two cells converging: the change of either.
    (
        "pair",
        'iterateCount="100" iterateDelta="0.3"',
        one(("A1", "A2/2+1", "0"), ("A2", "A1", "0")),
    ),
    # Not converging.
    ("oscillate_100", 'iterateCount="100"', one(("A1", "1-A1", "0"))),
    ("oscillate_7", 'iterateCount="7"', one(("A1", "1-A1", "0"))),
    ("overflow", 'iterateCount="400"', one(("A1", "A1*10+1", "0"))),
    # Already converged when opened.
    ("converged", 'iterateCount="100" iterateDelta="0.3"', one(("A1", "A1/2+1", "2"))),
    # Round 2. A dependent outside the loop: does its change count, is it recalculated each time.
    (
        "dependent_change",
        'iterateCount="100" iterateDelta="0.3"',
        one(HALF, ("B1", "A1*1000", "0")),
    ),
    (
        "dependent_read",
        'iterateCount="3"',
        one(COUNTER, ("B1", "A1*2", "0"), ("C1", "C1+B1", "0")),
    ),
    # A change between kinds, with a delta no number reaches: a number alone stops at once.
    ("kind_number", 'iterateCount="7" iterateDelta="1E9"', one(("A1", "1-A1", "0"))),
    ("kind_text", 'iterateCount="7" iterateDelta="1E9"', one(("A1", 'IF(A1=0,"x",0)', "0"))),
    (
        "kind_error",
        'iterateCount="7" iterateDelta="1E9"',
        one(("A1", "IF(ISERROR(A1),0,1/0)", "0")),
    ),
    ("kind_logical", 'iterateCount="7" iterateDelta="1E9"', one(("A1", "IF(A1=0,TRUE,0)", "0"))),
    # Sheet names in case, digits and beyond ASCII.
    *[
        (
            f"names_{label}",
            'iterateCount="1"',
            [(x, [("A1", f"'{y}'!A1+1", "0")]), (y, [("A1", f"'{x}'!A1+10", "0")])],
        )
        for label, x, y in (("case", "B", "a"), ("digits", "10", "9"), ("accent", "f", "é"))
    ],
    # Loops found only through INDIRECT and OFFSET.
    ("indirect", 'iterateCount="1"', one(("A1", 'INDIRECT("A5")+1', "0"), ("A5", "A1+10", "0"))),
    ("offset", 'iterateCount="1"', one(("A1", "OFFSET(A5,0,0)+1", "0"), ("A5", "A1+10", "0"))),
    # Round 3. A dependent above its loop: calculated in row order with it, a step behind, or after.
    ("dependent_before", 'iterateCount="3"', one(("A1", "A5*2", "0"), ("A5", "A5+1", "0"))),
]


# (name, cell to edit, number entered, sheets): iterateCount 3, a run on opening.
EDITS: list[tuple[str, str, str, Sheets]] = [
    ("edit_read", "s!B9", "2", one(("A1", "A1+B9", "0"), ("B9", "", "1"))),
    ("edit_unread", "s!C9", "2", one(("A1", "A1+B9", "0"), ("B9", "", "1"), ("C9", "", "1"))),
]


def build(path: Path, calc: str, sheets: Sheets) -> None:
    """The workbook: formulas, their cached values written into the XML, and ``calc``."""
    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    for name, cells in sheets:
        ws = wb.create_sheet(name)
        for ref, formula, cached in cells:
            ws[ref] = "=" + formula if formula else float(cached)  # "" is a constant
    plain = path.with_name("xlrecalc_iterate_probe_plain.xlsx")
    wb.save(plain)
    with zipfile.ZipFile(plain) as zin, zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            found = re.fullmatch(r"xl/worksheets/sheet(\d+)\.xml", item.filename)
            if found:
                text = data.decode()
                for ref, formula, cached in sheets[int(found[1]) - 1][1]:
                    if not formula or cached is None:  # None: no cached value, as openpyxl writes
                        continue
                    text, n = re.subn(
                        rf'(<c r="{ref}"[^>]*>\s*<f>[^<]*</f>)\s*<v\s*/>',
                        rf"\g<1><v>{cached}</v>",
                        text,
                    )
                    if n != 1:
                        raise RuntimeError(f"{ref}: not found")
                data = text.encode()
            if item.filename == "xl/workbook.xml":
                text, n = re.subn(r"<calcPr[^>]*/>", calc, data.decode())
                if n != 1:
                    raise RuntimeError("calcPr not found")
                data = text.encode()
            zout.writestr(item, data)
    plain.unlink()


def osascript(*args: str) -> str:
    result = osascript_or_exit(*args, seconds=STEP_LIMIT)
    if result.returncode:
        sys.exit(f"osascript failed: {result.stderr.strip()}")
    return result.stdout.rstrip("\n")


def saved_values(path: Path) -> dict[str, tuple[str, str]]:
    """Each formula cell's cached type and value, as ``sheet!A1``, from a copy Excel saved."""
    with zipfile.ZipFile(path) as z:
        book = ET.fromstring(z.read("xl/workbook.xml"))
        rels = ET.fromstring(z.read("xl/_rels/workbook.xml.rels"))
        targets = {r.get("Id"): r.get("Target", "") for r in rels}
        out: dict[str, tuple[str, str]] = {}
        for sheet in book.iter(f"{NS}sheet"):
            target = targets[sheet.get(REL)]
            part = target.lstrip("/") if target.startswith("/") else "xl/" + target
            for c in ET.fromstring(z.read(part)).iter(f"{NS}c"):
                if c.find(f"{NS}f") is None:
                    continue
                v = c.find(f"{NS}v")
                text = "" if v is None or v.text is None else v.text
                out[f"{sheet.get('name')}!{c.get('r')}"] = (c.get("t", "n"), text)
    return out


def main() -> None:
    count = osascript("-e", 'tell application "Microsoft Excel" to count of workbooks')
    if count != "0":
        sys.exit(f"Excel has {count} workbooks open: iteration is taken from the first one opened")
    CONTAINER.mkdir(parents=True, exist_ok=True)
    header = [
        "# suite: iterate_calc",
        f"# excel: {excel_version()} for Mac, locale en_US",
        edition_line(),
        f"# captured: {dt.date.today().isoformat()}",
        "# generator: capture/iterate_probe.py",
        "# method: a workbook per variant, cached values written into its XML, opened alone with",
        "#   alerts off; shown is what Excel shows after opening (value, displayed text),",
        "#   calculated the same after one calculate; saved_type and saved the cached value of a",
        "#   copy saved after opening again (opening and saving each run the iterations)",
        "# columns: variant calcPr cell formula cached shown shown_text calculated "
        "calculated_text saved_type saved",
    ]
    body: list[str] = []
    for name, attrs, sheets in VARIANTS:
        calc = f'<calcPr calcId="191029" iterate="1" {attrs}/>'
        path = CONTAINER / f"xlrecalc_iterate_probe_{name}.xlsx"
        copy = path.with_name(path.stem + "_saved.xlsx")
        copy.unlink(missing_ok=True)
        build(path, calc, sheets)
        refs = [f"{sheet}!{ref}" for sheet, cells in sheets for ref, _, _ in cells]
        fields = osascript(str(HERE / "iterate_read.applescript"), str(path), *refs).split(US)
        if fields[0] != "open":
            sys.exit(f"{path.name} did not open: {fields}")
        k = len(refs)
        first, second = fields[1 : 1 + 2 * k], fields[2 + 2 * k :]
        osascript(str(HERE / "save_copy.applescript"), str(path), str(copy))
        saved = saved_values(copy)
        cells = [(f"{sheet}!{ref}", f, c) for sheet, cs in sheets for ref, f, c in cs]
        for i, (ref, formula, cached) in enumerate(cells):
            kind, value = saved.get(ref, ("none", ""))
            row = [name, attrs, ref, formula, cached, first[2 * i], first[2 * i + 1]]
            body.append("\t".join([*row, second[2 * i], second[2 * i + 1], kind, value]))
            print(
                f"{name:18} {ref:8} {formula:12} shown {first[2 * i]:>22} "
                f"calc {second[2 * i]:>22} saved {value}",
                flush=True,
            )
        path.unlink()
        copy.unlink()
    # An edit: a constant the loop reads, and a cell nothing reads.
    for name, target, number, sheets in EDITS:
        path = CONTAINER / f"xlrecalc_iterate_probe_{name}.xlsx"
        build(path, '<calcPr calcId="191029" iterate="1" iterateCount="3"/>', sheets)
        refs = [f"{sheet}!{ref}" for sheet, cells in sheets for ref, _, _ in cells]
        fields = osascript(
            str(HERE / "iterate_edit.applescript"), str(path), target, number, *refs
        ).split(US)
        if fields[0] != "open":
            sys.exit(f"{path.name} did not open: {fields}")
        before, after = fields[1 : 1 + len(refs)], fields[2 + len(refs) :]
        cells = [(f"{sheet}!{ref}", f, c) for sheet, cs in sheets for ref, f, c in cs]
        for i, (ref, formula, cached) in enumerate(cells):
            row = [name, f"edit {target}={number}", ref, formula, cached, before[i], ""]
            body.append("\t".join([*row, after[i], "", "none", ""]))
            print(
                f"{name:18} {ref:8} {formula:12} before {before[i]:>8} after {after[i]}", flush=True
            )
        path.unlink()
    out = CAPTURES / "iterate_calc.tsv"
    out.write_text("\n".join([*header, *body]) + "\n", encoding="utf-8")
    print(f"{len(VARIANTS)} variants -> {out.name}")


if __name__ == "__main__":
    main()
