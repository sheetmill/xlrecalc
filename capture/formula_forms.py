"""Probe the file form of formulas: what Excel accepts in ``<f>`` and what it writes back.

    uv run python capture/formula_forms.py

Writes the formula corpus (``tests/formula_corpus.py``; external references and hostile
sheet names left out, since they need parts the workbook does not have) into a workbook with
the sheets, defined names and table the formulas refer to, one formula per cell with a
constant between formulas so Excel does not join them into shared formulas. Excel opens the
workbook and saves a copy (``save_copy.applescript``); the copy's ``<f>`` text is compared with
what was written. A workbook Excel refuses (error -50, a repair prompt with alerts off) is
split in halves until each refused formula is found alone. Writes
``tests/captures/formula_forms.tsv``: one row per formula, the text written and the text Excel
saved (Python ``repr``; ``'REFUSED'`` when Excel refuses a workbook holding it alone).
"""

from __future__ import annotations

import subprocess
import sys
import xml.etree.ElementTree as ET
import zipfile
from pathlib import Path

import openpyxl
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.worksheet.table import Table

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))

from tests.formula_corpus import CORPUS  # noqa: E402

HERE = Path(__file__).resolve().parent
OUTPUT = ROOT / "tests" / "captures" / "formula_forms.tsv"
BUILD = HERE / "build"
CONTAINER = Path.home() / "Library/Containers/com.microsoft.Excel/Data/Documents/xlrecalc"
NS = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
SHEETS = [
    "Sheet2", "Sheet3", "My Sheet", "It's", "2024", "Données", "Sheet 1", "Sheet 3",
    "Q1-Q2", "R&D", "Cash+Debt", "Plan (v2)", "v1.2", "a,b", "3-15-24", "2024 Budget", "Résumé",
    "Ops '25", "East", "West", "Jan-24", "Mar-24", "North (1)", "South (9)", "10 (a)", "20 (b)",
    "12 - Plan Notes ",
]  # fmt: skip
# External references need external-link parts this workbook does not have.
SKIP = ("[1]", "[2]", "[3]", "[4]", "[5]", "__import__", "); import", "K[0]")
COLUMN = 11  # K: away from the cells the formulas refer to


def corpus() -> list[str]:
    return [f for f in CORPUS if not any(s in f for s in SKIP)]


def build(path: Path, formulas: list[str]) -> None:
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Sheet1"
    for name in SHEETS:
        wb.create_sheet(name)
    data = wb["Sheet2"]
    for row in (
        ["Price", "Qty", "Unit Price", "'#Hash", "Col[1]", "__class__"],
        [1, 2, 3, 4, 5, 6],
    ):
        data.append(row)
    data.add_table(Table(displayName="Table1", ref="A1:F2"))
    for name, ref in [
        ("Rate", "Sheet1!$A$1"),
        ("_Rate", "Sheet1!$A$2"),
        ("my.name", "Sheet1!$A$3"),
        ("is_ok?", "Sheet1!$A$4"),
        ("\\backslash", "Sheet1!$A$5"),
    ]:
        wb.defined_names[name] = DefinedName(name, attr_text=ref)
    data.defined_names["Rate"] = DefinedName("Rate", attr_text="Sheet2!$A$2")
    for k, formula in enumerate(formulas):
        ws.cell(2 * k + 1, COLUMN, "=" + formula)
        ws.cell(2 * k + 2, COLUMN, 0)
    wb.save(path)


def saved_formulas(path: Path) -> list[str | None]:
    with zipfile.ZipFile(path) as z:
        root = ET.fromstring(z.read("xl/worksheets/sheet1.xml"))
    by_row: dict[int, str | None] = {}
    for c in root.iter(f"{NS}c"):
        f = c.find(f"{NS}f")
        if f is not None:
            row = int("".join(ch for ch in c.get("r", "") if ch.isdigit()))
            by_row[row] = f.text if f.text is not None else f"<shared si={f.get('si')}>"
    return [by_row.get(2 * k + 1) for k in range(len(by_row) and max(by_row) // 2 + 1)]


def excel_saves(formulas: list[str], tag: str) -> list[str | None] | None:
    """Excel's text for each formula, or None when Excel refuses the workbook."""
    src = CONTAINER / f"xlrecalc_forms_{tag}.xlsx"
    copy = CONTAINER / f"xlrecalc_forms_{tag}_saved.xlsx"
    for p in (src, copy):
        p.unlink(missing_ok=True)
    build(src, formulas)
    result = subprocess.run(
        ["osascript", str(HERE / "save_copy.applescript"), str(src), str(copy)],
        capture_output=True,
        text=True,
    )
    if result.returncode:
        if "(-50)" in result.stderr:
            return None
        raise RuntimeError(result.stderr)
    saved = saved_formulas(copy)
    return (saved + [None] * len(formulas))[: len(formulas)]


def probe(formulas: list[str], tag: str, out: dict[str, str | None]) -> None:
    saved = excel_saves(formulas, tag)
    if saved is not None:
        out.update(zip(formulas, saved, strict=True))
        return
    if len(formulas) == 1:
        out[formulas[0]] = "REFUSED"
        return
    half = len(formulas) // 2
    probe(formulas[:half], tag + "a", out)
    probe(formulas[half:], tag + "b", out)


def excel_version() -> str:
    result = subprocess.run(
        ["osascript", "-e", 'tell application "Microsoft Excel" to version'],
        capture_output=True,
        text=True,
    )
    return result.stdout.strip()


def write(path: Path, rows: list[tuple[str, str | None]], version: str) -> None:
    header = [
        "# suite: formula_forms",
        f"# excel: {version} for Mac, locale en_US",
        "# generator: capture/formula_forms.py (formulas from tests/formula_corpus.py)",
        "# method: openpyxl writes each formula into <f>; Excel opens the workbook and saves a",
        "#   copy; the column excel is the <f> text of the copy, or REFUSED when Excel refuses",
        "#   to open a workbook holding that formula alone (error -50 with alerts off)",
        "# columns: written, excel (both as Python repr)",
    ]
    body = [f"{w!r}\t{e!r}" for w, e in rows]
    path.write_text("\n".join([*header, *body]) + "\n", encoding="utf-8")


def read(path: Path) -> list[tuple[str, str | None]]:
    import ast

    rows = []
    for line in path.read_text(encoding="utf-8").splitlines():
        if line.startswith("#"):
            continue
        written, excel = line.split("\t")
        rows.append((ast.literal_eval(written), ast.literal_eval(excel)))
    return rows


def main() -> None:
    CONTAINER.mkdir(parents=True, exist_ok=True)
    BUILD.mkdir(exist_ok=True)
    formulas = corpus()
    out: dict[str, str | None] = {}
    probe(formulas, "x", out)
    lines = []
    same = changed = refused = 0
    for f in formulas:
        got = out[f]
        if got == f:
            same += 1
            continue
        if got == "REFUSED":
            refused += 1
        else:
            changed += 1
        lines.append(f"{f!r}\t{got!r}")
    summary = f"{len(formulas)} written: {same} kept, {changed} rewritten, {refused} refused"
    print(summary)
    print("\n".join(lines))
    write(OUTPUT, [(f, out[f]) for f in formulas], excel_version())
    for p in CONTAINER.glob("xlrecalc_forms_*.xlsx"):
        p.unlink()


if __name__ == "__main__":
    main()
