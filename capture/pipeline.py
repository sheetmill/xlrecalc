"""Capture formula results from Excel: build a workbook, have Excel compute and save it, read back.

    uv run python -m capture.pipeline SUITE [SUITE ...]     # or: all

A suite (``capture/suites.py``) is a list of ``Case``: a formula with placeholders ``{1}``,
``{2}``, ... for its input cells, and the input values. Each case is one row of sheet
``cases``: column A holds the formula, columns B onward the inputs, and ``{k}`` becomes the
reference to the k-th input of the same row. Floats are written as their ``repr`` text, never
openpyxl's ``%.16g``; texts as text cells (whitespace-only text as a formula, since openpyxl
writes it without ``xml:space="preserve"`` and Excel would read it as empty); ``None`` leaves
the cell empty; an ``ErrorValue`` is written as an error cell; a ``FormulaInput`` as that
formula (``EMPTY_TEXT``, ``=""``, is the only way to get an empty text: openpyxl does not write
an empty string).

Excel opens the workbook (openpyxl's files make it calculate everything on open) and saves a
copy (``save_copy.applescript``); the cached value and type of each formula cell are read from
the copy's XML, which gives the exact double and the real error code. Writes
``tests/captures/<suite>.tsv``. Workbooks live in Excel's container; nothing else is touched.
"""

from __future__ import annotations

import datetime as dt
import math
import os
import re
import subprocess
import sys
import xml.etree.ElementTree as ET
import zipfile
from dataclasses import dataclass
from pathlib import Path

import openpyxl
from openpyxl.utils import get_column_letter

from xlrecalc._values import ErrorValue

ROOT = Path(__file__).resolve().parents[1]
HERE = Path(__file__).resolve().parent
CAPTURES = ROOT / "tests" / "captures"
CONTAINER = Path.home() / "Library/Containers/com.microsoft.Excel/Data/Documents/xlrecalc"
NS = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
COLUMNS = ("suite", "tag", "case", "formula", "inputs", "expected_type", "expected_repr")


@dataclass(frozen=True)
class FormulaInput:
    """An input cell holding a formula, such as ``=""`` for a true empty string."""

    text: str

    def __repr__(self) -> str:
        return f"FormulaInput({self.text!r})"


EMPTY_TEXT = FormulaInput('=""')

Input = float | str | bool | ErrorValue | FormulaInput | None


@dataclass(frozen=True)
class Case:
    """One formula and its inputs. ``tag`` groups cases for reports (a function or a topic)."""

    tag: str
    formula: str
    inputs: tuple[Input, ...] = ()


@dataclass(frozen=True)
class Result:
    """What Excel saved for a formula cell: ``n``, ``str``, ``b`` or ``e``, and the value."""

    kind: str
    value: float | str | bool | ErrorValue


def formula_text(case: Case, row: int) -> str:
    """The case's formula with each ``{k}`` replaced by its input cell in ``row``."""
    return re.sub(r"\{(\d+)\}", lambda m: f"{get_column_letter(1 + int(m[1]))}{row}", case.formula)


def build(
    path: Path, cases: list[Case], date1904: bool = False, fixed: dict[str, Input] | None = None
) -> None:
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "cases"
    if date1904:
        wb.epoch = openpyxl.utils.datetime.CALENDAR_MAC_1904
    for ref, value in (fixed or {}).items():
        _set_input(ws[ref], value)
    for row, case in enumerate(cases, 1):
        ws.cell(row, 1, "=" + formula_text(case, row))
        for k, value in enumerate(case.inputs, 2):
            _set_input(ws.cell(row, k), value)
    wb.save(path)


def _set_input(cell: object, value: Input) -> None:
    if value is None:
        return
    if isinstance(value, FormulaInput):
        cell.value = value.text  # type: ignore[attr-defined]
        return
    if value == "":
        cell.value = EMPTY_TEXT.text  # type: ignore[attr-defined]  # openpyxl drops an empty text
        return
    if isinstance(value, str) and not value.strip():
        # openpyxl writes whitespace-only text without xml:space="preserve", and Excel reads
        # it as empty text; a formula keeps it.
        cell.value = '="' + value.replace('"', '""') + '"'  # type: ignore[attr-defined]
        return
    if isinstance(value, ErrorValue):
        cell.value = value.code  # type: ignore[attr-defined]
        cell.data_type = "e"  # type: ignore[attr-defined]
    elif isinstance(value, bool):
        cell.value = value  # type: ignore[attr-defined]
    elif isinstance(value, float):
        if not math.isfinite(value):
            raise ValueError(f"Excel has no {value!r}")
        text = repr(value)
        if float(text) != value:
            raise AssertionError(text)
        cell.value = text  # type: ignore[attr-defined]
        cell.data_type = "n"  # the writer emits the text as it is: Excel parses exactly this
    else:
        cell.value = value  # type: ignore[attr-defined]
        cell.data_type = "s"  # type: ignore[attr-defined]  # never a formula


def read_results(path: Path, count: int) -> list[Result | None]:
    """Column A's cached values, rows 1 to ``count``, from a workbook Excel saved."""
    with zipfile.ZipFile(path) as z:
        sheet = ET.fromstring(z.read("xl/worksheets/sheet1.xml"))
    out: list[Result | None] = [None] * count
    for c in sheet.iter(f"{NS}c"):
        ref = c.get("r", "")
        if not re.fullmatch(r"A\d+", ref):
            continue
        row = int(ref[1:])
        v = c.find(f"{NS}v")
        text = "" if v is None or v.text is None else v.text
        kind = c.get("t", "n")
        if row > count or c.find(f"{NS}f") is None:
            continue
        if kind == "n":
            out[row - 1] = Result("n", float(text))
        elif kind == "str":
            out[row - 1] = Result("str", text)
        elif kind == "b":
            out[row - 1] = Result("b", text == "1")
        elif kind == "e":
            out[row - 1] = Result("e", ErrorValue(text))
        else:
            raise ValueError(f"{ref}: unexpected type {kind!r}")
    return out


def run_excel(
    cases: list[Case], name: str, date1904: bool = False, fixed: dict[str, Input] | None = None
) -> list[Result | None]:
    CONTAINER.mkdir(parents=True, exist_ok=True)
    src = CONTAINER / f"xlrecalc_capture_{name}.xlsx"
    copy = CONTAINER / f"xlrecalc_capture_{name}_saved.xlsx"
    for p in (src, copy):
        p.unlink(missing_ok=True)
    build(src, cases, date1904, fixed)
    result = subprocess.run(
        ["osascript", str(HERE / "save_copy.applescript"), str(src), str(copy)],
        capture_output=True,
        text=True,
    )
    if result.returncode:
        raise RuntimeError(result.stderr.strip())
    results = read_results(copy, len(cases))
    src.unlink()
    if not os.environ.get("XLRECALC_CAPTURE_KEEP"):  # keep Excel's copy to read its formulas
        copy.unlink()
    return results


def excel_version() -> str:
    result = subprocess.run(
        ["osascript", "-e", 'tell application "Microsoft Excel" to version'],
        capture_output=True,
        text=True,
    )
    return result.stdout.strip()


def result_repr(result: Result | None) -> tuple[str, str]:
    if result is None:
        return "none", ""
    value = result.value
    if isinstance(value, ErrorValue):
        return "e", value.code
    return result.kind, repr(value)


def write_suite(
    name: str,
    generator: str,
    cases: list[Case],
    results: list[Result | None],
    date1904: bool,
    version: str,
    fixed: dict[str, Input] | None = None,
) -> Path:
    path = CAPTURES / f"{name}.tsv"
    header = [
        f"# suite: {name}",
        f"# excel: {version} for Mac, locale en_US",
        f"# date_system: {'1904' if date1904 else '1900'}",
        f"# captured: {dt.date.today().isoformat()}",
        f"# generator: {generator}",
        "# method: openpyxl writes the formula in column A and the inputs in B onward (floats as",
        "#   repr text); Excel opens the workbook, calculates and saves a copy; the cached value",
        "#   and type of A are read from the copy's XML (n: number, str: text, b: logical,",
        "#   e: error); inputs and expected_repr are Python repr, with ErrorValue('#N/A') for an",
        "#   error input and FormulaInput('=\"\"') for an input cell holding that formula",
        "# columns: " + " ".join(COLUMNS),
    ]
    if fixed:
        header.insert(-1, f"# fixed: {fixed!r}")
    lines = []
    for k, (case, result) in enumerate(zip(cases, results, strict=True), 1):
        kind, text = result_repr(result)
        inputs = repr(case.inputs)
        lines.append("\t".join([name, case.tag, str(k), case.formula, inputs, kind, text]))
    path.write_text("\n".join([*header, *lines]) + "\n", encoding="utf-8")
    return path


def main(argv: list[str]) -> None:
    from capture import suites

    names = list(suites.SUITES) if argv == ["all"] else argv
    version = excel_version()
    for name in names:
        suite = suites.SUITES[name]
        cases = suite.cases()
        results = run_excel(cases, name, suite.date1904, suite.fixed)
        path = write_suite(
            name,
            f"capture/suites.py:{suite.cases.__name__}",
            cases,
            results,
            suite.date1904,
            version,
            suite.fixed,
        )
        missing = sum(r is None for r in results)
        print(
            f"{name}: {len(cases)} cases -> {path.relative_to(ROOT)}"
            + (f" ({missing} without a result)" if missing else "")
        )


if __name__ == "__main__":
    # Run the imported module, not __main__: the suites import capture.pipeline, and its classes
    # must be the ones the checks above see.
    from capture import pipeline

    pipeline.main(sys.argv[1:])
