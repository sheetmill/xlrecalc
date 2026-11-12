"""Whether formulas pasted as text enter as a person typing them does (CAPTURING.md §2).

    uv run python -m capture.paste_probe

AppleScript's ``formula`` enters only legacy formulas, so the captures of what Excel writes for a
typed formula (``marker_typed``, ``let_typed`` and the LAMBDA rows ``lambda_probe --typed`` holds)
were typed by a person. This builds each of those workbooks again, has Excel paste the formulas
as one column of text from the clipboard (``paste_formulas.applescript``; the clipboard's text is
restored after), saves a copy, and compares every formula's anchor cell with the typed one: the
formula's text as the file holds it (prefixes included), its ``t`` and ``ref``, the dynamic-array
marker (``cm``), and the cached value and type. Prints the matches and every difference.
"""

from __future__ import annotations

import subprocess
import tempfile
from pathlib import Path

import openpyxl
from openpyxl.workbook.defined_name import DefinedName

from capture import lambda_probe, marker_probe
from capture.excel_watch import osascript_or_exit
from capture.pipeline import CAPTURES, CONTAINER
from xlrecalc._xlsx import Book

HERE = Path(__file__).resolve().parent


def _anchors(path: Path, rows: list[int]) -> dict[int, tuple[object, ...]]:
    """Each row's A cell: file form, t, ref, cm, type, value."""
    with Book(path) as book:
        cells = {c.row: c for c in book.cells("probe") if c.column == 1}
    out: dict[int, tuple[object, ...]] = {}
    for r in rows:
        c = cells.get(r)
        f = c.formula if c else None
        out[r] = (
            f.text if f else None,
            f.kind if f else None,
            f.ref if f else None,
            c.cm if c else None,
            c.type if c else None,
            c.value if c else None,
        )
    return out


def _paste(build: Path, formulas: dict[int, str], copy: Path) -> None:
    """Paste the formulas, each at its row of column A, into the built workbook; save ``copy``."""
    lines = [""] * max(formulas)
    for row, text in formulas.items():
        lines[row - 1] = text
    saved = subprocess.run(["pbpaste"], capture_output=True, check=False).stdout
    with tempfile.NamedTemporaryFile("w", suffix=".txt", encoding="utf-8", delete=False) as f:
        f.write("\n".join(lines))
        listing = f.name
    try:
        osascript_or_exit(
            str(HERE / "paste_formulas.applescript"), str(build), listing, str(copy), seconds=900
        )
    finally:
        subprocess.run(["pbcopy"], input=saved, check=False)
        Path(listing).unlink(missing_ok=True)


def compare(name: str, build: Path, formulas: dict[int, str], truth: Path) -> list[str]:
    copy = CONTAINER / f"xlrecalc_paste_{name}_saved.xlsx"
    copy.unlink(missing_ok=True)
    try:
        _paste(build, formulas, copy)
        got, want = _anchors(copy, list(formulas)), _anchors(truth, list(formulas))
    finally:
        build.unlink(missing_ok=True)
        copy.unlink(missing_ok=True)
    differ = [
        f"{name}\tA{r}\t{formulas[r]}\ttyped {want[r]!r}\tpasted {got[r]!r}"
        for r in formulas
        if got[r] != want[r]
    ]
    print(f"{name}: {len(formulas) - len(differ)} of {len(formulas)} match", flush=True)
    return differ


def main() -> None:
    CONTAINER.mkdir(parents=True, exist_ok=True)
    differ = []
    build = CONTAINER / "xlrecalc_paste_marker.xlsx"
    marker_probe.typed_workbook(build)
    formulas = {4 * k + 1: f for k, f in enumerate(marker_probe.TYPED)}
    differ += compare("marker", build, formulas, CAPTURES / "marker_typed.xlsx")
    build = CONTAINER / "xlrecalc_paste_lambda.xlsx"
    lambda_probe.typed_workbook(build)
    truth = lambda_probe.HELD / "lambda_typed.xlsx"
    with Book(truth) as book:  # the names as the person defined them, in their file form
        names = [(n.name, n.text) for n in book.defined_names]
    wb = openpyxl.load_workbook(build)
    for name, text in names:
        wb.defined_names[name] = DefinedName(name, attr_text=text)
    wb.save(build)
    formulas = {5 * k + 1: f for k, f in enumerate(lambda_probe.TYPED)}
    differ += compare("lambda", build, formulas, truth)
    for line in differ:
        print(line)


if __name__ == "__main__":
    main()
