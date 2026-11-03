"""What Excel computes and saves for circular references with iteration off (ARCHITECTURE.md §7.2).

    uv run python -m capture.circular_probe [--interactive]

Non-interactive. openpyxl writes cycles of every shape into column A (two cells, one cell on
itself, three cells, text, a reference in an untaken branch, cells depending on cycles) with
constants in column B; Excel opens it, calculates and saves a copy (``save_copy.applescript``),
and the copy's cached values and calculation settings are read from its XML. Then the cycle
cells of that copy are given sentinel cached values (``900 + row``) and the copy is opened with
alerts off (``open_read.applescript``) to see whether Excel shows the file's values for cells on
a cycle or recalculates them. Writes ``tests/captures/circular.tsv``.

With ``--interactive`` the openpyxl file is opened with alerts on instead
(``repair_probe.applescript``), so Excel shows its circular-reference warning and a person
dismisses it; column A is read and a copy saved. Writes ``tests/captures/circular_alerts.tsv``.
"""

from __future__ import annotations

import datetime as dt
import re
import subprocess
import sys
import zipfile
from pathlib import Path

import openpyxl

from capture.pipeline import CAPTURES, CONTAINER, excel_version, read_results, result_repr
from capture.repair_probe import saved_cells

HERE = Path(__file__).resolve().parent
US = "\x1f"

# (tag, formula in column A, constant in column B or None)
CASES: list[tuple[str, str, float | None]] = [
    ("pair", "A2+1", None),
    ("pair", "A1+1", None),
    ("self", "A3+1", None),
    ("dependent", "A1*2", None),
    ("control", "1+1", None),
    ("self_range", "A6+B6", 5.0),
    ("triple", "A8+1", None),
    ("triple", "A9+1", None),
    ("triple", "A7+1", None),
    ("text", '"x"&A11', None),
    ("text", "A10", None),
    ("iferror", "IFERROR(1/0,A12)", None),
    ("untaken", "IF(FALSE,A13,5)", None),
    ("dependent2", "A3+A1", None),
    ("dependent_const", "A4+B15", 1.0),
    ("control", "B16*3", 2.0),
]
CYCLE_TAGS = {"pair", "self", "self_range", "triple", "text", "iferror", "untaken"}


def build(path: Path) -> None:
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "probe"
    for r, (_, formula, constant) in enumerate(CASES, 1):
        ws.cell(r, 1, "=" + formula)
        if constant is not None:
            ws.cell(r, 2, constant)
    wb.save(path)


def calc_pr(path: Path) -> str:
    with zipfile.ZipFile(path) as z:
        workbook = z.read("xl/workbook.xml").decode()
    m = re.search(r"<calcPr[^>]*/>", workbook)
    return m[0] if m else ""


def with_sentinels(source: Path, target: Path) -> None:
    """A copy of Excel's saved file whose cycle cells hold the cached value 900 + row."""
    with zipfile.ZipFile(source) as z:
        parts = {info: z.read(info.filename) for info in z.infolist()}
    out = {}
    for info, data in parts.items():
        if info.filename == "xl/worksheets/sheet1.xml":
            sheet = data.decode()
            for r, (tag, _, _) in enumerate(CASES, 1):
                if tag not in CYCLE_TAGS:
                    continue
                sheet, n = re.subn(
                    rf'<c r="A{r}"([^>]*?)(?: t="\w+")?([^>]*)>(<f[^>]*>.*?</f>)'
                    r"(?:<v>[^<]*</v>|<v ?/>)</c>",
                    rf'<c r="A{r}"\1\2>\3<v>{900 + r}</v></c>',
                    sheet,
                    count=1,
                    flags=re.S,
                )
                if n != 1:
                    raise RuntimeError(f"A{r}: cell not found in the saved copy")
            data = sheet.encode()
        out[info] = data
    with zipfile.ZipFile(target, "w", zipfile.ZIP_DEFLATED) as z:
        for info, data in out.items():
            z.writestr(info.filename, data)


def osascript(*args: str) -> str:
    result = subprocess.run(["osascript", *args], capture_output=True, text=True)
    if result.returncode:
        sys.exit(f"osascript failed: {result.stderr.strip()}")
    return result.stdout.rstrip("\n")


def shown_values(path: Path) -> list[str]:
    """What Excel shows in column A when it opens ``path`` (alerts off), as value|text."""
    shown = osascript(str(HERE / "open_read.applescript"), str(path), str(len(CASES)))
    fields = shown.split(US)
    if fields[0] != "open":
        sys.exit(f"{path.name} did not open: {shown}")
    return [f"{fields[1 + 3 * i]}|{fields[2 + 3 * i]}" for i in range(len(CASES))]


def interactive() -> None:
    CONTAINER.mkdir(parents=True, exist_ok=True)
    src = CONTAINER / "xlrecalc_circular_alerts.xlsx"
    copy = CONTAINER / "xlrecalc_circular_alerts_saved.xlsx"
    for p in (src, copy):
        p.unlink(missing_ok=True)
    build(src)
    version = excel_version()
    print(f"opening {src.name} with alerts on: dismiss Excel's circular-reference warning")
    out = osascript(
        str(HERE / "repair_probe.applescript"), str(src), str(len(CASES)), str(copy)
    ).split("\n")
    seen = {int(f[0]): f[1:] for f in (line.split(US) for line in out[1:])}
    xml = saved_cells(copy)
    header = [
        "# suite: circular_alerts",
        f"# excel: {version} for Mac, locale en_US",
        f"# captured: {dt.date.today().isoformat()}",
        "# generator: capture/circular_probe.py --interactive",
        "# method: openpyxl writes the formulas into column A (constants in B, no cached values,",
        "#   fullCalcOnLoad); Excel opens it with alerts on, a person dismisses the",
        "#   circular-reference warning; formula, value and text are read from the open",
        "#   workbook, saved is the cell as Excel's saved copy holds it",
        f"# saved_calcPr: {calc_pr(copy)}",
        "# columns: row tag formula constant value text saved",
    ]
    body = []
    for r, (tag, formula, constant) in enumerate(CASES, 1):
        _, value, text = seen.get(r, ["", "", ""])
        body.append(
            "\t".join([str(r), tag, formula, repr(constant), value, text, repr(xml.get(r, ""))])
        )
        print(f"{r:3} {tag:15} {formula:17} {value!r:8} {text!r:8} {xml.get(r, '')}")
    (CAPTURES / "circular_alerts.tsv").write_text(
        "\n".join([*header, *body]) + "\n", encoding="utf-8"
    )
    for p in (src, copy):
        p.unlink(missing_ok=True)


def main() -> None:
    if "--interactive" in sys.argv[1:]:
        interactive()
        return
    CONTAINER.mkdir(parents=True, exist_ok=True)
    src = CONTAINER / "xlrecalc_circular.xlsx"
    copy = CONTAINER / "xlrecalc_circular_saved.xlsx"
    patched = CONTAINER / "xlrecalc_circular_patched.xlsx"
    for p in (src, copy, patched):
        p.unlink(missing_ok=True)
    build(src)
    version = excel_version()
    osascript(str(HERE / "save_copy.applescript"), str(src), str(copy))
    results = read_results(copy, len(CASES))
    saved_calc = calc_pr(copy)
    shown_saved = shown_values(copy)
    with_sentinels(copy, patched)
    shown_patched = shown_values(patched)
    header = [
        "# suite: circular",
        f"# excel: {version} for Mac, locale en_US",
        f"# captured: {dt.date.today().isoformat()}",
        "# generator: capture/circular_probe.py",
        "# method: openpyxl writes the formulas into column A (constants in B, no cached values,",
        "#   fullCalcOnLoad); Excel opens it, calculates and saves a copy; expected is the copy's",
        "#   cached value (type, repr) and saved its XML; shown_saved is what Excel shows when it",
        "#   opens that copy (value|text, alerts off). Then the copy's cycle cells get the cached",
        "#   value 900+row: shown_patched is what Excel shows when it opens that file",
        f"# saved_calcPr: {saved_calc}",
        "# columns: row tag formula constant expected_type expected_repr saved shown_saved"
        " shown_patched",
    ]
    body = []
    xml = saved_cells(copy)
    for r, ((tag, formula, constant), result, a, b) in enumerate(
        zip(CASES, results, shown_saved, shown_patched, strict=True), 1
    ):
        kind, text = result_repr(result)
        body.append(
            "\t".join(
                [str(r), tag, formula, repr(constant), kind, text, repr(xml.get(r, "")), a, b]
            )
        )
        print(f"{r:3} {tag:15} {formula:17} {xml.get(r, '')[:58]:58} {a:10} {b}")
    print(saved_calc)
    (CAPTURES / "circular.tsv").write_text("\n".join([*header, *body]) + "\n", encoding="utf-8")
    for p in (src, copy, patched):
        p.unlink(missing_ok=True)


if __name__ == "__main__":
    main()
