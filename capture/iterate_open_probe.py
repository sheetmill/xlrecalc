"""What Excel does on opening a file with iterative calculation on (ARCHITECTURE.md §7.2).

    uv run python -m capture.iterate_open_probe

Non-interactive. Each variant is one workbook holding a counter (a cell adding 1 to itself), a loop
that converges to 2, a dependent, an unrelated formula, and a loop with no cached values. Cached
values are written into its XML, sentinels where a calculation would change them. Each file is
opened alone, with alerts off, and what Excel shows is read (``open_read.applescript``); it is then
opened again and a copy saved (``save_copy.applescript``), and the copy's cached values are read.
One variant is a file ``recalculate`` wrote, which asks for a full calculation. Others leave out
the loop without cached values, have iteration off, or are a copy Excel itself saved (with its
calculation chain), as it stands and after ``recalculate``. Writes
``tests/captures/iterate_open.tsv``.

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

import openpyxl

import xlrecalc
from capture.excel_watch import osascript_or_exit
from capture.pipeline import CAPTURES, CONTAINER, excel_version, read_results, result_repr

HERE = Path(__file__).resolve().parent
US = "\x1f"
STEP_LIMIT = 90  # seconds

# (tag, formula in column A, cached value or None)
CASES: list[tuple[str, str, str | None]] = [
    ("counter", "A1+1", "10"),
    ("loop", "A3/2+1", "5"),
    ("loop", "A2", "5"),
    ("dependent", "A2+100", "999"),
    ("control", "B5*2", "999"),
    ("uncached", "A7/2+1", None),
    ("uncached", "A6", None),
]
CACHED = CASES[:5]  # no cell without a cached value
ID = 'calcId="191029"'  # this Excel's own engine: no recalculation for an older one
VARIANTS: dict[str, str] = {
    "flag": f'<calcPr {ID} iterate="1" fullCalcOnLoad="1"/>',
    "no_flag": f'<calcPr {ID} iterate="1"/>',
    "flag_count_1": f'<calcPr {ID} iterate="1" iterateCount="1" fullCalcOnLoad="1"/>',
    "flag_count_3_delta_0": (
        f'<calcPr {ID} iterate="1" iterateCount="3" iterateDelta="0" fullCalcOnLoad="1"/>'
    ),
}


def build(path: Path, calc: str, cases: list[tuple[str, str, str | None]]) -> None:
    """The cases in column A, with cached values and ``calc`` as the file's calculation settings."""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "probe"
    for r, (_, formula, _) in enumerate(cases, 1):
        ws.cell(r, 1, "=" + formula)
    ws["B5"] = 2
    plain = path.with_name("xlrecalc_iterate_plain.xlsx")
    wb.save(plain)
    with zipfile.ZipFile(plain) as zin, zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            if item.filename == "xl/worksheets/sheet1.xml":
                text = data.decode()
                for r, (_, _, cached) in enumerate(cases, 1):
                    if cached is None:
                        continue
                    text, n = re.subn(
                        rf'(<c r="A{r}"[^>]*>\s*<f>[^<]*</f>)\s*<v\s*/>',
                        rf"\g<1><v>{cached}</v>",
                        text,
                    )
                    if n != 1:
                        raise RuntimeError(f"A{r}: not found")
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


def shown(path: Path, count: int) -> list[str]:
    """What Excel shows in column A on opening ``path``, closed again without saving."""
    fields = osascript(str(HERE / "open_read.applescript"), str(path), str(count)).split(US)
    if fields[0] != "open":
        sys.exit(f"{path.name} did not open: {fields}")
    return [fields[1 + 3 * i] for i in range(count)]


def calc_pr(path: Path) -> str:
    with zipfile.ZipFile(path) as z:
        found = re.search(r"<calcPr[^>]*/>", z.read("xl/workbook.xml").decode())
    return found[0] if found else ""


def main() -> None:
    count = osascript("-e", 'tell application "Microsoft Excel" to count of workbooks')
    if count != "0":
        sys.exit(f"Excel has {count} workbooks open: iteration is taken from the first one opened")
    CONTAINER.mkdir(parents=True, exist_ok=True)
    version = excel_version()
    header = [
        "# suite: iterate_open",
        f"# excel: {version} for Mac, locale en_US",
        f"# captured: {dt.date.today().isoformat()}",
        "# generator: capture/iterate_open_probe.py",
        "# method: a workbook per variant, cached values written into its XML, opened alone with",
        "#   alerts off; shown is what Excel shows on opening, saved_type and saved the cached",
        "#   value of a copy saved after opening again",
        "# columns: variant calcPr saved_calcPr row tag formula cached shown saved_type saved",
    ]
    body: list[str] = []
    files: list[tuple[str, Path, list[tuple[str, str, str | None]]]] = []
    for name, calc in VARIANTS.items():
        path = CONTAINER / f"xlrecalc_iterate_{name}.xlsx"
        build(path, calc, CASES)
        files.append((name, path, CASES))
    for name, calc in (
        ("no_flag_all_cached", f'<calcPr {ID} iterate="1"/>'),
        ("no_flag_all_cached_off", f"<calcPr {ID}/>"),
    ):
        path = CONTAINER / f"xlrecalc_iterate_{name}.xlsx"
        build(path, calc, CACHED)
        files.append((name, path, CACHED))
    written = CONTAINER / "xlrecalc_iterate_written.xlsx"
    report = xlrecalc.recalculate(files[1][1], written)
    print("recalculate:", dict(report.unsupported), len(report.uncomputed))
    files.append(("written_by_recalculate", written, CASES))
    made = [path for _, path, _ in files]

    def run(name: str, path: Path, cases: list[tuple[str, str, str | None]]) -> Path:
        copy = path.with_name(path.stem + "_saved.xlsx")
        copy.unlink(missing_ok=True)
        made.append(copy)
        seen = shown(path, len(cases))
        osascript(str(HERE / "save_copy.applescript"), str(path), str(copy))
        results = read_results(copy, len(cases))
        saved_calc = calc_pr(copy)
        for r, ((tag, formula, cached), value, result) in enumerate(
            zip(cases, seen, results, strict=True), 1
        ):
            kind, text = result_repr(result)
            row = [name, calc_pr(path), saved_calc, str(r), tag, formula, repr(cached), value]
            body.append("\t".join([*row, kind, text]))
            print(
                f"{name:24} {tag:10} {formula:8} cached {cached!s:5} shown {value:14} saved {text}"
            )
        return copy

    excel_copy: Path | None = None
    for name, path, cases in files:
        copy = run(name, path, cases)
        if name == "no_flag_all_cached":
            excel_copy = copy
    assert excel_copy is not None
    # The copy's cached values are what its first run saved; the cases' column shows the sentinels.
    held = CONTAINER / "xlrecalc_iterate_excel_copy.xlsx"
    held.write_bytes(excel_copy.read_bytes())
    made.append(held)
    run("excel_copy", held, CACHED)
    rewritten = CONTAINER / "xlrecalc_iterate_excel_copy_written.xlsx"
    report = xlrecalc.recalculate(held, rewritten)
    made.append(rewritten)
    print("recalculate:", dict(report.unsupported), len(report.uncomputed))
    run("excel_copy_written", rewritten, CACHED)
    for path in made:
        path.unlink(missing_ok=True)
    (CAPTURES / "iterate_open.tsv").write_text("\n".join([*header, *body]) + "\n", encoding="utf-8")


if __name__ == "__main__":
    main()
