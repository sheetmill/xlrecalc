"""Which Excel actions run a loop whose cells have cached values (ARCHITECTURE.md §7.2).

    uv run python -m capture.iterate_cached_probe

Non-interactive. openpyxl writes a loop (``A1 = IFERROR(A2,0)+1``, ``A2 = A1+1``, 100 iterations a
run, each run adding 200) with no cached values; Excel saves a copy, ``X``, with its own cached
values, ``calcChain.xml`` and ``<f ca="1">``. ``X`` is kept, Excel's personal markers removed, as
``tests/captures/iterate_cached.xlsx``, the input of the end-to-end tests. Then Excel opens ``X``
and saves a copy, and opens ``X``, enters a number in a cell nothing reads (``C1``) and saves a
copy; the copies' cached values are read. Writes ``tests/captures/iterate_cached.tsv``.

Needs Excel with no workbook open, and refuses to run otherwise.
"""

from __future__ import annotations

import datetime as dt
import re
import sys
import zipfile
from pathlib import Path

from capture.iterate_probe import HERE, build, osascript, saved_values
from capture.pipeline import CAPTURES, CONTAINER, excel_version

CALC = '<calcPr calcId="124519" iterate="1" iterateCount="100" fullCalcOnLoad="1"/>'
LOOP = [("S", [("A1", "IFERROR(A2,0)+1", None), ("A2", "A1+1", None)])]


def strip(source: Path, target: Path) -> None:
    """``source`` without the folder it was saved in or the user who saved it."""
    with zipfile.ZipFile(source) as zin, zipfile.ZipFile(target, "w", zipfile.ZIP_DEFLATED) as out:
        for item in zin.infolist():
            data = zin.read(item.filename)
            if item.filename == "xl/workbook.xml":
                data = re.sub(
                    rb"<mc:AlternateContent[^>]*>.*?absPath.*?</mc:AlternateContent>",
                    b"",
                    data,
                    flags=re.S,
                )
                data = re.sub(rb"<x15ac:absPath[^>]*/>", b"", data)
            if item.filename == "docProps/core.xml":
                data = re.sub(
                    rb"<cp:lastModifiedBy>[^<]*</cp:lastModifiedBy>|<cp:lastModifiedBy/>", b"", data
                )
            out.writestr(item, data)


def main() -> None:
    count = osascript("-e", 'tell application "Microsoft Excel" to count of workbooks')
    if count != "0":
        sys.exit(f"Excel has {count} workbooks open: iteration is taken from the first one opened")
    CONTAINER.mkdir(parents=True, exist_ok=True)
    plain = CONTAINER / "xlrecalc_iter_cached_plain.xlsx"
    x = CONTAINER / "xlrecalc_iter_cached_x.xlsx"
    opened = CONTAINER / "xlrecalc_iter_cached_opened.xlsx"
    edited = CONTAINER / "xlrecalc_iter_cached_edited.xlsx"
    made = [plain, x, opened, edited]
    rows = []
    try:
        for path in made:
            path.unlink(missing_ok=True)
        build(plain, CALC, LOOP)
        osascript(str(HERE / "save_copy.applescript"), str(plain), str(x))
        strip(x, CAPTURES / "iterate_cached.xlsx")
        osascript(str(HERE / "save_copy.applescript"), str(x), str(opened))
        osascript(str(HERE / "edit_save_copy.applescript"), str(x), str(edited), "S!C1", "1")
        for step, path in (("x", x), ("open_save", opened), ("open_edit_save", edited)):
            values = saved_values(path)
            rows += ["\t".join([step, ref, *values[ref]]) for ref in ("S!A1", "S!A2")]
    finally:
        for path in made:
            path.unlink(missing_ok=True)
    header = [
        "# suite: iterate_cached",
        f"# excel: {excel_version()} for Mac, locale en_US",
        f"# captured: {dt.date.today().isoformat()}",
        "# generator: capture/iterate_cached_probe.py",
        "# method: x is Excel's saved copy of the openpyxl loop (kept as iterate_cached.xlsx);",
        "#   open_save is Excel opening x and saving a copy; open_edit_save is Excel opening x,",
        "#   entering 1 in S!C1 (read by nothing) and saving a copy",
        "# columns: step cell type value",
    ]
    (CAPTURES / "iterate_cached.tsv").write_text(
        "\n".join([*header, *rows]) + "\n", encoding="utf-8"
    )
    print("\n".join(rows))


if __name__ == "__main__":
    main()
