"""When does Excel recalculate a workbook on opening? Cached values planted wrong, one change each.

    uv run python -m capture.open_calc_probe
    uv run python -m capture.open_calc_probe --calc-ids N,N,...   (calcId variants only)

openpyxl writes a workbook (A1 = 2; A2:A6 formulas of it, as in ``writeback_probe``; A7 the
volatile ``RAND()*0+A1`` in a second one); Excel saves a copy of each, a file with Excel's own
cached values, ``calcPr`` and calculation chain. Each variant plants wrong cached values
(sentinels: 999, "SENT", FALSE, #N/A, 777, and 555 in A7) in that copy and changes one thing:
nothing; ``calcPr`` removed; ``calcId`` older or newer than the copy's, or removed;
``fullCalcOnLoad``; the volatile cell; manual calculation (``calcMode="manual"``), alone, with
``fullCalcOnLoad`` and with an old ``calcId``. Two more are openpyxl's own files: with no cached
values, and with the sentinels. Four look for another trigger, one difference between openpyxl's
file and Excel's each: openpyxl's file with Excel's ``calcId`` and no ``fullCalcOnLoad``; Excel's
copy without its calculation chain, without ``<fileVersion>``, and without ``docProps/app.xml`` (a
part is dropped with its relationship and content type). Each variant is opened with alerts off and
read (``open_read.applescript``: value, displayed text, formula, saved flag and calculation mode),
closed without saving, then opened again and saved as a copy whose cached values are read. A cell
showing its sentinel was taken from the file; one showing its right value was recalculated.

Excel's calculation mode belongs to the application and the first workbook opened sets it, so
the manual variants come last and ``calc_automatic.applescript`` sets the mode back after each.
Workbooks live in Excel's container and are removed after; the record is written to
``capture/build/open_calc_probe.out`` and the capture to ``tests/captures/open_calc.tsv``. Rerun
it after every Excel update: ``VERIFIED_CALC_ID`` (``_load.py``) follows the threshold it finds.
"""

from __future__ import annotations

import re
import subprocess
import sys
import zipfile
from pathlib import Path

import openpyxl

from capture.edition import edition_line
from capture.excel_watch import preflight
from capture.writeback_probe import (
    CONTAINER,
    SHEET,
    build_source,
    osascript,
    patch_sentinels,
    rewrite,
    set_cell,
)

HERE = Path(__file__).resolve().parent
OUT = HERE / "build" / "open_calc_probe.out"
CAPTURE = HERE.parent / "tests" / "captures" / "open_calc.tsv"
# calcIds around the threshold, and those of earlier Excel versions (2010 to 2019).
CALC_IDS = (145621, 152511, 162913, 171027, 181029, 190000, 191000, 191001, 191010, 191020)
CALC_IDS += (191025, 191027, 191028)
US = "\x1f"
ROWS = 7


def build_volatile(path: Path) -> None:
    build_source(path)
    wb = openpyxl.load_workbook(path)
    wb.active["A7"] = "=RAND()*0+A1"
    wb.save(path)


def calc_attributes(path: Path) -> str:
    """The workbook part's ``<fileVersion>`` and ``<calcPr>`` as written, each when present."""
    with zipfile.ZipFile(path) as z:
        workbook = z.read("xl/workbook.xml").decode()
    found = [m.group(0) for m in re.finditer(r"<(?:fileVersion|calcPr)\b[^>]*>", workbook)]
    return "".join(found) if found else "(neither)"


def edit_calc(  # type: ignore[no-untyped-def]
    src: Path, dst: Path, sheet, edit, drop: str | None = None
) -> None:
    """Copy ``src`` with the sheet patched and ``edit`` applied to workbook.xml's text.

    ``drop`` removes a part with its relationship and its content-type override, so the package
    stays whole (a dangling relationship would make Excel offer a repair).
    """
    rewrite(src, dst, sheet=sheet)
    with zipfile.ZipFile(dst) as zin:
        items = [(info, zin.read(info.filename)) for info in zin.infolist()]
    target = drop.rsplit("/", 1)[-1] if drop else None
    with zipfile.ZipFile(dst, "w", zipfile.ZIP_DEFLATED) as zout:
        for info, data in items:
            if info.filename == drop:
                continue
            if info.filename == "xl/workbook.xml":
                data = edit(data.decode()).encode()
            if target and info.filename.endswith(".rels"):
                data = re.sub(
                    rf'<Relationship\b[^>]*Target="[^"]*{re.escape(target)}"[^>]*/>',
                    "",
                    data.decode(),
                ).encode()
            if drop and info.filename == "[Content_Types].xml":
                data = re.sub(
                    rf'<Override\b[^>]*PartName="/{re.escape(drop)}"[^>]*/>', "", data.decode()
                ).encode()
            zout.writestr(info, data)


def _fileversion(find: str, replace: str):  # type: ignore[no-untyped-def]
    """An edit of workbook.xml's ``<fileVersion>`` element alone."""

    def edit(text: str) -> str:
        return re.sub(
            r"<fileVersion\b[^>]*/>", lambda m: re.sub(find, replace, m.group(0)), text, count=1
        )

    return edit


def saved_values(path: Path) -> str:
    with zipfile.ZipFile(path) as z:
        sheet = z.read(SHEET).decode()
    cells = re.findall(r'<c r="(A\d)"([^>]*?)(?:/>|>(.*?)</c>)', sheet, re.S)
    out = []
    for ref, attrs, body in cells:
        value = re.search(r"<v>(.*?)</v>", body or "")
        t = re.search(r'\st="([^"]*)"', attrs)
        out.append(f"{ref}={(t.group(1) + ':') if t else ''}{value.group(1) if value else '-'}")
    return " ".join(out)


def main() -> None:
    preflight()
    CONTAINER.mkdir(parents=True, exist_ok=True)
    OUT.parent.mkdir(exist_ok=True)
    lines: list[str] = []

    def log(text: str) -> None:
        print(text, flush=True)
        lines.append(text)

    for p in CONTAINER.glob("xlrecalc_oc_*.xlsx"):
        p.unlink()
    version = subprocess.run(
        ["osascript", "-e", 'tell application "Microsoft Excel" to version'],
        capture_output=True,
        text=True,
    ).stdout.strip()
    log(f"excel {version}; openpyxl {openpyxl.__version__}")
    src, vsrc = CONTAINER / "xlrecalc_oc_src.xlsx", CONTAINER / "xlrecalc_oc_vsrc.xlsx"
    base, vbase = CONTAINER / "xlrecalc_oc_base.xlsx", CONTAINER / "xlrecalc_oc_vbase.xlsx"
    build_source(src)
    build_volatile(vsrc)
    log("save copy of source: " + osascript("save_copy.applescript", str(src), str(base)))
    log(
        "save copy of volatile source: " + osascript("save_copy.applescript", str(vsrc), str(vbase))
    )
    log(f"base calcPr: {calc_attributes(base)}; openpyxl calcPr: {calc_attributes(src)}")
    calc_id = int(re.search(r'calcId="(\d+)"', calc_attributes(base)).group(1))  # type: ignore[union-attr]

    def with_id(n: int):  # type: ignore[no-untyped-def]
        return lambda t: re.sub(r'calcId="\d+"', f'calcId="{n}"', t)

    def volatile_sentinels(x: str) -> str:
        return set_cell(patch_sentinels(x), "A7", "RAND()*0+A1", "555", None)

    variants = {
        "patched": (base, patch_sentinels, lambda t: t),
        "no_calcpr": (base, patch_sentinels, lambda t: re.sub(r"<calcPr\b[^>]*/>", "", t)),
        "calcid_124519": (base, patch_sentinels, with_id(124519)),
        **{f"calcid_{n}": (base, patch_sentinels, with_id(n)) for n in CALC_IDS},
        "calcid_newer_by_1": (base, patch_sentinels, with_id(calc_id + 1)),
        "calcid_999999": (base, patch_sentinels, with_id(999999)),
        "fullcalconload": (
            base,
            patch_sentinels,
            lambda t: re.sub(r"<calcPr\b", '<calcPr fullCalcOnLoad="1"', t, count=1),
        ),
        "volatile": (vbase, volatile_sentinels, lambda t: t),
        # What else makes Excel recalculate: one difference between openpyxl's file and Excel's.
        "openpyxl_calcid_current": (
            src,
            patch_sentinels,
            lambda t: re.sub(r"\s*fullCalcOnLoad=\"1\"", "", with_id(calc_id)(t)),
        ),
        "no_calcchain": (base, patch_sentinels, lambda t: t, "xl/calcChain.xml"),
        "no_fileversion": (
            base,
            patch_sentinels,
            lambda t: re.sub(r"<fileVersion\b[^>]*/>", "", t),
        ),
        "no_app_properties": (base, patch_sentinels, lambda t: t, "docProps/app.xml"),
        # Which part of <fileVersion> matters (Excel 16.115 writes appName, lastEdited and
        # lowestEdited 7, and rupBuild).
        **{
            f"fileversion_{label}": (base, patch_sentinels, _fileversion(find, replace))
            for label, find, replace in (
                ("no_appname", r'\sappName="[^"]*"', ""),
                ("appname_calc", r'appName="[^"]*"', 'appName="Calc"'),
                ("no_rupbuild", r'\srupBuild="[^"]*"', ""),
                ("rupbuild_4505", r'rupBuild="[^"]*"', 'rupBuild="4505"'),
                ("lastedited_6", r'lastEdited="[^"]*"', 'lastEdited="6"'),
                ("lastedited_4", r'lastEdited="[^"]*"', 'lastEdited="4"'),
                ("lowestedited_4", r'lowestEdited="[^"]*"', 'lowestEdited="4"'),
            )
        },
        "openpyxl_no_values": (src, None, lambda t: t),
        "openpyxl_patched": (src, patch_sentinels, lambda t: t),
        "calcpr_without_calcid": (
            base,
            patch_sentinels,
            lambda t: re.sub(r'\scalcId="\d+"', "", t),
        ),
        "manual": (
            base,
            patch_sentinels,
            lambda t: re.sub(r"<calcPr\b", '<calcPr calcMode="manual"', t, count=1),
        ),
        "manual_fullcalconload": (
            base,
            patch_sentinels,
            lambda t: re.sub(
                r"<calcPr\b", '<calcPr calcMode="manual" fullCalcOnLoad="1"', t, count=1
            ),
        ),
        "manual_calcid_124519": (
            base,
            patch_sentinels,
            lambda t: re.sub(
                r"<calcPr\b", '<calcPr calcMode="manual"', with_id(124519)(t), count=1
            ),
        ),
    }
    if sys.argv[1:2] == ["--calc-ids"]:  # only these calcId variants, to place a threshold
        ids = [int(n) for n in sys.argv[2].split(",")]
        variants = {f"calcid_{n}": (base, patch_sentinels, with_id(n)) for n in ids}
    rows: list[str] = []
    try:
        for name, (origin, sheet, edit, *drop) in variants.items():
            path = CONTAINER / f"xlrecalc_oc_{name}.xlsx"
            edit_calc(origin, path, sheet, edit, *drop)
            log(f"\n{name}: calcPr {calc_attributes(path)}; file {saved_values(path)}")
            fields = osascript("open_read.applescript", str(path), str(ROWS)).split(US)
            if fields[0] != "open":
                log(f"  REFUSED {fields[1:]}")
                continue
            shown = [" | ".join(fields[1 + 3 * i : 4 + 3 * i]) for i in range(ROWS)]
            log(f"  opened: {fields[-2]} {fields[-1]}\n    " + "\n    ".join(shown))
            copy = CONTAINER / f"xlrecalc_oc_{name}_saved.xlsx"
            osascript("save_copy.applescript", str(path), str(copy))
            log(f"  saved copy: calcPr {calc_attributes(copy)}; {saved_values(copy)}")
            texts = [fields[2 + 3 * i] for i in range(1, ROWS)]  # displayed text of A2:A7
            mode = fields[-1].removeprefix("calc=calculation ")
            shown_texts = " ".join(map(repr, texts))
            fields_out = (name, calc_attributes(path), saved_values(path), shown_texts, mode)
            rows.append("\t".join((*fields_out, saved_values(copy))))
            if name.startswith("manual"):  # the application's mode back to automatic each time
                log("  restore: " + osascript("calc_automatic.applescript", str(base)))
    finally:
        log("\nrestore: " + osascript("calc_automatic.applescript", str(base)))
        OUT.write_text("\n".join(lines) + "\n", encoding="utf-8")
        if rows and sys.argv[1:2] != ["--calc-ids"]:
            header = [
                "# suite: open_calc (when Excel recalculates a workbook on opening)",
                f"# excel: {version} for Mac, locale en_US",
                edition_line(),
                "# generator: capture/open_calc_probe.py (cached values planted wrong in a copy",
                "#   Excel saved: 999, SENT, FALSE, #N/A, 777 in A2:A6, 555 in A7 where volatile;",
                "#   right values 6, x2, TRUE, #DIV/0!, 8, 2)",
                "# columns: variant, fileVersion and calcPr as written, the file's cells,",
                "#   displayed text of A2:A7 on opening, calculation mode after opening, the cells",
                "#   of the copy Excel saved",
            ]
            CAPTURE.write_text("\n".join([*header, *rows]) + "\n", encoding="utf-8")
        for p in CONTAINER.glob("xlrecalc_oc_*.xlsx"):
            p.unlink()


if __name__ == "__main__":
    sys.exit(main())
