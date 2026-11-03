"""Probe how Excel treats a workbook whose cached formula values were patched into the XML.

    cd bench && uv run python ../capture/writeback_probe.py

Questions: does Excel open a file whose ``<v>`` and ``t`` were patched without a repair prompt;
does it show the patched values or recalculate; what removing ``calcChain.xml`` or setting
``fullCalcOnLoad`` changes; and whether Excel keeps or rebuilds ``calcChain.xml`` when it saves.

openpyxl writes the source workbook; Excel saves a copy (``save_copy.applescript``), which is a
file with Excel's own cached values and calculation chain. Each variant patches one of the two.
Patched values are deliberately wrong (sentinels), so a cell showing its sentinel was taken from
the file and a cell showing the right result was recalculated. ``open_read.applescript`` opens
each variant with alerts off: a file Excel would offer to repair fails to open with error -50.
Workbooks live in Excel's container; output goes to ``capture/build/writeback_probe.out``.
"""

from __future__ import annotations

import re
import subprocess
import sys
import zipfile
from pathlib import Path

import openpyxl

HERE = Path(__file__).resolve().parent
BUILD = HERE / "build"
CONTAINER = Path.home() / "Library/Containers/com.microsoft.Excel/Data/Documents/xlrecalc"
US = "\x1f"
ROWS = 7

# Column A. Correct results: 2, 6, x2, TRUE, #DIV/0!, 8, (empty).
FORMULAS = {"A2": "A1*3", "A3": '"x"&amp;A1', "A4": "A1&gt;1", "A5": "1/0", "A6": "SUM(A1:A2)"}
# Sentinels: (type attribute or None, value text).
SENTINELS = {
    "A2": (None, "999"),
    "A3": ("str", "SENT"),
    "A4": ("b", "0"),
    "A5": ("e", "#N/A"),
    "A6": (None, "777"),
}
SHEET = "xl/worksheets/sheet1.xml"


def build_source(path: Path) -> None:
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "probe"
    ws["A1"] = 2
    ws["A2"] = "=A1*3"
    ws["A3"] = '="x"&A1'
    ws["A4"] = "=A1>1"
    ws["A5"] = "=1/0"
    ws["A6"] = "=SUM(A1:A2)"
    wb.save(path)


def _cell_pattern(ref: str) -> re.Pattern[str]:
    return re.compile(rf'<c r="{ref}"(?P<attrs>[^>]*?)(?:/>|>(?P<body>.*?)</c>)', re.S)


def set_cell(xml: str, ref: str, formula: str | None, value: str, t: str | None) -> str:
    """Replace a cell element, keeping its style; add the row if the sheet has none."""
    pattern = _cell_pattern(ref)
    match = pattern.search(xml)
    style = ""
    if match:
        s = re.search(r'\ss="(\d+)"', match.group("attrs"))
        style = f' s="{s.group(1)}"' if s else ""
    t_attr = f' t="{t}"' if t else ""
    f_elem = f"<f>{formula}</f>" if formula is not None else ""
    new = f'<c r="{ref}"{style}{t_attr}>{f_elem}<v>{value}</v></c>'
    if match:
        return xml[: match.start()] + new + xml[match.end() :]
    row = int(re.sub(r"\D", "", ref))
    return xml.replace("</sheetData>", f'<row r="{row}">{new}</row></sheetData>', 1)


def patch_sentinels(xml: str) -> str:
    for ref, (t, value) in SENTINELS.items():
        xml = set_cell(xml, ref, FORMULAS[ref], value, t)
    return xml


def rewrite(
    src: Path, dst: Path, sheet=None, drop_chain=False, drop_chain_part_only=False, full_calc=None
) -> None:
    """Copy the package, editing only the named parts."""
    with zipfile.ZipFile(src) as zin, zipfile.ZipFile(dst, "w", zipfile.ZIP_DEFLATED) as zout:
        for info in zin.infolist():
            data = zin.read(info.filename)
            name = info.filename
            if name == "xl/calcChain.xml" and (drop_chain or drop_chain_part_only):
                continue
            if name == SHEET and sheet is not None:
                data = sheet(data.decode()).encode()
            elif name == "xl/_rels/workbook.xml.rels" and drop_chain:
                data = re.sub(rb"<Relationship [^>]*calcChain[^>]*/>", b"", data)
            elif name == "[Content_Types].xml" and drop_chain:
                data = re.sub(rb"<Override [^>]*calcChain[^>]*/>", b"", data)
            elif name == "xl/workbook.xml" and full_calc is not None:
                text = data.decode()
                text = re.sub(r'\sfullCalcOnLoad="[^"]*"', "", text)
                if full_calc:
                    text = re.sub(r"<calcPr\b", '<calcPr fullCalcOnLoad="1"', text, count=1)
                data = text.encode()
            zout.writestr(info, data)


def osascript(script: str, *args: str) -> str:
    result = subprocess.run(
        ["osascript", str(HERE / script), *args], capture_output=True, text=True
    )
    if result.returncode:
        raise RuntimeError(result.stderr.strip())
    return result.stdout.strip()


def parts(path: Path) -> str:
    with zipfile.ZipFile(path) as z:
        names = z.namelist()
        workbook = z.read("xl/workbook.xml").decode()
        sheet = z.read(SHEET).decode()
        chain = z.read("xl/calcChain.xml").decode() if "xl/calcChain.xml" in names else None
    calc = re.findall(r"<calcPr[^>]*>", workbook)
    cells = re.search(r"<sheetData>.*</sheetData>", sheet, re.S)
    lines = [
        f"  parts: {', '.join(sorted(names))}",
        f"  calcPr: {calc}",
        f"  sheetData: {cells.group(0) if cells else None}",
    ]
    lines.append(f"  calcChain: {re.findall(r'<c [^>]*/>', chain) if chain else None}")
    return "\n".join(lines)


def main() -> None:
    CONTAINER.mkdir(parents=True, exist_ok=True)
    BUILD.mkdir(exist_ok=True)
    out: list[str] = []

    def log(text: str) -> None:
        print(text, flush=True)
        out.append(text)

    log(
        "excel "
        + subprocess.run(
            ["osascript", "-e", 'tell application "Microsoft Excel" to version'],
            capture_output=True,
            text=True,
        ).stdout.strip()
    )
    src = CONTAINER / "xlrecalc_src.xlsx"
    base = CONTAINER / "xlrecalc_excel.xlsx"
    for p in CONTAINER.glob("xlrecalc_*.xlsx"):
        p.unlink()
    build_source(src)
    log(f"source (openpyxl {openpyxl.__version__}):\n{parts(src)}")
    log("save copy of source: " + osascript("save_copy.applescript", str(src), str(base)))
    log(f"excel-saved base:\n{parts(base)}")

    variants = {
        "O2_openpyxl_patched": {"src": src, "sheet": patch_sentinels},
        "O3_openpyxl_patched_nofullcalc": {
            "src": src,
            "sheet": patch_sentinels,
            "full_calc": False,
        },
        "E0_excel_untouched": {"src": base},
        "E1_patched_chain_kept": {"src": base, "sheet": patch_sentinels},
        "E2_patched_chain_removed": {"src": base, "sheet": patch_sentinels, "drop_chain": True},
        "E3_patched_chain_kept_fullcalc": {
            "src": base,
            "sheet": patch_sentinels,
            "full_calc": True,
        },
        "E4_patched_chain_removed_fullcalc": {
            "src": base,
            "sheet": patch_sentinels,
            "drop_chain": True,
            "full_calc": True,
        },
        "E5_formula_to_constant_chain_kept": {
            "src": base,
            "sheet": lambda x: set_cell(x, "A6", None, "5", None),
        },
        "E6_formula_added_chain_kept": {
            "src": base,
            "sheet": lambda x: set_cell(x, "A7", "A1+1", "555", None),
        },
        "E7_chain_part_deleted_rel_left": {
            "src": base,
            "sheet": patch_sentinels,
            "drop_chain_part_only": True,
        },
        "C1_control_bad_number": {
            "src": base,
            "sheet": lambda x: set_cell(x, "A2", "A1*3", "abc", None),
        },
    }
    paths = {}
    for name, spec in variants.items():
        dst = CONTAINER / f"xlrecalc_{name}.xlsx"
        spec = dict(spec)
        rewrite(spec.pop("src"), dst, **spec)
        paths[name] = dst

    log("\nopen and read (value | text | formula per row of A1:A7):")
    for name, path in paths.items():
        fields = osascript("open_read.applescript", str(path), str(ROWS)).split(US)
        if fields[0] == "open":
            rows = [" | ".join(fields[1 + 3 * i : 4 + 3 * i]) for i in range(ROWS)]
            log(f"{name}: OPEN {fields[-2]} {fields[-1]}\n    " + "\n    ".join(rows))
        else:
            log(f"{name}: REFUSED {fields[1:]}")

    log("\ncalcChain after Excel saves a copy:")
    for name in (
        "O2_openpyxl_patched",
        "E2_patched_chain_removed",
        "E5_formula_to_constant_chain_kept",
        "E6_formula_added_chain_kept",
        "E1_patched_chain_kept",
    ):
        copy = CONTAINER / f"xlrecalc_saved_{name}.xlsx"
        try:
            osascript("save_copy.applescript", str(paths[name]), str(copy))
            log(f"{name}:\n{parts(copy)}")
        except RuntimeError as exc:
            log(f"{name}: save failed: {exc}")

    (BUILD / "writeback_probe.out").write_text("\n".join(out) + "\n")


if __name__ == "__main__":
    sys.exit(main())
