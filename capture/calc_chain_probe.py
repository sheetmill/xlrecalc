"""Whether Excel's calculation chain sets the order in which it iterates (ARCHITECTURE.md §7.2).

    uv run python -m capture.calc_chain_probe

Non-interactive. Two loops, each in a workbook with one iteration a run: one within a sheet
(``A1`` and ``A5``), one across two sheets. Excel saves each once, writing its ``calcChain.xml``;
from that copy come three files, with every cached value set back to 0: the chain as Excel wrote
it, the chain reversed, and no chain. Each is opened alone: what Excel shows after opening and
after one ``calculate`` is read (``iterate_read.applescript``); then it is opened again and a copy
saved, whose values and chain are read. That copy, values set back to 0 again, is opened once
more: the order a chain has after a save. Writes ``tests/captures/calc_chain.tsv``.

Iteration is an application setting that Excel takes from the first workbook it opens, so this
needs Excel with no workbook open, and refuses to run otherwise. Each Excel step has a time limit;
when one passes, the probe stops.
"""

from __future__ import annotations

import datetime as dt
import re
import sys
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET

from capture.iterate_probe import HERE, NS, REL, US, Sheets, build, osascript, saved_values
from capture.pipeline import CAPTURES, CONTAINER, excel_version

CALC = '<calcPr calcId="191029" iterate="1" iterateCount="1"/>'
LOOPS: list[tuple[str, Sheets]] = [
    ("rows", [("s", [("A1", "A5+1", "0"), ("A5", "A1+10", "0")])]),
    ("sheets", [("one", [("A1", "two!A1+1", "0")]), ("two", [("A1", "one!A1+10", "0")])]),
]
CHAIN = "xl/calcChain.xml"


def chain_of(path: Path) -> list[tuple[str, str]]:
    """The chain's cells in order, as (sheet name, cell); empty without one."""
    with zipfile.ZipFile(path) as z:
        if CHAIN not in z.namelist():
            return []
        book = ET.fromstring(z.read("xl/workbook.xml"))
        ids = {s.get("sheetId"): s.get("name", "") for s in book.iter(f"{NS}sheet")}
        out, sheet = [], ""
        for c in ET.fromstring(z.read(CHAIN)).iter(f"{NS}c"):
            sheet = c.get("i", sheet)
            out.append((ids.get(sheet, f"?{sheet}"), c.get("r", "")))
    return out


def rewrite(source: Path, target: Path, order: list[tuple[str, str]] | None) -> None:
    """``source`` with every formula's cached value 0 and the chain ``order``, or none."""
    with zipfile.ZipFile(source) as z:
        book = ET.fromstring(z.read("xl/workbook.xml"))
        ids = {s.get("name"): s.get("sheetId") for s in book.iter(f"{NS}sheet")}
        rels = ET.fromstring(z.read("xl/_rels/workbook.xml.rels"))
        targets = {r.get("Id"): r.get("Target", "") for r in rels}
        sheet_parts = {
            "xl/" + targets[s.get(REL)].lstrip("/").removeprefix("xl/")
            for s in book.iter(f"{NS}sheet")
        }
        with zipfile.ZipFile(target, "w", zipfile.ZIP_DEFLATED) as out:
            for item in z.infolist():
                data = z.read(item.filename)
                if item.filename in sheet_parts:
                    text = re.sub(  # a formula: Excel writes <f ca="1"> in a loop
                        r"(<c r=\"[A-Z]+\d+\"[^>]*>\s*<f[^>]*>[^<]*</f>\s*)<v>[^<]*</v>",
                        r"\g<1><v>0</v>",
                        data.decode(),
                    )
                    data = text.encode()
                if item.filename == CHAIN:
                    if order is None:
                        continue
                    cells = "".join(f'<c r="{r}" i="{ids[s]}"/>' for s, r in order)
                    data = (
                        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
                        f'<calcChain xmlns="{NS[1:-1]}">{cells}</calcChain>'
                    ).encode()
                if order is None and item.filename == "xl/_rels/workbook.xml.rels":
                    data = re.sub(rb"<Relationship [^>]*calcChain[^>]*/>", b"", data)
                if order is None and item.filename == "[Content_Types].xml":
                    data = re.sub(rb"<Override [^>]*calcChain[^>]*/>", b"", data)
                out.writestr(item, data)


def main() -> None:
    made: list[Path] = []
    try:
        run(made)
    finally:  # what a stopped run made goes too
        for path in made:
            path.unlink(missing_ok=True)


def run(made: list[Path]) -> None:
    count = osascript("-e", 'tell application "Microsoft Excel" to count of workbooks')
    if count != "0":
        sys.exit(f"Excel has {count} workbooks open: iteration is taken from the first one opened")
    CONTAINER.mkdir(parents=True, exist_ok=True)
    body: list[str] = []

    def save(path: Path) -> Path:
        copy = path.with_name(path.stem + "_saved.xlsx")
        copy.unlink(missing_ok=True)
        made.append(copy)
        osascript(str(HERE / "save_copy.applescript"), str(path), str(copy))
        return copy

    def observe(loop: str, variant: str, path: Path, refs: list[str]) -> Path:
        fields = osascript(str(HERE / "iterate_read.applescript"), str(path), *refs).split(US)
        if fields[0] != "open":
            sys.exit(f"{path.name} did not open: {fields}")
        k = len(refs)
        shown, calculated = fields[1 : 1 + 2 * k : 2], fields[2 + 2 * k :: 2]
        copy = save(path)
        saved = saved_values(copy)
        chain_in = " ".join(f"{s}!{r}" for s, r in chain_of(path)) or "none"
        chain_out = " ".join(f"{s}!{r}" for s, r in chain_of(copy)) or "none"
        for i, ref in enumerate(refs):
            row = [loop, variant, chain_in, ref, shown[i], calculated[i], saved[ref][1], chain_out]
            body.append("\t".join(row))
            print("  ".join(row), flush=True)
        return copy

    for loop, sheets in LOOPS:
        refs = [f"{s}!{r}" for s, cells in sheets for r, _, _ in cells]
        plain = CONTAINER / f"xlrecalc_chain_{loop}.xlsx"
        made.append(plain)
        build(plain, CALC, sheets)
        first = save(plain)  # Excel's own chain
        excel = chain_of(first)
        variants = {"excel_chain": excel, "reversed_chain": excel[::-1], "no_chain": None}
        for variant, order in variants.items():
            path = CONTAINER / f"xlrecalc_chain_{loop}_{variant}.xlsx"
            made.append(path)
            rewrite(first, path, order)
            copy = observe(loop, variant, path, refs)
            again = CONTAINER / f"xlrecalc_chain_{loop}_{variant}_again.xlsx"
            made.append(again)
            rewrite(copy, again, chain_of(copy) or None)
            observe(loop, variant + "_after_save", again, refs)
    header = [
        "# suite: calc_chain",
        f"# excel: {excel_version()} for Mac, locale en_US",
        f"# captured: {dt.date.today().isoformat()}",
        "# generator: capture/calc_chain_probe.py",
        "# method: a loop saved by Excel, then its chain kept, reversed or removed, cached values",
        "#   set to 0, iterateCount 1; shown after opening, calculated after one calculate, saved",
        "#   the value in a copy saved after opening again, chain_out that copy's chain;",
        "#   *_after_save is that copy, values set to 0 again, opened once more",
        "# columns: loop variant chain_in cell shown calculated saved chain_out",
    ]
    (CAPTURES / "calc_chain.tsv").write_text("\n".join([*header, *body]) + "\n", encoding="utf-8")
    print(f"-> calc_chain.tsv, {len(body)} rows")


if __name__ == "__main__":
    main()
