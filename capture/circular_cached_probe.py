"""What Excel does with a loop whose file holds cached values, iteration off (ARCHITECTURE.md §7.2).

    uv run python -m capture.circular_cached_probe

Non-interactive. Capture ``circular`` had no cached values to start from. Here each file holds a
loop of two cells, a cell on itself, a text loop, cells depending on the loop and a control, all
with cached values that are sentinels, and optionally a loop with none. A file is made either by
openpyxl, or from a copy Excel itself saved with its calculation chain, with the sentinels written
into it; with and without ``fullCalcOnLoad``. Each file is opened once per action (nothing, a
recalculation, a full recalculation, a rebuild, each followed by a read of what Excel shows and a
saved copy, whose cached values are read from its XML). Writes
``tests/captures/circular_cached.tsv``.

Excel takes its calculation settings from the first workbook it opens, so the probe refuses to run
with a workbook open. Each Excel step has a time limit; when one passes, the probe stops and leaves
Excel as it is.
"""

from __future__ import annotations

import datetime as dt
import re
import sys
import zipfile
from pathlib import Path

import openpyxl

from capture.excel_watch import osascript_or_exit
from capture.pipeline import CAPTURES, CONTAINER, excel_version
from capture.repair_probe import saved_cells

HERE = Path(__file__).resolve().parent
US = "\x1f"
STEP_LIMIT = 90  # seconds
ID = 'calcId="191029"'  # this Excel's own engine: no recalculation for an older one

# (tag, formula in column A, cached value, cached type)
Case = tuple[str, str, str | None, str | None]
HELD: list[Case] = [
    ("pair", "A2+1", "901", None),
    ("pair", "A1+1", "902", None),
    ("self", "A3+1", "903", None),
    ("dependent", "A1*2", "904", None),
    ("dependent2", "A4+1", "905", None),
    ("control", "B6*2", "906", None),
    ("text", '"x"&A8', "held", "str"),
    ("text", "A7", "also", "str"),
]
EMPTY: list[Case] = [
    ("no_start", "A10+1", None, None),
    ("no_start", "A9+1", None, None),
    ("no_start_dependent", "A9*2", None, None),
]
ACTIONS = ("none", "calculate", "full", "rebuild")
# A loop reading a constant (B1), with room for formulas an edit adds (A5, A6).
EDITED: list[Case] = [
    ("pair", "A2+B1", "901", None),
    ("pair", "A1+1", "902", None),
    ("dependent", "A1*2", "903", None),
    ("control", "B1*2", "904", None),
]
EDITS = ("set_input", "set_formula", "break", "new_cycle", "new_dependent")


def _patch_cell(sheet: str, row: int, cached: str | None, kind: str | None) -> str:
    """Column A of ``row`` with this cached value (no value: as openpyxl writes it)."""
    pattern = (
        rf'<c r="A{row}"([^>]*?)(?: t="\w+")?([^>]*)>(<f[^>]*>.*?</f>)(?:<v>[^<]*</v>|<v ?/>)?</c>'
    )
    t = f' t="{kind}"' if kind else ""
    value = "<v />" if cached is None else f"<v>{cached}</v>"
    sheet, n = re.subn(pattern, rf'<c r="A{row}"\1\2{t}>\3{value}</c>', sheet, count=1, flags=re.S)
    if n != 1:
        raise RuntimeError(f"A{row}: cell not found")
    return sheet


def patched(source: Path, target: Path, cases: list[Case], flag: bool, calc_id: str = ID) -> None:
    """A copy of ``source`` with cached values for ``cases``, and the calculation settings asked."""
    with zipfile.ZipFile(source) as zin, zipfile.ZipFile(target, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            if item.filename == "xl/worksheets/sheet1.xml":
                text = data.decode()
                for r, (_, _, cached, kind) in enumerate(cases, 1):
                    text = _patch_cell(text, r, cached, kind)
                data = text.encode()
            if item.filename == "xl/workbook.xml":
                full = ' fullCalcOnLoad="1"' if flag else ""
                settings = f"<calcPr {calc_id}{full}/>"
                text, n = re.subn(r"<calcPr[^>]*/>", settings, data.decode())
                if n != 1:
                    raise RuntimeError("calcPr not found")
                data = text.encode()
            zout.writestr(item, data)


def plain(path: Path, cases: list[Case]) -> None:
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "probe"
    for r, (_, formula, _, _) in enumerate(cases, 1):
        ws.cell(r, 1, "=" + formula)
    if cases is EDITED:
        ws["B1"] = 1
    else:
        ws["B6"] = 2
    wb.save(path)


def osascript(*args: str) -> str:
    result = osascript_or_exit(*args, seconds=STEP_LIMIT)
    if result.returncode:
        sys.exit(f"osascript failed: {result.stderr.strip()}")
    return result.stdout.rstrip("\n")


def act(path: Path, count: int, action: str, copy: Path) -> list[tuple[str, str]]:
    """What Excel shows in column A after ``action``, with a copy saved after that."""
    copy.unlink(missing_ok=True)
    fields = osascript(
        str(HERE / "circular_cached.applescript"), str(path), str(count), action, str(copy)
    ).split(US)
    if fields[0] != "open":
        sys.exit(f"{path.name} did not open: {fields}")
    return [(fields[1 + 2 * i], fields[2 + 2 * i]) for i in range(count)]


def main() -> None:
    count = osascript("-e", 'tell application "Microsoft Excel" to count of workbooks')
    if count != "0":
        sys.exit(f"Excel has {count} workbooks open: its settings come from the first one opened")
    CONTAINER.mkdir(parents=True, exist_ok=True)
    version = excel_version()
    made: list[Path] = []

    def place(name: str) -> Path:
        path = CONTAINER / f"xlrecalc_cc_{name}.xlsx"
        path.unlink(missing_ok=True)
        made.append(path)
        return path

    held = HELD
    both = HELD + EMPTY
    base = {
        "held": (place("base_held"), held),
        "both": (place("base_both"), both),
    }
    variants: list[tuple[str, Path, list[Case]]] = []
    for label, (source, cases) in base.items():
        plain(source, cases)
        for flag in (False, True):
            name = f"openpyxl_{label}_{'flag' if flag else 'no_flag'}"
            target = place(name)
            patched(source, target, cases, flag)
            variants.append((name, target, cases))
        # An older engine's id, which is how a file from an earlier Excel arrives.
        name = f"openpyxl_{label}_old_engine"
        target = place(name)
        patched(source, target, cases, False, 'calcId="124519"')
        variants.append((name, target, cases))
        # Excel's own file, with its calculation chain: saved once, then given the sentinels.
        saved = place(f"excel_saved_{label}")
        act(source, len(cases), "none", saved)
        for flag in (False, True):
            name = f"excel_{label}_{'flag' if flag else 'no_flag'}"
            target = place(name)
            patched(saved, target, cases, flag)
            variants.append((name, target, cases))

    source = place("base_edited")
    plain(source, EDITED)
    saved = place("excel_saved_edited")
    act(source, len(EDITED), "none", saved)
    target = place("excel_edited")
    patched(saved, target, EDITED, False)
    edited = ("excel_edited", target)

    header = [
        "# suite: circular_cached",
        f"# excel: {version} for Mac, locale en_US",
        f"# captured: {dt.date.today().isoformat()}",
        "# generator: capture/circular_cached_probe.py",
        "# method: iteration off; a real loop whose file holds sentinel cached values, in a file",
        "#   openpyxl built or Excel saved, with or without fullCalcOnLoad; opened alone with",
        "#   alerts off, one action per opening; shown is value|text after it, saved is the cell",
        "#   as the copy saved after it holds it",
        "# columns: variant action row tag formula cached shown saved",
    ]
    body: list[str] = []
    for name, path, cases in variants:
        for action in ACTIONS:
            copy = place("copy")
            seen = act(path, len(cases), action, copy)
            xml = saved_cells(copy)
            for r, ((tag, formula, cached, _), (value, text)) in enumerate(
                zip(cases, seen, strict=True), 1
            ):
                fields = [name, action, str(r), tag, formula, repr(cached), f"{value}|{text}"]
                body.append("\t".join([*fields, repr(xml.get(r, ""))]))
            print(name, action, [v for v, _ in seen], flush=True)
    name, path = edited
    for action in EDITS:
        copy = place("copy")
        seen = act(path, 6, action, copy)
        xml = saved_cells(copy)
        for r, (value, text) in enumerate(seen, 1):
            tag, formula, cached, _ = (
                EDITED[r - 1] if r <= len(EDITED) else ("added", "", None, None)
            )
            fields = [name, action, str(r), tag, formula, repr(cached), f"{value}|{text}"]
            body.append("\t".join([*fields, repr(xml.get(r, ""))]))
        print(name, action, [v for v, _ in seen], flush=True)
    for path in made:
        path.unlink(missing_ok=True)
    (CAPTURES / "circular_cached.tsv").write_text(
        "\n".join([*header, *body]) + "\n", encoding="utf-8"
    )


if __name__ == "__main__":
    main()
