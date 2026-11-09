"""What Excel's repair does to a formula it refuses in a file (ARCHITECTURE.md §9.5).

    uv run python -m capture.repair_probe [--uncached]

The one interactive capture here: Excel refuses to open a file holding such a formula unless
a person accepts its repair dialog, so this opens the workbook with alerts on and waits for
that. The workbook holds every formula ``capture/formula_forms.py`` found refused, one per row
of column A, between control formulas. Every formula cell is given a deliberately wrong cached
value (``100 + row``) and the workbook's calculation settings are those of a file Excel saved
(``calcId="191029"``, no ``fullCalcOnLoad``), so Excel shows cached values as they are: after
the repair, each refused formula's cell shows whether its cached value was kept, replaced or
dropped. The script reads column A, saves a copy of the repaired workbook, and reads that
copy's cells too. Writes ``tests/captures/repair.tsv``.

With ``--uncached`` the refused formulas keep openpyxl's empty ``<v />`` (no cached value) while
the controls keep theirs, and column C holds ``=ISBLANK(A{row})`` for every row, so the saved
copy shows what a refused formula without a cached value becomes. Writes
``tests/captures/repair_uncached.tsv``.
"""

from __future__ import annotations

import datetime as dt
import re
import subprocess
import sys
import xml.etree.ElementTree as ET
import zipfile
from pathlib import Path

import openpyxl
from openpyxl.worksheet.table import Table

from capture.edition import edition_line
from capture.pipeline import CAPTURES, CONTAINER, NS, excel_version

HERE = Path(__file__).resolve().parent
US = "\x1f"
REFUSED = [
    "#SPILL!",
    "#CALC!",
    "A1#",
    "Sheet2!A1#",
    "Table1[@[Unit Price]]",
    "[@[Unit Price]]*2",
    "-\t1",
    "SUM( )",
]
CONTROLS = ["1+1", "B1*3", '"ok"']


def rows() -> list[tuple[str, str]]:
    out = [("control", CONTROLS[0])]
    for k, formula in enumerate(REFUSED):
        out.append(("refused", formula))
        out.append(("control", CONTROLS[(k + 1) % len(CONTROLS)]))
    return out


def build(path: Path, cases: list[tuple[str, str]], uncached: bool = False) -> None:
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "probe"
    ws["B1"] = 2
    data = wb.create_sheet("Sheet2")
    data.append(["Price", "Unit Price"])
    data.append([1, 2])
    data.add_table(Table(displayName="Table1", ref="A1:B2"))
    for r, (_, formula) in enumerate(cases, 1):
        ws.cell(r, 1, "=" + formula)
        if uncached:
            ws.cell(r, 3, f"=ISBLANK(A{r})")
    wb.save(path)
    skip = {r for r, (tag, _) in enumerate(cases, 1) if tag == "refused"} if uncached else set()
    _trust_cached_values(path, len(cases), skip)


def _trust_cached_values(path: Path, count: int, skip: set[int]) -> None:
    """Give each formula cell of column A the cached value 100 + row, as Excel's own files.

    Rows in ``skip`` keep openpyxl's empty ``<v />``; column C's formulas get the cached value 0.
    """
    with zipfile.ZipFile(path) as z:
        parts = {name: z.read(name) for name in z.namelist()}
    sheet = parts["xl/worksheets/sheet1.xml"].decode()
    for r in range(1, count + 1):
        sheet = re.sub(
            rf'(<c r="C{r}"[^>]*>)(<f>.*?</f>)<v ?/>', r"\1\2<v>0</v>", sheet, count=1, flags=re.S
        )
        if r in skip:
            continue
        sheet = re.sub(
            rf'(<c r="A{r}"[^>]*>)(<f>.*?</f>)<v ?/>',
            lambda m, r=r: f"{m[1]}{m[2]}<v>{100 + r}</v>",
            sheet,
            count=1,
            flags=re.S,
        )
    parts["xl/worksheets/sheet1.xml"] = sheet.encode()
    workbook = parts["xl/workbook.xml"].decode()
    workbook = re.sub(r"<calcPr[^>]*/>", '<calcPr calcId="191029"/>', workbook)
    parts["xl/workbook.xml"] = workbook.encode()
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
        for name, data in parts.items():
            z.writestr(name, data)


def saved_cells(path: Path, column: str = "A") -> dict[int, str]:
    """Each cell of a column of the saved copy as compact XML."""
    with zipfile.ZipFile(path) as z:
        root = ET.fromstring(z.read("xl/worksheets/sheet1.xml"))
    out = {}
    for c in root.iter(f"{NS}c"):
        ref = c.get("r", "")
        if re.fullmatch(rf"{column}\d+", ref):
            out[int(ref[1:])] = _plain(ET.tostring(c, encoding="unicode"))
    return out


def _plain(xml: str) -> str:
    """An element's XML without ElementTree's namespace prefix and declaration."""
    xml = re.sub(r' xmlns:\w+="[^"]*"', "", xml)
    return re.sub(r"(</?)\w+:", r"\1", xml)


def main() -> None:
    uncached = "--uncached" in sys.argv[1:]
    suite = "repair_uncached" if uncached else "repair"
    cases = rows()
    CONTAINER.mkdir(parents=True, exist_ok=True)
    src = CONTAINER / f"xlrecalc_{suite}.xlsx"
    copy = CONTAINER / f"xlrecalc_{suite}_saved.xlsx"
    for p in (src, copy):
        p.unlink(missing_ok=True)
    build(src, cases, uncached)
    version = excel_version()
    print(f"opening {src.name} with alerts on: accept Excel's repair dialogs")
    result = subprocess.run(
        ["osascript", str(HERE / "repair_probe.applescript"), str(src), str(len(cases)), str(copy)],
        capture_output=True,
        text=True,
    )
    if result.returncode:
        sys.exit(f"osascript failed: {result.stderr.strip()}")
    lines = result.stdout.rstrip("\n").split("\n")
    repaired_name = lines[0].split(US)[1]
    seen = {int(f[0]): f[1:] for f in (line.split(US) for line in lines[1:])}
    saved = saved_cells(copy)
    saved_c = saved_cells(copy, "C") if uncached else {}
    header = [
        f"# suite: {suite}",
        f"# excel: {version} for Mac, locale en_US",
        edition_line(),
        f"# captured: {dt.date.today().isoformat()}",
        "# generator: capture/repair_probe.py",
        "# method: openpyxl writes each formula into column A with the cached value 100+row and",
        *(
            [
                "#   (refused formulas: no cached value, openpyxl's <v />); column C holds",
                "#   =ISBLANK(A<row>) cached 0, and saved_c is that cell in the saved copy",
            ]
            if uncached
            else []
        ),
        '#   calcPr calcId="191029" (Excel trusts cached values); Excel opens it with alerts',
        "#   on and a person accepts the repair; formula, value and text are read from the",
        f"#   repaired workbook ({repaired_name}); saved is the cell as its saved copy holds it",
        "# columns: row tag written formula_after value_after text_after saved"
        + (" saved_c" if uncached else "")
        + " (Python repr)",
    ]
    body = []
    for r, (tag, formula) in enumerate(cases, 1):
        formula_after, value, text = seen.get(r, ["", "", ""])
        body.append(
            "\t".join(
                [
                    str(r),
                    tag,
                    repr(formula),
                    repr(formula_after),
                    repr(value),
                    repr(text),
                    repr(saved.get(r, "")),
                    *([repr(saved_c.get(r, ""))] if uncached else []),
                ]
            )
        )
        print(
            f"{r:3} {tag:8} {formula!r:26} -> {formula_after!r:24} {value!r:10} {saved.get(r, '')}"
        )
    (CAPTURES / f"{suite}.tsv").write_text("\n".join([*header, *body]) + "\n", encoding="utf-8")
    src.unlink()
    copy.unlink(missing_ok=True)


if __name__ == "__main__":
    main()
