"""The volatile functions the documents list are the ones the registry marks volatile."""

from __future__ import annotations

import math
import re
from pathlib import Path

import pytest

from xlrecalc._functions import SPECS

ROOT = Path(__file__).resolve().parents[1]


def _registry() -> set[str]:
    """The registry's volatile functions as the documents name them (no file prefix)."""
    return {name.rsplit(".", 1)[-1] if name.startswith("_XLFN.") else name
            for name, spec in SPECS.items() if spec.volatile}  # fmt: skip


def _names(text: str) -> set[str]:
    return set(re.findall(r"`([A-Z][A-Z.]*)`", text))


def test_architecture_lists_the_registry_volatile_functions() -> None:
    text = (ROOT / "ARCHITECTURE.md").read_text(encoding="utf-8")
    paragraph = re.search(r"\*\*Volatility\.\*\*(.*?)\n\n", text, re.S)
    assert paragraph, "ARCHITECTURE.md: the Volatility paragraph is missing"
    listed = re.search(r"marks volatile, as in Excel, are (.*?)\(", paragraph[1], re.S)
    assert listed
    assert _names(listed[1]) == _registry()


def test_the_glossary_lists_the_registry_volatile_functions() -> None:
    text = (ROOT / "docs" / "glossary.md").read_text(encoding="utf-8")
    entry = re.search(r"\nVolatile function\n: (.*?)\n\n", text, re.S)
    assert entry, "docs/glossary.md: the Volatile function entry is missing"
    listed = entry[1].split(":", 1)[1].split(". ", 1)[0]
    assert _names(listed) == _registry()


def test_no_other_list_of_volatile_functions() -> None:
    """Elsewhere, a mention points to the list: no sentence on volatility names three or more."""
    volatile = _registry()
    for path in [ROOT / "ARCHITECTURE.md", ROOT / "README.md", ROOT / "QUIRKS.md"]:
        text = path.read_text(encoding="utf-8")
        text = re.sub(r"\*\*Volatility\.\*\*.*?\n\n", "", text, flags=re.S)
        for sentence in re.split(r"(?<=[.;])\s", text):
            if "volatile" in sentence.lower():
                named = _names(sentence) & volatile
                assert len(named) < 3, f"{path.name}: {sorted(named)} beside 'volatile'"


def _volatile_edit() -> tuple[str, dict[str, list[list[str]]]]:
    """Capture ``volatile_edit``: its workbook part and its rows by variant."""
    text = (Path(__file__).parent / "captures" / "volatile_edit.tsv").read_text(encoding="utf-8")
    lines = text.splitlines()
    workbook = next(line for line in lines if line.startswith("# workbook: "))[
        len("# workbook: ") :
    ]
    rows: dict[str, list[list[str]]] = {}
    for line in lines:
        if line and not line.startswith("#"):
            fields = line.split("\t")
            rows.setdefault(fields[0], []).append(fields)
    return workbook, rows


# An unmarked cell volatile through a defined name, which Excel has not computed since opening, is
# recomputed by an edit neither as volatile (as the engine keeps it too) nor as a dependent
# through the name's references, which the engine follows (QUIRKS.md).
VOLATILE_EDIT_DEVIATIONS = {
    ("name_dependent_no_ca", "probe!B1"),
    ("name_offset_plus", "probe!B1"),
    ("name_static_dependent", "probe!B1"),
}


@pytest.mark.parametrize("read_first", [False, True])
@pytest.mark.parametrize("variant", list(_volatile_edit()[1]))
def test_an_edit_recomputes_what_excel_recomputes(variant: str, read_first: bool) -> None:
    """Capture ``volatile_edit``: each file as Excel saved and the probe planted it, opened, fully
    calculated first where the variant asks, edited (C1, or A1, set to 2), then calculated in
    full; every cell as Excel showed it at each step, but the deviations listed (strict: a
    listed cell that matches fails)."""
    import io

    from capture.volatile_edit_cases import EDITED, FULL_FIRST, FULL_ON_LOAD, VARIANTS

    import xlrecalc
    from tests.xlsx_files import package

    workbook, rows = _volatile_edit()
    if variant in FULL_ON_LOAD:
        workbook = workbook.replace("<calcPr ", '<calcPr fullCalcOnLoad="1" ', 1)
    cells = rows[variant]
    name = VARIANTS[variant][1]
    names = (
        ""
        if name is None
        else '<definedNames><definedName name="vn">'
        + name.replace('"', "&quot;")
        + "</definedName></definedNames>"
    )
    data = package(
        f'<row r="1">{"".join(c[2] for c in cells)}</row>',
        workbook_extra=workbook + names,
        sheets='<sheet name="probe" sheetId="1" r:id="rId1"/>',
    )
    wb = xlrecalc.load(io.BytesIO(data))
    refs = [c[1] for c in cells]
    unread = [None] * len(refs)
    steps: list[list[object]] = [[wb.value(r) for r in refs] if read_first else unread]
    if variant in FULL_FIRST:
        wb.recalculate(full=True)
        steps.append([wb.value(r) for r in refs])
    else:
        steps.append(unread)
    wb.set(EDITED.get(variant, "probe!C1"), 2)
    steps.append([wb.value(r) for r in refs])
    wb.recalculate(full=True)
    steps.append([wb.value(r) for r in refs])
    differ = set()
    for i, c in enumerate(cells):
        for step in range(4):
            got = steps[step][i]
            if got is None or c[3 + step] == "-":
                continue
            shown = float(c[3 + step])  # AppleScript's text of the value: 15 digits
            if not (type(got) is float and math.isclose(got, shown, rel_tol=1e-11)):
                differ.add((variant, c[1]))
    assert differ == {d for d in VOLATILE_EDIT_DEVIATIONS if d[0] == variant}
