"""The volatile functions the documents list are the ones the registry marks volatile."""

from __future__ import annotations

import itertools
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
    """Capture ``volatile_edit``: its workbook part and its rows by variant, each row as long as
    the variant's steps make it."""
    from capture.volatile_edit_cases import columns

    text = (Path(__file__).parent / "captures" / "volatile_edit.tsv").read_text(encoding="utf-8")
    lines = text.splitlines()
    workbook = next(line for line in lines if line.startswith("# workbook: "))[
        len("# workbook: ") :
    ]
    rows: dict[str, list[list[str]]] = {}
    for line in lines:
        if line and not line.startswith("#"):
            fields = line.split("\t")
            assert len(fields) == 3 + columns(fields[0]), f"a row of {fields[0]}: {len(fields)}"
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


def _differ(shown: list[str], got: list[object], random: bool) -> bool:
    """Whether the engine's values differ from those Excel showed at the steps both have ('-' and
    ``None`` are steps without a reading). A random cell agrees when each value lies in [2, 3)
    where Excel's does and changes where Excel's changes."""
    pairs = [(s, g) for s, g in zip(shown, got, strict=True) if s != "-" and g is not None]
    if not random:
        return not all(
            type(g) is float and math.isclose(g, float(s), rel_tol=1e-11)  # 15 digits shown
            for s, g in pairs
        )
    if any(type(g) is not float for _, g in pairs):
        return True
    inside = [(2 <= float(s) < 3, 2 <= g < 3) for s, g in pairs]  # type: ignore[operator]
    changed = [(a[0] != b[0], a[1] != b[1]) for a, b in itertools.pairwise(pairs)]
    return any(x != y for x, y in inside + changed)


@pytest.mark.parametrize("read_first", [False, True])
@pytest.mark.parametrize("variant", list(_volatile_edit()[1]))
def test_an_edit_recomputes_what_excel_recomputes(variant: str, read_first: bool) -> None:
    """Capture ``volatile_edit``: each file as Excel saved and the probe planted it, opened, then
    taken through the variant's steps (edits, ``recalculate()``, ``recalculate(full=True)``);
    every cell as Excel showed it after each step, but the deviations listed (strict: a listed
    cell that matches fails). In manual calculation a read evaluates what is stale, where Excel
    shows the cached values until it calculates, so the reading after an edit there is skipped."""
    import io

    from capture.volatile_edit_cases import (
        FULL_ON_LOAD,
        MANUAL,
        NAMES,
        RANDOM,
        VARIANTS,
        edit,
        steps,
    )

    import xlrecalc
    from tests.xlsx_files import package

    workbook, rows = _volatile_edit()
    if variant in FULL_ON_LOAD:
        workbook = workbook.replace("<calcPr ", '<calcPr fullCalcOnLoad="1" ', 1)
    if variant in MANUAL:
        workbook = workbook.replace("<calcPr ", '<calcPr calcMode="manual" ', 1)
    cells = rows[variant]
    name = VARIANTS[variant][1]
    defined = {**({} if name is None else {"vn": name}), **NAMES.get(variant, {})}
    names = "".join(
        f'<definedName name="{key}">' + text.replace('"', "&quot;") + "</definedName>"
        for key, text in defined.items()
    )
    data = package(
        f'<row r="1">{"".join(c[2] for c in cells)}</row>',
        workbook_extra=workbook + (f"<definedNames>{names}</definedNames>" if names else ""),
        sheets='<sheet name="probe" sheetId="1" r:id="rId1"/>',
    )
    wb = xlrecalc.load(io.BytesIO(data))
    refs = [c[1] for c in cells]
    unread: list[object] = [None] * len(refs)
    got: list[list[object]] = [[wb.value(r) for r in refs] if read_first else unread]
    after = steps(variant)
    if after[0] != "full":
        got.append(unread)
    for step in after:
        if step == "full":
            wb.recalculate(full=True)
        elif step == "calc":
            wb.recalculate()
        else:
            wb.set(*edit(step))
            if variant in MANUAL:
                got.append(unread)
                continue
        got.append([wb.value(r) for r in refs])
    differ = {
        (variant, c[1])
        for i, c in enumerate(cells)
        if _differ(c[3:], [g[i] for g in got], (variant, c[1]) in RANDOM)
    }
    assert differ == {d for d in VOLATILE_EDIT_DEVIATIONS if d[0] == variant}
