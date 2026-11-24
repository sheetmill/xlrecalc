"""Capture ``spill_edit``: after an edit makes a dynamic array spill further or less, every cell
reads what Excel shows, whichever is read first (§7.7).

Each case of ``capture/spill_edit_cases.py`` is written as a file with its dynamic arrays in
Excel's form and no cached values, so the engine computes every formula on opening, as Excel did.
After opening and after each edit, the case's cells are read in their listed order, in reverse,
each one first, and in one pass (``Workbook.values_at``), and compared with the capture.
"""

from __future__ import annotations

import io
from pathlib import Path

import pytest
from capture.spill_edit_cases import CASES

import xlrecalc
from tests.xlsx_files import DOC, MAIN, PKG, package

HERE = Path(__file__).parent
METADATA = (HERE / "captures" / "arrays_metadata.xml").read_text("utf-8")


def _captured() -> dict[str, dict[str, list[float | None]]]:
    rows: dict[str, dict[str, list[float | None]]] = {}
    text = (HERE / "captures" / "spill_edit.tsv").read_text(encoding="utf-8")
    for line in text.splitlines():
        if line and not line.startswith("#"):
            case, ref, *shown = line.split("\t")
            rows.setdefault(case, {})[ref] = [None if s == "" else float(s) for s in shown]
    return rows


CAPTURED = _captured()


def _xml_text(text: str) -> str:
    return text.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")


def _sheet(cells: dict[str, object]) -> str:
    """A worksheet part of the cells (address -> content), in row and column order."""
    by_row: dict[int, list[str]] = {}
    for where, content in cells.items():
        row = int("".join(ch for ch in where if ch.isdigit()))
        if isinstance(content, float):
            xml = f'<c r="{where}"><v>{content!r}</v></c>'
        elif content[0] == "#":  # type: ignore[index]
            text = _xml_text(content[1])  # type: ignore[index]
            xml = f'<c r="{where}" cm="1"><f t="array" ref="{where}">{text}</f></c>'
        else:
            xml = f'<c r="{where}"><f>{_xml_text(content[1])}</f></c>'  # type: ignore[index]
        by_row.setdefault(row, []).append((len(where), where, xml))  # type: ignore[arg-type]
    rows = "".join(
        f'<row r="{r}">' + "".join(x for _, _, x in sorted(by_row[r])) + "</row>"  # type: ignore[misc]
        for r in sorted(by_row)
    )
    return f'<worksheet xmlns="{MAIN}"><sheetData>{rows}</sheetData></worksheet>'


def _book(case: str) -> xlrecalc.Workbook:
    cells, _, _ = CASES[case]
    sheets: dict[str, dict[str, object]] = {"probe": {}}
    for ref, content in cells.items():
        sheet, where = ref.split("!")
        sheets.setdefault(sheet, {})[where] = content
    names = list(sheets)
    rels = "".join(
        f'<Relationship Id="rId{i}" Type="{DOC}/worksheet" Target="worksheets/sheet{i}.xml"/>'
        for i in range(1, len(names) + 1)
    )
    parts: dict[str, str | bytes] = {
        "xl/metadata.xml": METADATA,
        "xl/_rels/workbook.xml.rels": f'<Relationships xmlns="{PKG}">{rels}'
        f'<Relationship Id="rId8" Type="{DOC}/sheetMetadata" Target="metadata.xml"/>'
        "</Relationships>",
    }
    for i, name in enumerate(names[1:], start=2):
        parts[f"xl/worksheets/sheet{i}.xml"] = _sheet(sheets[name])
    data = package(
        sheet_xml=_sheet(sheets["probe"]),
        sheets="".join(
            f'<sheet name="{n}" sheetId="{i + 1}" r:id="rId{i + 1}"/>' for i, n in enumerate(names)
        ),
        workbook_extra='<fileVersion appName="xl"/><calcPr calcId="191029"/>',
        extra_parts=parts,
    )
    return xlrecalc.load(io.BytesIO(data))


def _orders(refs: tuple[str, ...]) -> list[tuple[str, list[str]]]:
    out = [("listed", list(refs)), ("reversed", list(reversed(refs)))]
    out += [(f"{r} first", [r, *(x for x in refs if x != r)]) for r in refs]
    return out


@pytest.mark.parametrize("batch", [False, True], ids=["one_by_one", "values_at"])
@pytest.mark.parametrize("case", list(CASES))
def test_each_cell_reads_what_excel_shows_after_each_edit(case: str, batch: bool) -> None:
    _, steps, refs = CASES[case]
    shown = CAPTURED[case]
    assert set(shown) == set(refs)
    for label, order in _orders(refs):
        wb = _book(case)
        for at in range(1 + len(steps)):
            if at:
                target, number = steps[at - 1][len("edit:") :].split("=")
                wb.set(target, float(number))
            values = wb.values_at(order) if batch else [wb.value(r) for r in order]
            for ref, got in zip(order, values, strict=True):
                assert got == shown[ref][at], (label, at, ref)
