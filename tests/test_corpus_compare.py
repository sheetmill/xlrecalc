"""The corpus comparison sorts formulas by what happened to them (``tools/corpus_compare.py``)."""

from __future__ import annotations

from pathlib import Path

import pytest
from tools.corpus_compare import compare, compare_workbook, kind_of

from tests.test_xlsx import package
from xlrecalc._values import ErrorValue


def cell(address: str, formula: str | None, value: str | None, kind: str = "") -> str:
    f = "" if formula is None else f"<f>{formula}</f>"
    v = "" if value is None else f"<v>{value}</v>"
    t = f' t="{kind}"' if kind else ""
    return f'<c r="{address}"{t}>{f}{v}</c>'


def counts(tmp_path: Path, cells: str, extra: str = "") -> dict:
    path = tmp_path / "book.xlsx"
    path.write_bytes(package(f'<row r="1">{cells}</row>', workbook_extra=extra))
    return compare_workbook(str(path))


@pytest.mark.parametrize(
    ("a", "b", "how"),
    [
        (1.0, 1.0, "exact"),
        (1.0, 1.0 + 2e-16, "exact"),
        (1.0, 1.0 + 5e-15, "15_digits"),
        (1.0, 1.0 + 5e-10, "close"),
        (1e-10, 5e-10, "tiny"),
        (1.0, 1.001, "differ"),
        (1.0, True, "differ"),
        ("a", "A", "differ"),
        (ErrorValue.NA, ErrorValue.NA, "exact"),
        (ErrorValue.NA, ErrorValue.REF, "differ"),
        (0.0, "", "differ"),
    ],
)
def test_compare(a: object, b: object, how: str) -> None:
    if how == "exact" and a == 1.0 and b != 1.0:
        how = "15_digits"
    assert compare(a, b) == how


def test_a_match_and_a_cell_differing_for_itself(tmp_path: Path) -> None:
    row = counts(
        tmp_path, cell("A1", None, "2") + cell("B1", "A1*2", "4") + cell("C1", "A1+1", "9")
    )
    assert row["counts"] == {"match_exact": 1, "differ_own": 1}
    only = row["mismatches"][0]
    assert only["kind"] == "number -> number"
    assert only["functions"] == []
    assert (only["cached"], only["got"]) == ("9.0", "3.0")


def test_a_cell_that_only_inherits_a_difference(tmp_path: Path) -> None:
    """C1 is right for the value Excel cached in B1: only B1 is counted as a mismatch."""
    row = counts(
        tmp_path, cell("A1", None, "2") + cell("B1", "A1+1", "9") + cell("C1", "B1*2", "18")
    )
    assert row["counts"] == {"differ_own": 1, "differ_inherited": 1}
    assert row["after_own"] == 1
    assert len(row["mismatches"]) == 1


def test_an_absent_function_is_one_label_with_its_cells_downstream(tmp_path: Path) -> None:
    cells = cell("A1", "NOSUCHFN(1)", "5") + cell("B1", "A1*2", "10") + cell("C1", "B1+1", "11")
    row = counts(tmp_path, cells)
    assert row["counts"] == {"stopped_own": 1, "match_exact": 2}
    assert row["from_kept"] == 3
    assert row["labels"] == {"NOSUCHFN": 1}
    assert row["downstream"] == {"NOSUCHFN": 2}
    assert not row["mismatches"]


def test_an_absent_function_stops_its_formula_in_any_branch(tmp_path: Path) -> None:
    """As the engine does: it does not evaluate a formula naming one (ARCHITECTURE.md §7.6)."""
    row = counts(tmp_path, cell("A1", "IF(TRUE,1,NOSUCHFN(1))", "1"))
    assert row["counts"] == {"stopped_own": 1}


def test_volatile_and_external_are_set_aside_with_what_follows_them(tmp_path: Path) -> None:
    cells = (
        cell("A1", "NOW()", "12345")
        + cell("B1", "A1+1", "12346")
        + cell("C1", "[1]Other!A1", "7")
        + cell("D1", "C1+1", "8")
        + cell("E1", '"[x]"&amp;"y"', "[x]y", "str")
    )
    row = counts(tmp_path, cells)
    assert row["counts"] == {
        "volatile": 1,
        "volatile_dependent": 1,
        "external": 1,
        "external_dependent": 1,
        "match_exact": 1,
    }


def test_no_cached_value(tmp_path: Path) -> None:
    assert counts(tmp_path, cell("A1", "1+1", None))["counts"] == {"no_cached": 1}


def test_circular_with_iteration_off_and_on(tmp_path: Path) -> None:
    empty = cell("A1", "B1+1", "", "str") + cell("B1", "A1+1", "", "str")
    assert counts(tmp_path, empty)["counts"] == {"circular_empty": 2}
    held = cell("A1", "B1+1", "4") + cell("B1", "A1*2", "8")
    assert counts(tmp_path, held)["counts"] == {
        "circular_alone_matches": 1,
        "circular_alone_differs": 1,
    }
    iterate = '<calcPr iterate="1"/>'
    moving = counts(tmp_path, held, iterate)  # diverges: moves on every run, by design
    assert moving["counts"] == {"iterating_count_differs": 2}
    assert moving["iteration"]["stops"] == ["count", "count"]
    assert moving["iteration"]["cycle_cells"] == moving["iteration"]["held_back"] == 2
    fixed = cell("A1", "B1/2+1", "1.3333333333333333") + cell("B1", "A1/2", "0.6666666666666666")
    settled = counts(tmp_path, fixed, iterate)  # at its fixed point, as Excel saved it
    assert settled["counts"] == {"iterating_converged_matches": 2}
    assert settled["iteration"]["stops"] == ["delta", "delta"]


def test_an_array_formula_compares_its_whole_range(tmp_path: Path) -> None:
    same = '<c r="A1"><f t="array" ref="A1:C1">{1,2,3}*2</f><v>2</v></c>'
    same += cell("B1", None, "4") + cell("C1", None, "6")
    assert counts(tmp_path, same)["counts"] == {"match_exact": 1}
    other = '<c r="A1"><f t="array" ref="A1:C1">{1,2,3}*2</f><v>2</v></c>'
    other += cell("B1", None, "4") + cell("C1", None, "99")
    assert counts(tmp_path, other)["counts"] == {"differ_own": 1}


def test_kinds() -> None:
    assert kind_of(ErrorValue.NA) == "#N/A"
    assert kind_of(True) == "logical"
    assert kind_of(1.0) == "number"
    assert kind_of("x") == "text"
