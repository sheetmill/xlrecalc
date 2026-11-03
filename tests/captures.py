"""Load the capture suites in ``tests/captures`` (written by ``capture/pipeline.py``)."""

from __future__ import annotations

import ast
import datetime as dt
from dataclasses import dataclass
from pathlib import Path

from xlrecalc._values import BLANK, ErrorValue, Value
from xlrecalc._xlsx._cells import unescape

CAPTURES = Path(__file__).with_name("captures")
EMPTY_TEXT = '=""'


@dataclass(frozen=True)
class Formula:
    """An input cell holding a formula (display form, with ``=``)."""

    text: str


@dataclass(frozen=True)
class Row:
    suite: str
    tag: str
    case: int
    formula: str
    inputs: tuple[Value, ...]
    expected: Value  # what Excel saved for the formula cell

    @property
    def id(self) -> str:
        return f"{self.suite}:{self.case}:{self.tag}"


@dataclass(frozen=True)
class Suite:
    name: str
    header: dict[str, str]
    rows: list[Row]

    @property
    def captured(self) -> dt.date:
        return dt.date.fromisoformat(self.header["captured"])

    @property
    def date1904(self) -> bool:
        return self.header.get("date_system") == "1904"

    @property
    def fixed(self) -> dict[str, Value]:
        """Cells written once for the whole suite (lookup tables)."""
        text = self.header.get("fixed")
        if not text:
            return {}
        tree = ast.parse(text, mode="eval").body
        assert isinstance(tree, ast.Dict)
        return {
            ast.literal_eval(k): _literal(v)  # type: ignore[arg-type, misc]
            for k, v in zip(tree.keys, tree.values, strict=True)
        }


def _literal(node: ast.expr) -> Value:
    """An input or expected value: Python literals, ``ErrorValue(...)``, ``FormulaInput(...)``."""
    if isinstance(node, ast.Call) and isinstance(node.func, ast.Name):
        arg = ast.literal_eval(node.args[0])
        if node.func.id == "ErrorValue":
            return ErrorValue(arg)
        if node.func.id == "FormulaInput" and arg == EMPTY_TEXT:
            return ""
        if node.func.id == "FormulaInput" and arg.startswith('="') and arg.endswith('"'):
            return arg[2:-1].replace('""', '"')
        if node.func.id == "FormulaInput":
            return Formula(arg)
        raise ValueError(f"unknown input {ast.unparse(node)}")
    value = ast.literal_eval(node)
    if value is None:
        return BLANK
    if isinstance(value, int) and not isinstance(value, bool):
        return float(value)
    return value  # type: ignore[no-any-return]


def _inputs(text: str) -> tuple[Value, ...]:
    tree = ast.parse(text, mode="eval").body
    assert isinstance(tree, ast.Tuple)
    return tuple(_literal(e) for e in tree.elts)


def _expected(kind: str, text: str) -> Value:
    if kind == "e":
        return ErrorValue(text)
    if kind == "none":
        return BLANK
    value = ast.literal_eval(text)
    if kind == "str":  # as written in the saved file: decode _xHHHH_ as Excel reads it
        return unescape(value)
    return float(value) if kind == "n" else value  # type: ignore[no-any-return]


def load(name: str) -> Suite:
    header: dict[str, str] = {}
    rows: list[Row] = []
    for line in (CAPTURES / f"{name}.tsv").read_text(encoding="utf-8").splitlines():
        if line.startswith("# ") and ": " in line:
            key, _, value = line[2:].partition(": ")
            header.setdefault(key, value)
            continue
        if line.startswith("#"):
            continue
        suite, tag, case, formula, inputs, kind, expected = line.split("\t")
        rows.append(Row(suite, tag, int(case), formula, _inputs(inputs), _expected(kind, expected)))
    return Suite(name, header, rows)
