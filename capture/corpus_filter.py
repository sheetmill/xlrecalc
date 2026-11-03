"""Decide which formulas from a downloaded workbook may go into a workbook Excel computes.

Formulas found in third-party or real-world workbooks are never opened in Excel as they are.
The capture tooling writes a fresh workbook holding only accepted formulas and the constant
inputs they refer to (ARCHITECTURE.md §10). ``reject`` names the reason a formula is kept out:

- ``dde``: a DDE link (``app|'topic'!item``), which asks Excel to start or call a program;
- ``external``: a reference to another workbook (``[1]Sheet1!A1``, ``'[1]Sheet 1'!A1``,
  ``[1]!Name``), or ``INDIRECT`` of a text naming one;
- ``side effects: NAME``: a function that reaches outside the workbook (files, network, other
  programs, the machine) or reports on the environment;
- ``hyperlink``: ``HYPERLINK``, harmless, but its result is not worth capturing;
- ``unparseable``: the tokenizer or parser refuses it.

``tally`` counts the reasons over many formulas, for the report of a corpus run.
"""

from __future__ import annotations

import collections
import re
from collections.abc import Iterable

from xlrecalc._formula import FormulaSyntaxError, Kind, TokenizeError, parse, tokenize

# Functions with side effects, file or network access, or answers about the machine.
SIDE_EFFECTS = frozenset(
    {
        "WEBSERVICE",
        "FILTERXML",
        "ENCODEURL",
        "RTD",
        "CALL",
        "REGISTER.ID",
        "INFO",
        "CELL",
        "IMAGE",
        "STOCKHISTORY",
        "STOCKSERIES",
        "TRANSLATE",
        "DETECTLANGUAGE",
        "COPILOT",
        "PY",
        "SQL.REQUEST",
        "EUROCONVERT",
        "CUBEMEMBER",
        "CUBEVALUE",
        "CUBESET",
        "CUBESETCOUNT",
        "CUBERANKEDMEMBER",
        "CUBEMEMBERPROPERTY",
        "CUBEKPIMEMBER",
        "GETPIVOTDATA",
    }
)
_PREFIXES = ("_xlfn._xlws.", "_xlfn.", "_xlws.", "_xludf.")
# app|topic!item, with an optionally quoted topic: what the formula bar shows for a DDE link.
_DDE = re.compile(r"[A-Za-z_][\w.]*\|")


def function_name(text: str) -> str:
    """A function token's name without Excel's file-form prefixes, upper-cased."""
    name = text.upper()
    for prefix in _PREFIXES:
        if name.startswith(prefix.upper()):
            return name[len(prefix) :]
    return name


def reject(formula: str) -> str | None:
    """Why a formula (file form, no ``=``) must not be captured, or ``None`` to accept it."""
    if _DDE.search(_outside_strings(formula)):
        return "dde"
    try:
        toks = tokenize(formula)
        parse(formula)
    except (TokenizeError, FormulaSyntaxError):
        return "unparseable"
    indirect = False
    for tok in toks:
        if tok.kind is Kind.PREFIX and "[" in tok.text:
            return "external"
        if tok.kind is Kind.FUNCTION:
            name = function_name(tok.text)
            if name in SIDE_EFFECTS:
                return f"side effects: {name}"
            if name == "HYPERLINK":
                return "hyperlink"
            indirect = indirect or name == "INDIRECT"
        if indirect and tok.kind is Kind.STRING and "[" in tok.text:
            return "external"
    return None


def tally(formulas: Iterable[str]) -> tuple[list[str], collections.Counter[str]]:
    """The accepted formulas, and how many were rejected for each reason."""
    accepted: list[str] = []
    rejected: collections.Counter[str] = collections.Counter()
    for formula in formulas:
        reason = reject(formula)
        if reason is None:
            accepted.append(formula)
        else:
            rejected[reason] += 1
    return accepted, rejected


def _outside_strings(formula: str) -> str:
    """The formula with string literals blanked, so a ``|`` inside a string is not DDE."""
    return re.sub(r'"(?:[^"]|"")*"', '""', formula)
