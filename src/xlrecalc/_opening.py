"""Which formulas are volatile, and what opening a workbook computes (capture ``volatile_names``).

One place for both, read by ``load`` and by ``RecalcCheck``, so the two cannot disagree.

A formula is volatile when it calls a volatile function (``RAND``, ``NOW``, ``OFFSET``,
``INDIRECT``...) in its own text, or refers to or calls a defined name that does, directly or
through other names (``N`` = ``OFFSET(...)``, ``M`` = ``N+0``, a ``LAMBDA`` name whose body calls
``NOW()``). A function name inside a string literal is not a call; a name defined but not used
changes nothing. A name counts as volatile when any of its definitions, in any scope, is.

What Excel does on opening a workbook whose calculation properties let it keep the cached values
(capture ``volatile_names``): it recalculates the cells whose formula carries ``ca="1"``, the
mark Excel writes on every volatile cell and on every cell depending on one, and keeps the cached
values of the rest, a volatile cell without the mark included; it recalculates every formula
when a cell calls a volatile function in its own text without the mark (a file another program
wrote). In manual calculation it does this when saving, not on opening.
"""

from __future__ import annotations

import re
from collections.abc import Iterable
from dataclasses import dataclass
from typing import Literal

import xlio

from xlrecalc._formula._tokens import Kind, TokenizeError, tokenize
from xlrecalc._functions import SPECS

__all__ = [
    "VERIFIED_CALC_ID",
    "VOLATILE_FUNCTIONS",
    "Rule",
    "Volatility",
    "recalculates_on_opening",
    "shows_cached",
]

VOLATILE_FUNCTIONS = frozenset(n.upper() for n, s in SPECS.items() if s.volatile)
"""The volatile functions, upper case, with their file prefix where they have one."""

_KNOWN = frozenset(n.upper() for n in SPECS)
_WORD = r"(?<![\w.])(?:_xl[a-z]+\.)?(?:{})(?![\w.])"  # a word, after a file prefix or none


class Volatility:
    """Whether a formula is volatile, given the workbook's defined names.

    Args:
        names: Every defined name as (name, formula text), in any scope.
    """

    def __init__(self, names: Iterable[tuple[str, str]]) -> None:
        texts: dict[str, list[str]] = {}
        for name, text in names:
            texts.setdefault(name.casefold(), []).append(text)
        uses = {name: _uses(bodies) for name, bodies in texts.items()}
        # A name is volatile when a body calls a volatile function or uses a volatile name:
        # a fixed point, reached without recursion (chains of names may be long).
        volatile = {name for name, (calls, _) in uses.items() if calls & VOLATILE_FUNCTIONS}
        dependents: dict[str, set[str]] = {}
        for name, (_, used) in uses.items():
            for other in used:
                dependents.setdefault(other, set()).add(name)
        pending = list(volatile)
        while pending:
            for user in dependents.get(pending.pop(), ()):
                if user not in volatile:
                    volatile.add(user)
                    pending.append(user)
        self.names = frozenset(volatile)
        """The defined names (folded) that are volatile."""
        words = [re.escape(f.rsplit(".", 1)[-1]) for f in VOLATILE_FUNCTIONS]
        words += [re.escape(n) for n in self.names]
        self._own = re.compile(_WORD.format("|".join(words)), re.IGNORECASE)

    def __call__(self, text: str) -> bool:
        """Whether the formula (file form, no leading ``=``) is volatile."""
        if not self._own.search(text):  # most formulas: no word that could make it volatile
            return False
        calls, used = _uses((text,))
        return bool(calls & VOLATILE_FUNCTIONS or used & self.names)

    def own(self, text: str) -> bool:
        """Whether the formula calls a volatile function in its own text."""
        if not self._own.search(text):
            return False
        return bool(_uses((text,))[0] & VOLATILE_FUNCTIONS)


def _uses(texts: Iterable[str]) -> tuple[frozenset[str], frozenset[str]]:
    """The functions (upper case) and the names (folded) the formulas call or refer to.

    A call of a defined name (``F(1)``, a name that is not a function) counts as both. A name
    after an external prefix (``[1]!N``) is another workbook's and is left out. A text that does
    not tokenize uses nothing.
    """
    calls: set[str] = set()
    used: set[str] = set()
    for text in texts:
        try:
            tokens = tokenize(text)
        except TokenizeError:
            continue
        external = False
        for token in tokens:
            if token.kind is Kind.PREFIX:
                external = token.text.startswith("[") or "]" in token.text
                continue
            if token.kind is Kind.FUNCTION:
                calls.add(token.text.upper())
                if token.text.upper() not in _KNOWN:  # a call of a defined name (a LAMBDA)
                    used.add(token.text.casefold())
            elif token.kind is Kind.NAME and not external:
                used.add(token.text.casefold())
            external = False
    return frozenset(calls), frozenset(used)


VERIFIED_CALC_ID = 191028
"""The lowest ``calcId`` at which Excel trusts a file's cached values on opening, as verified on
the newest Excel build captured: Excel 16.115 for Mac (its own ``calcId`` 191029) recalculates a
file of ``calcId`` 191027 or lower and trusts one of 191028 or higher (capture ``open_calc``,
``capture/open_calc_probe.py``, rerun after every Excel update)."""


def recalculates_on_opening(
    calc: xlio.CalcProperties, app_name: str | None, saving: bool = False
) -> bool:
    """Whether the calculation properties make Excel recalculate every formula on opening.

    The file-level part of the rule (capture ``open_calc``): it does when the workbook has no
    ``calcPr`` or one without ``calcId``, asks for it (``fullCalcOnLoad``), was calculated by an
    older Excel (``calcId`` below ``VERIFIED_CALC_ID``), or was not saved by Excel (``app_name``,
    ``<fileVersion appName>``, is not ``"xl"``: openpyxl writes no ``<fileVersion>``, LibreOffice
    ``"Calc"``). In manual calculation it shows the cached values on opening whatever the reason,
    and saves them recalculated where a reason holds: ``saving`` asks about the copy it saves
    after opening. Otherwise the formulas decide (``_volatile``: cells marked ``ca="1"``, volatile
    cells and their dependents; every formula when a volatile function is unmarked).
    """
    if shows_cached(calc, saving):
        return False
    return (
        calc.full_calc_on_load
        or calc.calc_id is None
        or calc.calc_id < VERIFIED_CALC_ID
        or app_name != "xl"
    )


def shows_cached(calc: xlio.CalcProperties, saving: bool) -> bool:
    """Whether Excel shows every cached value on opening: manual calculation, not saving."""
    return calc.calc_mode == "manual" and not saving


Opening = Literal["all", "some", "none"]
"""What a load computes: every formula, some (those without a cached value, the volatile ones,
those Excel recalculates on opening, and what depends on them), or none."""


@dataclass(frozen=True, slots=True)
class Rule:
    """What a load computes for one workbook and ``recalc_on_load``, before its formulas are read.

    Attributes:
        everything: Every formula, from the file alone: ``"always"``, or ``"excel"`` where the
            calculation properties make Excel recalculate on opening.
        volatile: Volatile formulas are computed: ``"never"`` (at every calculation, as Excel
            computes them; ``"excel"`` computes what Excel recalculates on opening instead, and
            ``"always"`` everything).
        marked: Formulas marked ``ca="1"`` are computed: ``"excel"`` where Excel recalculates on
            opening, outside iterative calculation (where the mark is on a loop's cells, which a
            trusting load keeps: captures ``iterate_open``, ``iterate_cached``).
        unmarked: A volatile function in a formula's own text without ``ca="1"`` means every
            formula: ``"excel"`` where Excel recalculates on opening.
    """

    everything: bool
    volatile: bool
    marked: bool
    unmarked: bool

    @classmethod
    def of(cls, calc: xlio.CalcProperties, app_name: str | None, mode: str, saving: bool) -> Rule:
        """The rule for a workbook's calculation properties and ``<fileVersion appName>``."""
        cached = mode == "excel" and shows_cached(calc, saving)
        everything = mode == "always" or (
            mode == "excel" and recalculates_on_opening(calc, app_name, saving)
        )
        excel = mode == "excel" and not cached and not everything
        return cls(everything, mode == "never", excel and not calc.iterate, excel)

    def decide(self, uncached: bool, volatile: bool, marked: bool, unmarked: bool) -> Opening:
        """The answer, from what the formulas showed (each: some formula was so)."""
        if self.everything or (self.unmarked and unmarked):
            return "all"
        if uncached or (self.volatile and volatile) or (self.marked and marked):
            return "some"
        return "none"
