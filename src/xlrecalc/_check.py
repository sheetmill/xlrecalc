"""Whether loading a workbook computes any formula, without building an engine."""

from __future__ import annotations

import os
from typing import BinaryIO

import xlio

from xlrecalc._limits import ReadLimits, io_limits
from xlrecalc._load import RECALC_ON_LOAD, RecalcOnLoad, defined_names, file_app
from xlrecalc._opening import Opening, Rule, Volatility

__all__ = ["RecalcCheck"]


class RecalcCheck:
    """Whether ``load`` would compute any formula of a workbook, from what the caller reads.

    For a program that reads a workbook's cells itself and wants Excel's values without
    building an engine where none is needed: the answer is ``load``'s, from the same rule
    (capture ``open_calc`` for the file, ``volatile_names`` for the formulas). Like ``hashlib``'s
    objects, it is fed and then asked:

    >>> check = RecalcCheck("book.xlsx")  # doctest: +SKIP
    >>> if check.result is None:  # doctest: +SKIP
    ...     for formula, cached, ca in formula_cells:  # every worksheet's, from the caller's read
    ...         check.add(formula, cached=cached, ca=ca)
    >>> check.finish()  # doctest: +SKIP
    'none'

    Creating it reads only the workbook part (``calcPr``, ``<fileVersion>``, the defined names),
    which alone decides most files: an older ``calcId``, ``fullCalcOnLoad``, a file another
    program wrote. A caller holding the workbook open already still passes ``source``; the
    workbook part is read again, under a millisecond.

    Args:
        source: The workbook, as a path or a binary file object, as ``load`` takes it.
        recalc_on_load: As ``load``'s: ``"excel"``, ``"always"`` or ``"never"``.
        saving: As ``load``'s: the load is followed by a save, as in ``recalculate(path)``.
        limits: Bounds on what is read, as ``load``'s.

    Raises:
        ValueError: ``recalc_on_load`` is not one of the three.
        PackageError: The file is not a workbook xlrecalc reads.
    """

    def __init__(
        self,
        source: str | os.PathLike[str] | BinaryIO,
        *,
        recalc_on_load: RecalcOnLoad = "excel",
        saving: bool = False,
        limits: ReadLimits | None = None,
    ) -> None:
        if recalc_on_load not in RECALC_ON_LOAD:
            raise ValueError(f"recalc_on_load must be one of {RECALC_ON_LOAD}: {recalc_on_load!r}")
        with xlio.open(source, limits=io_limits(limits)) as book:
            self._rule = Rule.of(book.calc, file_app(book), recalc_on_load, saving)
            names = [(name, text) for name, text, _ in defined_names(book)]
        self._volatility = Volatility(names)
        self._uncached = self._volatile = self._marked = self._unmarked = False
        self._final: Opening | None = None

    def add(self, formula: str | None, *, cached: bool = True, ca: bool = False) -> None:
        """One formula cell of a worksheet.

        Args:
            formula: Its text as the file stores it (no leading ``=``), or ``None`` for a
                dependent cell of a shared formula, whose text is its master's.
            cached: Whether the file holds a value for it (``<v>``).
            ca: Whether its formula carries ``ca="1"``, the mark Excel writes on a cell it
                recalculates on opening.

        Raises:
            ValueError: After ``finish``.
        """
        if self._final is not None:
            raise ValueError("RecalcCheck.add after finish")
        if not cached:
            self._uncached = True
        if ca:
            self._marked = True
        if formula is None:
            return
        if not ca and not self._unmarked and self._volatility.own(formula):
            self._unmarked = self._volatile = True
        elif not self._volatile and self._volatility(formula):
            self._volatile = True

    @property
    def result(self) -> Opening | None:
        """``"all"`` as soon as every formula is certain to be computed, else ``None``.

        After ``finish``, its answer. ``"all"`` needs no more formulas: the caller may stop
        feeding them and call ``finish``.
        """
        if self._final is not None:
            return self._final
        answer = self._decide()
        return answer if answer == "all" else None

    def finish(self) -> Opening:
        """The answer, once every formula cell has been added; ``add`` is refused after it.

        Returns:
            ``"all"``: every formula is computed. ``"some"``: some are (those without a cached
            value, the volatile ones, those Excel recalculates on opening, and what depends on
            them); their values are ``load``'s. ``"none"``: every cached value stands.
        """
        if self._final is None:
            self._final = self._decide()
        return self._final

    def _decide(self) -> Opening:
        return self._rule.decide(self._uncached, self._volatile, self._marked, self._unmarked)
