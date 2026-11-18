"""``recalculate``: formulas evaluated and results written back (ARCHITECTURE.md §9.2, §9.3)."""

from __future__ import annotations

import os

import xlio

from xlrecalc._formula._fastkey import translate
from xlrecalc._limits import ReadLimits, io_limits
from xlrecalc._load import RecalcOnLoad, load, recalc_mode
from xlrecalc._quirks import Quirks
from xlrecalc._report import Report
from xlrecalc._save import array_edits, lookups

__all__ = ["recalculate"]


def recalculate(
    source: str | os.PathLike[str],
    destination: str | os.PathLike[str] | None = None,
    *,
    recalc_on_load: RecalcOnLoad = "excel",
    full_calc_on_load: bool = True,
    limits: ReadLimits | None = None,
    quirks: Quirks | None = None,
    full: bool | None = None,
) -> Report:
    """Recalculate a workbook and write the results into the file as its cached values.

    Readers that take cached values, such as pandas and openpyxl's data-only mode, then see the
    results. Formulas are evaluated as Excel evaluates them on opening the file and saving it
    (``recalc_on_load``, as :func:`load`): every one in a file Excel recalculates in full (one
    openpyxl wrote, for instance); otherwise those without a cached value, the cells Excel
    marked ``ca="1"`` (volatile cells and their dependents), and what depends on a changed
    value. In manual calculation it computes what Excel recalculates when it saves, where
    :func:`load` keeps every cached value, as Excel shows them on opening. Only the cached
    values of formula cells whose value changed are rewritten; every other part of the file is
    copied byte for byte.

    Args:
        source: The xlsx file.
        destination: Where to write the result; by default ``source`` itself. Either way the
            file is written to a temporary file next to it and moved into place, so a failure
            never leaves a partial file.
        recalc_on_load: ``"excel"``, Excel's rule (as :func:`load`); ``"always"``, evaluate
            every formula, trusting no cached value; ``"never"``, trust every cached value
            the file holds but the volatile formulas' (a formula without one is evaluated in
            every mode).
        full_calc_on_load: Ask Excel to recalculate when it opens the file
            (``<calcPr fullCalcOnLoad="1">``), so a value computed differently here never
            persists in what Excel shows. With ``False`` the file asks nothing (a request it
            held is removed, as Excel removes it when it saves), and Excel shows the written
            values unless its own rules recalculate (an older ``calcId``, a volatile cell and
            its dependents). A file with the setting also makes a later ``recalculate`` of it
            evaluate every formula again.
        limits: Limits on what reading and computing the workbook may cost (``ReadLimits``):
            the package's parts and sizes, the cells an array may hold, and the evaluations of
            one run of iterative calculation.
        quirks: Which Excel quirks to reproduce; all by default.
        full: Deprecated, to be removed in 1.0: ``True`` is ``recalc_on_load="always"``.

    Returns:
        What could not be done: unsupported functions, uncomputed cells (whose cached values
        are left as they were), cells that keep the cached value of a function or feature not
        implemented (``from_kept``), circular references, and formulas Excel refuses in a file
        (left as written; Excel's repair turns each into its cached value on open).

    Raises:
        PackageError: The file is not a readable xlsx package.
        LimitError: The file exceeds ``limits``.
        OSError: The file cannot be read, or the result cannot be written.
    """
    mode = recalc_mode(recalc_on_load, "full", full, stacklevel=2)
    # A load, then a save: in manual calculation Excel saves recalculated what it showed as cached.
    engine = load(source, quirks=quirks, limits=limits, recalc_on_load=mode, saving=True)
    engine.recalculate(full=mode == "always")
    engine.save_run()  # an opening, then a save: Excel's saved copy (§7.2)
    target = source if destination is None else destination
    edits, fills = array_edits(engine)
    xlio.patch(
        source,
        target,
        cached=lookups(engine),
        cells=edits,
        fills=fills,
        translate=translate,
        full_calc_on_load=full_calc_on_load,
        limits=io_limits(limits),
    )
    return engine.report
