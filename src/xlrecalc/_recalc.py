"""``recalculate``: a workbook's formulas evaluated and their results written back (§9.2, §9.3)."""

from __future__ import annotations

import os

from xlrecalc._formula._fastkey import translate
from xlrecalc._load import load
from xlrecalc._quirks import Quirks
from xlrecalc._report import Report
from xlrecalc._save import array_edits, lookups
from xlrecalc._xlsx import ReadLimits
from xlrecalc._xlsx._write import write

__all__ = ["recalculate"]


def recalculate(
    source: str | os.PathLike[str],
    destination: str | os.PathLike[str] | None = None,
    *,
    full: bool = False,
    full_calc_on_load: bool = True,
    limits: ReadLimits | None = None,
    quirks: Quirks | None = None,
) -> Report:
    """Recalculate a workbook and write the results into the file as its cached values.

    Readers that take cached values, such as pandas and openpyxl's data-only mode, then see
    the results. Formulas whose cached values the file already holds and trusts (a file Excel
    saved) are not evaluated unless something they depend on is; a file openpyxl wrote holds
    none, so every formula is. Only the cached values of formula cells whose value changed are
    rewritten; every other part of the file is copied byte for byte.

    Args:
        source: The xlsx file.
        destination: Where to write the result; by default ``source`` itself. Either way the
            file is written to a temporary file next to it and moved into place, so a failure
            never leaves a partial file.
        full: Evaluate every formula, trusting no cached value.
        full_calc_on_load: Ask Excel to recalculate when it opens the file
            (``<calcPr fullCalcOnLoad="1">``), so a value computed differently here never
            persists in what Excel shows. With ``False`` the file's setting stays as it was and
            Excel shows the written values. A file with the setting also makes a later
            ``recalculate`` of it evaluate every formula again.
        limits: Limits on what reading the file may decompress.
        quirks: Which Excel quirks to reproduce; all by default.

    Returns:
        What could not be done: unsupported functions, uncomputed cells (whose cached values
        are left as they were), circular references, and formulas Excel refuses in a file
        (left as written; Excel's repair turns each into its cached value on open).

    Raises:
        PackageError: The file is not a readable xlsx package.
        LimitError: The file exceeds ``limits``.
        OSError: The file cannot be read, or the result cannot be written.
    """
    engine = load(source, quirks=quirks, limits=limits, trust_cached=not full)
    engine.recalculate(full=full)
    target = source if destination is None else destination
    edits, fills = array_edits(engine)
    write(
        source,
        target,
        lookups(engine),
        edits=edits,
        fills=fills,
        translate=translate,
        full_calc_on_load=full_calc_on_load,
        limits=limits,
    )
    return engine.report
