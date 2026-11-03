"""Recalculate Excel workbooks in pure Python and write the results back as cached values."""

from __future__ import annotations

from xlrecalc._quirks import Quirks
from xlrecalc._recalc import recalculate
from xlrecalc._report import CellRef, Report
from xlrecalc._values import ErrorValue
from xlrecalc._workbook import CellValue, FormulaError, Workbook, load
from xlrecalc._xlsx import LimitError, PackageError, ReadLimits

__version__ = "0.1.0"

__all__ = [
    "CellRef",
    "CellValue",
    "ErrorValue",
    "FormulaError",
    "LimitError",
    "PackageError",
    "Quirks",
    "ReadLimits",
    "Report",
    "Workbook",
    "__version__",
    "load",
    "recalculate",
]
