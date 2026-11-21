"""Recalculate Excel workbooks in pure Python and write the results back as cached values."""

from __future__ import annotations

from xlio import LimitError, PackageError

from xlrecalc._check import RecalcCheck
from xlrecalc._conditional import ConditionalFormula
from xlrecalc._evaluate import Formula, Reference, UnsupportedError, evaluate
from xlrecalc._limits import ReadLimits
from xlrecalc._quirks import Quirks
from xlrecalc._recalc import recalculate
from xlrecalc._report import CellRef, Report
from xlrecalc._values import ErrorValue
from xlrecalc._workbook import CellValue, FormulaError, Loader, Workbook, load

__version__ = "0.7.0"

__all__ = [
    "CellRef",
    "CellValue",
    "ConditionalFormula",
    "ErrorValue",
    "Formula",
    "FormulaError",
    "LimitError",
    "Loader",
    "PackageError",
    "Quirks",
    "ReadLimits",
    "RecalcCheck",
    "Reference",
    "Report",
    "UnsupportedError",
    "Workbook",
    "__version__",
    "evaluate",
    "load",
    "recalculate",
]
