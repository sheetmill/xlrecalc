"""Reading xlsx packages: standard library only, no imports from the rest of xlrecalc.

Internal, and kept apart so it can become a package of its own (ARCHITECTURE.md §9.1).
"""

from xlrecalc._xlsx._book import Book, CalcProperties, DefinedName, SheetInfo, Table
from xlrecalc._xlsx._cells import Cell, Formula, unescape
from xlrecalc._xlsx._package import LimitError, PackageError, ReadLimits

__all__ = [
    "Book",
    "CalcProperties",
    "Cell",
    "DefinedName",
    "Formula",
    "LimitError",
    "PackageError",
    "ReadLimits",
    "SheetInfo",
    "Table",
    "unescape",
]
