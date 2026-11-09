"""The worksheet functions (ARCHITECTURE.md §8). Importing this package registers them all."""

from xlrecalc._functions import (
    _bonds,
    _database,
    _datetime,
    _dynamic,
    _engineering,
    _financial,
    _logical,
    _lookup,
    _math,
    _reference,
    _reshape,
    _statistical,
    _text,
)
from xlrecalc._functions._registry import REGISTRY, SPECS, Spec, function

__all__ = ["REGISTRY", "SPECS", "Spec", "function"]

del (
    _bonds,
    _database,
    _datetime,
    _dynamic,
    _engineering,
    _financial,
    _logical,
    _lookup,
    _math,
    _reference,
    _reshape,
    _statistical,
    _text,
)
