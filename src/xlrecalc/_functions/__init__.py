"""The worksheet functions (ARCHITECTURE.md §8). Importing this package registers them all."""

from collections import ChainMap
from collections.abc import Callable, Mapping
from typing import Any

from xlrecalc._formula._prefixes import PREFIXES
from xlrecalc._functions import (
    _bessel,
    _bonds,
    _complex,
    _continuous,
    _convert,
    _database,
    _datetime,
    _distributions,
    _dynamic,
    _engineering,
    _financial,
    _lambda,
    _logical,
    _lookup,
    _makers,
    _math,
    _reference,
    _regression,
    _reshape,
    _statistical,
    _text,
)
from xlrecalc._functions._registry import REGISTRY, SPECS, Spec, function
from xlrecalc._values import ErrorValue

__all__ = ["BARE_NAMES", "REGISTRY", "SPECS", "Spec", "function", "registry_for"]

del (
    _bessel,
    _bonds,
    _complex,
    _continuous,
    _convert,
    _database,
    _datetime,
    _distributions,
    _dynamic,
    _engineering,
    _financial,
    _lambda,
    _logical,
    _lookup,
    _makers,
    _math,
    _reference,
    _regression,
    _reshape,
    _statistical,
    _text,
)


def _unknown_name(ctx: Any, *args: Any) -> ErrorValue:
    """A call of a name the platform's Excel does not know: ``#NAME?``, whatever its arguments."""
    return ErrorValue.NAME


# The names each platform's Excel does not know, in the file forms captured (function_names,
# encodeurl): they compile to #NAME?. Excel for Windows knows ENCODEURL and FILTERXML only with
# _xlfn. (and computes FILTERXML, and EUROCONVERT with its add-in, which this package does not).
_UNKNOWN: dict[str, tuple[str, ...]] = {
    "mac": (
        "ENCODEURL",
        "_XLFN.ENCODEURL",
        "FILTERXML",
        "_XLFN.FILTERXML",
        "EUROCONVERT",
        "_XLFN.EUROCONVERT",
    ),
    "windows": ("ENCODEURL", "FILTERXML"),
}
# The names Excel writes with a file prefix (_xlfn.XLOOKUP), found without it in a file another
# program wrote: Excel shows #NAME? for each on both platforms (capture bare_names). SORT and FILTER
# without theirs make Excel refuse the file (the formula is refused instead), and LET without its
# prefix was not captured. In a file Excel saved, such a name is a macro's or an add-in's function,
# which Excel computed: it is not one of these (registry_for(platform, bare=False)).
BARE_NAMES: frozenset[str] = frozenset(PREFIXES) - {"SORT", "FILTER", "LET"}
_REGISTRIES: dict[tuple[str, bool], Mapping[str, Callable[..., Any]]] = {}


def registry_for(platform: str, bare: bool = True) -> Mapping[str, Callable[..., Any]]:
    """The functions a formula compiles against: ``REGISTRY`` and the platform's unknown names.

    With ``bare``, the names of ``BARE_NAMES`` are among the unknown ones (``#NAME?``): a file
    another program wrote. Without, as in a file Excel saved, they are not in the registry at all.
    """
    found = _REGISTRIES.get((platform, bare))
    if found is None:
        names = (*_UNKNOWN[platform], *sorted(BARE_NAMES)) if bare else _UNKNOWN[platform]
        unknown = dict.fromkeys(names, _unknown_name)
        found = _REGISTRIES[platform, bare] = ChainMap(unknown, REGISTRY)  # REGISTRY as it grows
    return found
