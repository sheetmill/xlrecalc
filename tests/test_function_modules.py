"""Every module of ``xlrecalc._functions`` is loaded by ``import xlrecalc``, or listed as not yet.

A module of functions committed without its import in ``_functions/__init__.py`` registers
nothing: its functions stay absent and every capture of them goes unchecked. A module may be
committed before its functions are registered (while their tolerances are measured, say), but
only by name in ``NOT_REGISTERED``, and a listed module that the package loads fails too.
"""

from __future__ import annotations

import pkgutil
import subprocess
import sys
from pathlib import Path

import xlrecalc._functions as functions

# Modules committed before their functions are registered; each is imported by the package once
# they are, and taken off this list in the same change.
NOT_REGISTERED: frozenset[str] = frozenset({"_complex"})

_PREFIX = "xlrecalc._functions."


def _modules() -> set[str]:
    return {m.name for m in pkgutil.iter_modules([str(Path(functions.__file__).parent)])}


def _loaded_by_import() -> set[str]:
    """The function modules a fresh interpreter has loaded after ``import xlrecalc`` alone."""
    code = (
        "import sys, xlrecalc\n"
        f"print(' '.join(m[{len(_PREFIX)}:] for m in sys.modules if m.startswith({_PREFIX!r})))"
    )
    out = subprocess.run([sys.executable, "-c", code], capture_output=True, text=True, check=True)
    return set(out.stdout.split())


def test_every_function_module_is_loaded_or_listed():
    loaded = _loaded_by_import()
    assert not _modules() - loaded - NOT_REGISTERED, "imported by nothing and not listed"
    assert not NOT_REGISTERED & loaded, "listed as not registered, but the package loads it"
    assert _modules() >= NOT_REGISTERED, "listed, but no such module"
