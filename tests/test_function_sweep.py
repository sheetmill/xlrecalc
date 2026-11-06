"""Every function over a whole sheet of two cells, generated from the registry.

A small file must not make any function use time or memory in proportion to a range's size.
Each function runs in a child process under the memory watchdog, with a time limit, in each of
the forms ``forms`` makes: past the limit a formula keeps the value the file holds and is reported
(ARCHITECTURE.md §7.7), and no function may raise.
"""

from __future__ import annotations

from concurrent.futures import ThreadPoolExecutor
from xml.sax.saxutils import escape

import pytest

from tests.test_array_limit import recalculated, two_sheets
from tests.watchdog import GIB, bounded, in_child
from xlrecalc._engine import Engine
from xlrecalc._functions._registry import SPECS, Spec

WHOLE = "D!A:XFD"
SCALAR = {"number": "1", "text": '"a"', "logical": "TRUE", "value": "1"}
BLANK_MATCHING = ['""', '"<>5"']
REFERENCE_KINDS = ("raw", "array", "matrix", "branch")
NAMES = sorted(
    SPECS
)  # at import: a test elsewhere registers a function of its own, in this process


def _argument_lists(spec: Spec) -> list[list[str]]:
    """The kinds to fill for each length: the fewest, every declared one, a variadic's repeats.

    A function taking pairs refuses an odd count at once, so lengths of both parities are tried.
    """
    most = len(spec.kinds) if spec.max_args is not None else len(spec.kinds) + 3
    lengths = range(max(spec.min_args, 1), max(most, spec.min_args, 1) + 1)
    return [[spec.kinds[min(i, len(spec.kinds) - 1)] for i in range(n)] for n in lengths]


def forms(name: str) -> list[tuple[str, str | None]]:
    """(formula, declared array range or None) calling a function over a whole sheet.

    The ranges go where a reference is taken, everywhere at once, then in the first argument
    alone, then with the later ones matching blank cells; each also as an array formula.
    """
    spec = SPECS[name]
    calls: list[list[str]] = []
    for kinds in _argument_lists(spec) if spec.kinds else []:
        takes = [k in REFERENCE_KINDS for k in kinds]
        plain = [SCALAR.get(k, "1") for k in kinds]
        calls.append([WHOLE] * len(kinds))
        calls.append([WHOLE if t else v for t, v in zip(takes, plain, strict=True)])
        calls.append([WHOLE, *plain[1:]])
        for criterion in BLANK_MATCHING:
            later = zip(takes[1:], plain[1:], strict=True)
            calls.append([WHOLE, *(criterion if t else v for t, v in later)])
            calls.append([WHOLE if i % 2 == 0 else criterion for i in range(len(kinds))])
            calls.append([WHOLE, *(WHOLE if i % 2 else criterion for i in range(len(kinds) - 1))])
    unique = dict.fromkeys(f"{name}({','.join(args)})" for args in calls)
    return [(formula, form) for formula in unique for form in (None, "A1:A1")]


def evaluate(name: str) -> list[tuple[str, str]]:
    """Each form's (formula, the exception a function raised, or ""), in this process."""
    out = []
    for formula, array in forms(name):
        report = recalculated(two_sheets(escape(formula), array=array)).report
        out.append((formula, "; ".join(report.exceptions.values())))
    return out


def outcome(name: str) -> str:
    """What went wrong for a function, or "": a killed or failed child, or an exception."""
    try:
        results = in_child(evaluate, name, limit=GIB, seconds=30)
    except AssertionError as exc:
        return f"{name}: {str(exc)[:300]}"
    return "\n".join(f"{formula}: {why}" for formula, why in results if why)


def _sweep(names: list[str]) -> None:
    with ThreadPoolExecutor(4) as pool:  # at most 4 GiB together, past which the session stops
        assert [line for line in pool.map(outcome, names) if line] == []


def test_functions_taking_a_reference_over_a_whole_sheet():
    names = [n for n in NAMES if set(SPECS[n].kinds) & set(REFERENCE_KINDS)]
    assert len(names) > 50
    _sweep(names)


@pytest.mark.slow
def test_the_other_functions_over_a_whole_sheet():
    names = [n for n in NAMES if not set(SPECS[n].kinds) & set(REFERENCE_KINDS)]
    assert len(names) > 50
    _sweep(names)


@pytest.mark.parametrize(
    "formula",
    [
        "COMBIN(1E300,5E299)",
        "COMBIN(2147483647,1073741823)",
        "PERMUT(2147483647,2147483647)",
        "_xlfn.COMBINA(2147483647,2147483647)",
        "_xlfn.PERMUTATIONA(2147483647,2147483647)",
        "FACT(1E300)",
        "LCM(9007199254740992,9007199254740991,9007199254740989)",
        "MULTINOMIAL(1E15,1E15)",
        '_xlfn.DECIMAL(REPT("Z",255),36)',
        "_xlfn.BASE(9007199254740991,2,255)",
    ],
)
@bounded
def test_integer_functions_stay_bounded_on_huge_arguments(formula: str) -> None:
    """They use exact integers, which a formula must not be able to make huge."""
    engine = Engine.from_cells({"S": {"A1": "=" + formula}})
    assert engine.get(0, 0, 0) is not None
