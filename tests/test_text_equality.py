"""Which equality of text each kind of comparison uses on the Mac (capture ``collate_functions``).

The operators' ``=`` folds ``ß`` to ``ss``; a criterion's equality and a binary search follow it,
an exact lookup's scan ignores case only, and a binary search finds the first of equal entries.
"""

from xlrecalc import evaluate
from xlrecalc._compare import compare, criterion, lookup_equal, lookup_key


def test_operators_fold_sharp_s():
    assert compare("ß", "ss", "=", platform="mac") is True
    assert compare("ae", "æ", "=", platform="mac") is False


def test_criterion_equality_is_the_operators():
    assert criterion("ß", platform="mac")("ss")
    assert criterion("ss", platform="mac")("ß")
    assert criterion("SS", platform="mac")("ß")
    assert not criterion("ae", platform="mac")("æ")
    assert not criterion("<>ß", platform="mac")("ss")


def test_exact_lookup_ignores_case_only():
    assert lookup_equal("A", "a")
    assert not lookup_equal("ß", "ss")
    assert lookup_key("ß") != lookup_key("ss")
    assert lookup_key("ABC") == lookup_key("abc")
    assert evaluate('MATCH("ß",{"ss";"ß"},0)', platform="mac") == 2.0
    assert evaluate('_xlfn.XLOOKUP("ß",{"ss";"ß"},{1;2})', platform="mac") == 2.0


# SORT of the table of capture collate_functions, as Excel gave it
SORTED = (
    "-a", "a", "A", "a b", "a_b", "a-b", "a10", "a9", "aa", "ab", "ae", "æ", "b", "co-op", "co'op",
    "coop", "naive", "naïve", "oe", "œ", "resume", "résumé", "ss", "ß",
)  # fmt: skip


def _xmatch(key: str, words: tuple[str, ...], search: int) -> object:
    array = ";".join(f'"{w}"' for w in words)
    return evaluate(f'_xlfn.XMATCH("{key}",{{{array}}},0,{search})', platform="mac")


def test_binary_search_finds_the_first_equal_entry():
    # Excel's values (capture collate_functions); the bisection reaches "A" and "ß" first
    assert _xmatch("a", SORTED, 2) == 2.0
    assert _xmatch("A", SORTED, 2) == 2.0
    assert _xmatch("ss", SORTED, 2) == 23.0
    assert _xmatch("ß", SORTED, 2) == 23.0
    assert _xmatch("a", SORTED[::-1], -2) == 22.0  # our choice: the first in the file's order
