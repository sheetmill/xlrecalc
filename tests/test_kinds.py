"""The text "#N/A" and the error #N/A stay distinct everywhere the engine compares (§4.2).

ErrorValue equals its text under Python's ``==`` for interoperation with xlformat; these tests
check that no comparison in the engine relies on that. The rows captured from Excel for the same
questions are in tests/captures/compare.tsv (tags text_vs_error and text_or_error).
"""

from xlrecalc._compare import compare, criterion, lookup_equal, lookup_key
from xlrecalc._values import BLANK, ErrorValue

NA = ErrorValue.NA


def test_python_equality_is_only_for_interoperation():
    assert NA == "#N/A"  # the documented interoperation
    assert not isinstance(NA, str)


def test_operators():
    assert compare("#N/A", "#N/A", "=") is True
    assert compare(NA, "#N/A", "=") is NA  # an error operand gives the error
    assert compare("#N/A", NA, "<>") is NA


def test_exact_lookup():
    assert lookup_equal("#N/A", "#N/A")
    assert not lookup_equal("#N/A", NA)
    assert not lookup_equal(1.0, True)
    assert not lookup_equal("", BLANK)


def test_criteria():
    # A criterion spelled like an error code matches the error, never the text (captured).
    assert criterion("#N/A")(NA)
    assert not criterion("#N/A")("#N/A")
    assert criterion(NA)(NA)
    assert not criterion(NA)("#N/A")
    assert criterion("#*")("#N/A")
    assert not criterion("#*")(NA)


def test_lookup_keys():
    keys = {lookup_key("#N/A"): "text", lookup_key(NA): "error"}
    assert len(keys) == 2
    assert lookup_key("ABC") == lookup_key("abc")
    assert lookup_key(1.0) != lookup_key(True)
    assert lookup_key(0.0) != lookup_key(BLANK) != lookup_key("")
