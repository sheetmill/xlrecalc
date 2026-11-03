"""ErrorValue (ARCHITECTURE.md §4.2)."""

import pickle

import pytest

from xlrecalc._values import ErrorValue


def test_listed_codes_are_singletons():
    assert ErrorValue("#N/A") is ErrorValue.NA
    assert ErrorValue("#DIV/0!") is ErrorValue.DIV0
    assert pickle.loads(pickle.dumps(ErrorValue.NA)) is ErrorValue.NA


def test_unlisted_codes_are_new_values_that_compare_by_code():
    a, b = ErrorValue("#FUTURE!"), ErrorValue("#FUTURE!")
    assert a == b
    assert hash(a) == hash(b)


def test_not_a_str():
    assert not isinstance(ErrorValue.NA, str)
    assert str(ErrorValue.NA) == "#N/A"


def test_equals_its_text_for_interoperation():
    assert ErrorValue.NA == "#N/A"
    assert "#N/A" == ErrorValue.NA  # noqa: SIM300 - the reflected comparison is the point
    assert {ErrorValue.NA: 1}["#N/A"] == 1


def test_equals_xlformat_error():
    xlformat = pytest.importorskip("xlformat")
    assert ErrorValue.NA == xlformat.ExcelError.NA
    assert xlformat.ExcelError.NA == ErrorValue.NA
    assert hash(ErrorValue.NA) == hash(xlformat.ExcelError.NA)


@pytest.mark.parametrize("text", ["", "N/A", "#n/a", "hello", "#"])
def test_refuses_text_that_is_not_a_code(text):
    with pytest.raises(ValueError, match="not an error code"):
        ErrorValue(text)


def test_immutable():
    with pytest.raises(AttributeError):
        ErrorValue.NA.code = "#VALUE!"  # type: ignore[misc]
