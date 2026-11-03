"""The filter that keeps unsafe formulas from downloaded workbooks out of captures."""

import pytest
from capture.corpus_filter import reject, tally


@pytest.mark.parametrize(
    ("formula", "reason"),
    [
        ("cmd|' /C calc'!A0", "dde"),
        ("=cmd|' /C calc'!A0", "dde"),
        ("1+cmd|'/c calc'!'A0'", "dde"),
        ("Excel|'[Book1]Sheet1'!R1C1", "dde"),
        ("msexcel|'\\..\\..\\Windows\\System32\\cmd.exe /c calc'!_xlbgnm.A1", "dde"),
        ("[1]Sheet1!A1", "external"),
        ("'[1]Sheet 1'!$A$1+1", "external"),
        ("[1]!Rate", "external"),
        ("[1]!'NGH1,PRIM ACT 1'", "external"),
        ("SUM([2]Data!A1:A10)", "external"),
        ('INDIRECT("[other.xlsx]Sheet1!A1")', "external"),
        ('WEBSERVICE("https://example.com")', "side effects: WEBSERVICE"),
        ('_xlfn.WEBSERVICE("https://example.com")', "side effects: WEBSERVICE"),
        ('_xlfn.FILTERXML(A1,"//a")', "side effects: FILTERXML"),
        ('RTD("prog",,"topic")', "side effects: RTD"),
        ('CALL("kernel32","GetTickCount","J")', "side effects: CALL"),
        ('REGISTER.ID("kernel32","GetTickCount","J")', "side effects: REGISTER.ID"),
        ('INFO("directory")', "side effects: INFO"),
        ('CELL("filename",A1)', "side effects: CELL"),
        ('_xlfn.IMAGE("https://example.com/a.png")', "side effects: IMAGE"),
        ('_xlfn._xlws.PY(0,"import os")', "side effects: PY"),
        ('HYPERLINK("https://example.com","x")', "hyperlink"),
        ("1+", "unparseable"),
        ("SUM(1", "unparseable"),
    ],
)
def test_rejected(formula, reason):
    assert reject(formula.removeprefix("=")) == reason


@pytest.mark.parametrize(
    "formula",
    [
        "SUM(A1:A10)",
        'IF(A1="a|b",1,2)',
        '"cmd|calc"&A1',
        "Sheet2!A1+'My Sheet'!B2",
        'INDIRECT("A"&ROW())',
        "VLOOKUP(A1,Table1[[#All],[Price]],2,FALSE)",
        "_xlfn.XLOOKUP(A1,B:B,C:C)",
        "Table1[[#This Row],[Price]]",
    ],
)
def test_accepted(formula):
    assert reject(formula) is None


def test_tally():
    accepted, rejected = tally(["SUM(A1)", "[1]S!A1", "[1]S!B1", "cmd|'x'!A0", 'CELL("type")'])
    assert accepted == ["SUM(A1)"]
    assert rejected == {"external": 2, "dde": 1, "side effects: CELL": 1}
