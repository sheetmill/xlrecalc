"""Cases for ``capture/dynamic_probe.py``: the functions that make arrays.

Each case is a dynamic-array formula in file form at the anchor of its own block of ``BLOCK``
rows in column B; it may put constants in its way (``way``: offsets from the anchor), read it
from plain formulas in column N (``readers``: ``{a}`` is the anchor), or be written as a plain
formula without the dynamic-array marker (``legacy``). The data, in columns AA on, holds a
first and a second set of values for each kind of input. New cases go at the end of a suite.
"""

from __future__ import annotations

from dataclasses import dataclass, field

from xlrecalc._values import ErrorValue

BLOCK = 12
NA, DIV0 = ErrorValue.NA, ErrorValue.DIV0

SEQ = "_xlfn.SEQUENCE"
UNQ = "_xlfn.UNIQUE"
SRT = "_xlfn._xlws.SORT"
SBY = "_xlfn.SORTBY"
FLT = "_xlfn._xlws.FILTER"
ANC = "_xlfn.ANCHORARRAY"


@dataclass(frozen=True)
class DynCase:
    formula: str
    way: dict[tuple[int, int], object] = field(default_factory=dict)
    readers: tuple[str, ...] = ()
    legacy: bool = False


# Data: column letter -> values from row 1 (None is a blank; "=..." is a plain formula).
DATA: dict[str, list[object]] = {
    # first set
    "AA": [3.0, 1.0, 4.0, 1.0, 5.0, 9.0, 2.0, 6.0],
    "AB": ["pear", "Apple", "apple", "Pear", "fig", "APPLE", "kiwi", "fig"],
    "AC": [2.0, "b", True, NA, None, "A", -1.0, False, "a", 10.0, DIV0, '=""'],
    "AD": ["b", "a", "c", "a", "b", "a"],
    "AE": [1.0, 2.0, 3.0, 4.0, 5.0, 6.0],
    "AF": ["1", 1.0, "1", 2.0, True, "TRUE", 1.0],
    # a table by rows for by_col: AG1:AL2
    "AG": [3.0, 1.0],
    "AH": ["x", 2.0],
    "AI": [1.0, 1.0],
    "AJ": [3.0, 1.0],
    "AK": ["X", 2.0],
    "AL": [2.0, 3.0],
    # matrices: AN1:AO2, AP1:AR2, AS1:AS3, and 2x2 ones holding a text, a blank, a logical,
    # an error, a number as text
    "AN": [1.0, 3.0],
    "AO": [2.0, 4.0],
    "AP": [1.0, -1.0],
    "AQ": [0.0, 3.0],
    "AR": [2.0, 1.0],
    "AS": [2.0, 0.5, -1.0],
    "AT": [1.0, 3.0],
    "AU": ["x", 4.0],
    "AV": [1.0, 3.0],
    "AW": [None, 4.0],
    "AX": [1.0, 3.0],
    "AY": [True, 4.0],
    "AZ": [1.0, 3.0],
    "BA": [NA, 4.0],
    "BB": ["1", 3.0],
    "BC": [2.0, 4.0],
    # second set
    "BE": [0.5, -2.0, 7.0, 7.0, 0.0, -2.0, 1e10, 3.25, 0.5, -0.1],
    "BF": ["Zeta", "alpha", "Beta", "zeta", "ALPHA", "beta", "gamma", "Gamma", "delta", "Delta"],
    "BG": [2.0, 1.0, 2.0, 3.0, 1.0, 2.0],
    "BH": ["p", "q", "r", "s", "t", "u"],
    "BJ": [2.0, -1.0, 0.0],
    "BK": [-1.0, 2.0, -1.0],
    "BL": [0.0, -1.0, 2.0],
    "BM": [0.1, 0.2, 0.3],
    "BN": [0.3],
    "BO": [0.2],
    "BP": [0.1],
    "BR": [True, "x", 5.0, None, NA, '=""', -3.0, False, "X", 5.0],
    # tables for the reshaping functions: BT1:BV4 (first set), BX1:CA3 (second set)
    "BT": [1.0, 2.0, 3.0, 4.0],
    "BU": ["a", None, "c", '=""'],
    "BV": [True, NA, False, 0.5],
    "BX": [-1.5, 0.0, 1e10],
    "BY": ["x", DIV0, False],
    "BZ": [7.0, "z", None],
    "CA": ["Y", 8.0, "w"],
}


READ = (f"SUM({ANC}({{a}}))", f"ROWS({ANC}({{a}}))", f"COLUMNS({ANC}({{a}}))")

SEQUENCE = [
    DynCase(f"{SEQ}(3)", readers=READ),
    DynCase(f"{SEQ}(3,2)"),
    DynCase(f"{SEQ}(2,3,10,5)"),
    DynCase(f"{SEQ}(4,1,-1,-0.5)"),
    DynCase(f"{SEQ}(1,5,0,0)"),
    DynCase(f"{SEQ}(2.7)"),
    DynCase(f"{SEQ}(3,2.9)"),
    DynCase(f"{SEQ}(0)"),
    DynCase(f"{SEQ}(-1)"),
    DynCase(f'{SEQ}("3")'),
    DynCase(f"{SEQ}(TRUE)"),
    DynCase(f"{SEQ}(3,,5)"),
    DynCase(f"{SEQ}(,3)"),
    DynCase(f"{SEQ}(3,1,AA1)"),
    DynCase(f"{SEQ}(AA2:AA3)"),
    DynCase(f"{SEQ}({{2,3}})"),
    DynCase(f"{SEQ}(1048577)"),
    DynCase(f"{SEQ}(2,2,9E+307,9E+307)"),
    DynCase(f'{SEQ}(3,1,"a")'),
    DynCase(f"{SEQ}(#N/A)"),
    DynCase(f"{SEQ}(5)", way={(2, 0): "in the way"}, readers=READ),
    DynCase(f"{SEQ}(3)*2"),
    DynCase(f"SUM({SEQ}(10))"),
    DynCase(f"{SEQ}(3)", legacy=True),
    DynCase("SEQUENCE(3)"),
    # second set
    DynCase(f"{SEQ}(5,1,100,-7)", readers=READ),
    DynCase(f"{SEQ}(2,4,0.5,0.25)"),
    DynCase(f"{SEQ}(1,1,7,3)"),
    DynCase(f"{SEQ}(2,3)", way={(1, 2): 0.0}),
    DynCase(f"{SEQ}(0.5)"),
    DynCase(f"{SEQ}(2,0)"),
    DynCase(f"{SEQ}(3,1,1,0.1)"),
]

MMULT = [
    DynCase("MMULT(AN1:AO2,AN1:AO2)", readers=READ),
    DynCase("MMULT(AN1:AO2,AP1:AR2)"),
    DynCase("MMULT(AP1:AR2,AN1:AO2)"),
    DynCase("MMULT(AP1:AR2,AS1:AS3)"),
    DynCase("MMULT({1,2},{3;4})"),
    DynCase("MMULT({1,2;3,4},{5;6})"),
    DynCase("MMULT(AT1:AU2,AN1:AO2)"),
    DynCase("MMULT(AV1:AW2,AN1:AO2)"),
    DynCase("MMULT(AX1:AY2,AN1:AO2)"),
    DynCase("MMULT(AZ1:BA2,AN1:AO2)"),
    DynCase("MMULT(BB1:BC2,AN1:AO2)"),
    DynCase('MMULT({"1",2;3,4},{1;1})'),
    DynCase("MMULT({1,TRUE},{1;1})"),
    DynCase("MMULT(2,3)"),
    DynCase("MMULT(AN1,AO1)"),
    DynCase("MMULT({9E+307,9E+307},{10;10})"),
    DynCase("MMULT(AN1:AO2,AN1:AO2)", way={(1, 1): "x"}),
    DynCase("MMULT(AN1:AO2,AN1:AO2)", legacy=True),
    DynCase("SUM(MMULT(AN1:AO2,AN1:AO2))"),
    DynCase("MMULT(#N/A,AN1:AO2)"),
    DynCase("MMULT(AN1:AO2,AS1:AS3)"),
    # second set
    DynCase("MMULT(BJ1:BL3,BJ1:BL3)", readers=READ),
    DynCase("MMULT(BJ1:BL3,BM1:BM3)"),
    DynCase("MMULT(BN1:BP1,BM1:BM3)"),
    DynCase("MMULT(BM1:BM3,BN1:BP1)"),
    DynCase("MMULT(BN1:BP1,BJ1:BL3)"),
    DynCase("MMULT({0.1,0.2,0.3},{0.3;0.2;0.1})"),
    DynCase("MMULT(BJ1:BL3,BN1:BP1)"),
]

UNIQUE = [
    DynCase(f"{UNQ}(AA1:AA8)", readers=READ),
    DynCase(f"{UNQ}(AB1:AB8)"),
    DynCase(f"{UNQ}(AA1:AA8,,TRUE)"),
    DynCase(f"{UNQ}(AB1:AB8,FALSE,TRUE)"),
    DynCase(f"{UNQ}(AD1:AE6)"),
    DynCase(f"{UNQ}(AD1:AD6)"),
    DynCase(f"{UNQ}(AG1:AL2,TRUE)"),
    DynCase(f"{UNQ}(AG1:AL1,TRUE)"),
    DynCase(f"{UNQ}(AG1:AL1,TRUE,TRUE)"),
    DynCase(f"{UNQ}(AC1:AC12)"),
    DynCase(f"{UNQ}(AF1:AF7)"),
    DynCase(f"{UNQ}({{1,1,2}})"),
    DynCase(f"{UNQ}({{1;1;2}})"),
    DynCase(f"{UNQ}({{1,1,2}},TRUE)"),
    DynCase(f"{UNQ}({{1;1;2;2}},,TRUE)"),
    DynCase(f'{UNQ}(AA1:AA8,"x")'),
    DynCase(f"{UNQ}(5)"),
    DynCase(f"{UNQ}(AA1:AA8)", way={(3, 0): 1.0}),
    DynCase(f"{UNQ}(AA1:AA8)", legacy=True),
    DynCase(f"ROWS({UNQ}(AB1:AB8))"),
    DynCase(f"{UNQ}(AA1:AA8,1,0)"),
    # second set
    DynCase(f"{UNQ}(BE1:BE10)", readers=READ),
    DynCase(f"{UNQ}(BF1:BF10)"),
    DynCase(f"{UNQ}(BF1:BF10,,TRUE)"),
    DynCase(f"{UNQ}(BG1:BH6)"),
    DynCase(f"{UNQ}(BG1:BG6,,TRUE)"),
    DynCase(f"{UNQ}(BR1:BR10)"),
    DynCase(f"{UNQ}(BR1:BR10,,TRUE)"),
    DynCase("UNIQUE(AA1:AA8)"),
]

SORT = [
    DynCase(f"{SRT}(AA1:AA8)", readers=READ),
    DynCase(f"{SRT}(AA1:AA8,,-1)"),
    DynCase(f"{SRT}(AB1:AB8)"),
    DynCase(f"{SRT}(AC1:AC12)"),
    DynCase(f"{SRT}(AC1:AC12,,-1)"),
    DynCase(f"{SRT}(AD1:AE6)"),
    DynCase(f"{SRT}(AD1:AE6,2,-1)"),
    DynCase(f"{SRT}(AD1:AE6,{{1,2}},{{1,-1}})"),
    DynCase(f"{SRT}(AD1:AE6,{{1,2}},-1)"),
    DynCase(f"{SRT}(AG1:AL2,1,1,TRUE)"),
    DynCase(f"{SRT}(AG1:AL2,2,-1,TRUE)"),
    DynCase(f"{SRT}(AD1:AE6,3)"),
    DynCase(f"{SRT}(AD1:AE6,0)"),
    DynCase(f"{SRT}(AD1:AE6,1,0)"),
    DynCase(f"{SRT}(AD1:AE6,1,2)"),
    DynCase(f'{SRT}(AD1:AE6,"1")'),
    DynCase(f"{SRT}({{3,1,2}})"),
    DynCase(f"{SRT}({{3,1,2}},1,1,TRUE)"),
    DynCase(f"{SRT}({{3;1;2}})"),
    DynCase(f"{SRT}(AF1:AF7)"),
    DynCase(f"{SRT}(5)"),
    DynCase(f"{SRT}(AA1:AA8)", way={(4, 0): "x"}),
    DynCase(f"{SRT}(AA1:AA8)", legacy=True),
    DynCase(f"{SRT}(AD1:AE6,1.9)"),
    # second set
    DynCase(f"{SRT}(BE1:BE10)", readers=READ),
    DynCase(f"{SRT}(BF1:BF10)"),
    DynCase(f"{SRT}(BF1:BF10,,-1)"),
    DynCase(f"{SRT}(BG1:BH6)"),
    DynCase(f"{SRT}(BG1:BH6,1,-1)"),
    DynCase(f"{SRT}(BR1:BR10)"),
    DynCase(f"{SRT}(BR1:BR10,,-1)"),
]

SORTBY = [
    DynCase(f"{SBY}(AE1:AE6,AD1:AD6)", readers=READ),
    DynCase(f"{SBY}(AE1:AE6,AD1:AD6,-1)"),
    DynCase(f"{SBY}(AD1:AE6,AD1:AD6,1,AE1:AE6,-1)"),
    DynCase(f"{SBY}(AD1:AE6,AD1:AD6,-1,AE1:AE6,1)"),
    DynCase(f"{SBY}(AB1:AB8,AA1:AA8)"),
    DynCase(f"{SBY}(AA1:AA8,AB1:AB8)"),
    DynCase(f"{SBY}(AG1:AL2,AG2:AL2)"),
    DynCase(f"{SBY}(AG1:AL2,AG1:AL1,-1)"),
    DynCase(f"{SBY}(AE1:AE6,AD1:AD5)"),
    DynCase(f"{SBY}(AE1:AE6,AG1:AL1)"),
    DynCase(f"{SBY}(AE1:AE6,{{3;1;2;6;5;4}})"),
    DynCase(f"{SBY}({{10;20;30}},{{3;1;2}})"),
    DynCase(f"{SBY}(AE1:AE6,AD1:AD6,0)"),
    DynCase(f"{SBY}(AE1:AE6,AD1:AD6,2)"),
    DynCase(f"{SBY}(AC1:AC12,AC1:AC12)"),
    DynCase(f"{SBY}(AE1:AE6,AD1:AD6)", way={(1, 0): 0.0}),
    DynCase(f"{SBY}(AE1:AE6,AD1:AD6)", legacy=True),
    DynCase(f"{SBY}(AE1:AE6,AD1:AE6)"),
    # second set
    DynCase(f"{SBY}(BH1:BH6,BG1:BG6)", readers=READ),
    DynCase(f"{SBY}(BH1:BH6,BG1:BG6,-1)"),
    DynCase(f"{SBY}(BF1:BF10,BE1:BE10)"),
    DynCase(f"{SBY}(BE1:BE10,BF1:BF10,-1)"),
    DynCase(f"{SBY}(BR1:BR10,BR1:BR10,-1)"),
    DynCase(f"{SBY}(BG1:BH6,BG1:BG6,1,BH1:BH6,-1)"),
    DynCase("SORTBY(AE1:AE6,AD1:AD6)"),
]

FILTER = [
    DynCase(f"{FLT}(AD1:AE6,AE1:AE6>2)", readers=READ),
    DynCase(f'{FLT}(AE1:AE6,AD1:AD6="a")'),
    DynCase(f"{FLT}(AE1:AE6,AE1:AE6>10)"),
    DynCase(f'{FLT}(AE1:AE6,AE1:AE6>10,"none")'),
    DynCase(f"{FLT}(AE1:AE6,AE1:AE6>10,{{1,2}})"),
    DynCase(f"{FLT}(AE1:AE6,AE1:AE6>10,AA1:AA3)"),
    DynCase(f"{FLT}(AE1:AE6,AE1:AE6-3)"),
    DynCase(f"{FLT}(AE1:AE6,AD1:AD6)"),
    DynCase(f"{FLT}(AA1:AA8,AC1:AC8)"),
    DynCase(f"{FLT}(AE1:AE6,AE1:AE5>2)"),
    DynCase(f"{FLT}(AG1:AL2,AG2:AL2=1)"),
    DynCase(f"{FLT}(AG1:AL2,AG1:AL1=3)"),
    DynCase(f"{FLT}({{1;2;3}},{{TRUE;FALSE;TRUE}})"),
    DynCase(f"{FLT}({{1,2,3}},{{1,0,1}})"),
    DynCase(f"{FLT}(AD1:AE6,AD1:AE6=1)"),
    DynCase(f"{FLT}(AE1:AE6,TRUE)"),
    DynCase(f'{SRT}({FLT}(AE1:AE6,AD1:AD6<>"b"),,-1)'),
    DynCase(f'{FLT}(AE1:AE6,(AD1:AD6="a")*(AE1:AE6>2))'),
    DynCase(f"{FLT}(AE1:AE6,AE1:AE6>2)", way={(1, 0): "x"}),
    DynCase(f"{FLT}(AE1:AE6,AE1:AE6>2)", legacy=True),
    DynCase(f"{FLT}(AE1:AE6,AE1:AE6>10,#N/A)"),
    DynCase(f'{FLT}(AE1:AE6,AE1:AE6>10,"")'),
    # second set
    DynCase(f"{FLT}(BG1:BH6,BG1:BG6=2)", readers=READ),
    DynCase(f"{FLT}(BF1:BF10,BE1:BE10>0)"),
    DynCase(f"{FLT}(BE1:BE10,BE1:BE10>1E11)"),
    DynCase(f"{FLT}(BE1:BE10,BE1:BE10>1E11,0)"),
    DynCase(f"{FLT}(BR1:BR10,BE1:BE10>0)"),
    DynCase(f"{FLT}(BE1:BE10,BR1:BR10)"),
]

# The reshaping functions, over BT1:BV4 (T1, first set) and BX1:CA3 (T2, second set).
T1, T2 = "BT1:BV4", "BX1:CA3"
CC, CR = "_xlfn.CHOOSECOLS", "_xlfn.CHOOSEROWS"
TK, DR, EX = "_xlfn.TAKE", "_xlfn.DROP", "_xlfn.EXPAND"
VS, HS = "_xlfn.VSTACK", "_xlfn.HSTACK"
TC, TR = "_xlfn.TOCOL", "_xlfn.TOROW"
WR, WC = "_xlfn.WRAPROWS", "_xlfn.WRAPCOLS"

CHOOSE = [
    DynCase(f"{CC}({T1},1)", readers=READ),
    DynCase(f"{CC}({T1},1,3)"),
    DynCase(f"{CC}({T1},-1)"),
    DynCase(f"{CC}({T1},-3)"),
    DynCase(f"{CC}({T1},0)"),
    DynCase(f"{CC}({T1},4)"),
    DynCase(f"{CC}({T1},-4)"),
    DynCase(f"{CC}({T1},{{1,1,-1}})"),
    DynCase(f"{CC}({T1},{{2;3}})"),
    DynCase(f"{CC}({T1},2.9)"),
    DynCase(f"{CC}({T1},-1.9)"),
    DynCase(f'{CC}({T1},"2")'),
    DynCase(f"{CC}({T1},TRUE)"),
    DynCase(f"{CC}({T1},#N/A)"),
    DynCase(f"{CC}({T1},)"),
    DynCase(f"{CC}({T1},BT1:BT2)"),
    DynCase(f"{CC}({T1},BT1:BT2)", legacy=True),
    DynCase(f"{CC}({T1},1)", legacy=True),
    DynCase(f"{CC}(5,1)"),
    DynCase(f"{CC}({{1,2,3;4,5,6}},{{3,1}})"),
    DynCase(f"{CC}({T1},1,{{2,3}})"),
    DynCase(f"{CC}({T1},{{1,0}})"),
    DynCase(f"{CC}({T1},{{1,#N/A}})"),
    DynCase(f"{CC}({T1},BU2)"),
    DynCase(f"{CC}({T1},2)", way={(1, 0): "in the way"}),
    DynCase(f"CHOOSECOLS({T1},1)"),
    DynCase(f'{CC}({T1},"x")'),
    DynCase(f"{CC}({T1},BT1)"),
    DynCase(f"SUM({CC}({T1},1))"),
    DynCase(f"{CR}({T1},1)", readers=READ),
    DynCase(f"{CR}({T1},2)"),
    DynCase(f"{CR}({T1},-1)"),
    DynCase(f"{CR}({T1},{{1;1;-1}})"),
    DynCase(f"{CR}({T1},{{2,4}})"),
    DynCase(f"{CR}({T1},0)"),
    DynCase(f"{CR}({T1},5)"),
    DynCase(f"{CR}({T1},-5)"),
    DynCase(f"{CR}({T1},1.5)"),
    DynCase(f"{CR}({T1},-1.5)"),
    DynCase(f"{CR}({T1},0.5)"),
    DynCase(f'{CR}("x",1)'),
    DynCase(f'{CR}("x",2)'),
    DynCase(f"{CR}({T1},{{1,#N/A}})"),
    DynCase(f"{CR}({T1},BT1:BT2)"),
    DynCase(f"{CR}({T1},3)", legacy=True),
    DynCase(f"CHOOSEROWS({T1},1)"),
    # second set
    DynCase(f"{CC}({T2},-2,2)", readers=READ),
    DynCase(f"{CR}({T2},3,1)"),
    DynCase(f"{CR}({T2},-3)"),
    DynCase(f"{CC}({T2},{{4,4}})"),
    DynCase(f"{CR}({T2},{{2;2}},1)"),
    DynCase(f"{CC}({T2},5)"),
    DynCase(f"{CR}({T2},4)"),
    DynCase(f"{CC}(BX1:BX3,1,1,1)"),
    DynCase(f"{CR}({{1;2;3}},-1,-2,-3)"),
    DynCase(f"{CC}({T2},BY2)"),
    # follow-up: plain formulas over whole columns, lifting, direct errors
    DynCase(f"{CC}(BT:BT,1)", legacy=True),
    DynCase(f"{CR}(BT:BT,1)", legacy=True),
    DynCase(f"{CC}({T1},{{1,2;3,1}})"),
    DynCase(f"{CC}(#N/A,1)"),
    DynCase(f"{CR}({{1,#N/A;3,4}},2)"),
    DynCase(f"{CC}(BY2,1)"),
    # follow-up: an empty array argument
    DynCase(f"{CC}(,1)"),
    DynCase(f"{CR}(,-1)"),
]

DROPTAKE = [
    DynCase(f"{TK}({T1},2)", readers=READ),
    DynCase(f"{TK}({T1},-2)"),
    DynCase(f"{TK}({T1},,2)"),
    DynCase(f"{TK}({T1},2,-1)"),
    DynCase(f"{TK}({T1},0)"),
    DynCase(f"{TK}({T1},10)"),
    DynCase(f"{TK}({T1},-10,-10)"),
    DynCase(f"{TK}({T1},1.9)"),
    DynCase(f'{TK}({T1},"2")'),
    DynCase(f"{TK}({T1},TRUE)"),
    DynCase(f"{TK}({T1},#N/A)"),
    DynCase(f"{TK}({T1},{{1,2}})"),
    DynCase(f"{TK}({T1},,)"),
    DynCase(f"{TK}(5,1)"),
    DynCase(f"{TK}({{1,2,3;4,5,6}},-1,2)"),
    DynCase(f"{TK}({T1},1,0)"),
    DynCase(f"{TK}({T1},-1.9)"),
    DynCase(f"{TK}({T1},BT1:BT2)"),
    DynCase(f"{TK}({T1},2)", legacy=True),
    DynCase(f"{TK}({T1},BT1:BT2)", legacy=True),
    DynCase(f"TAKE({T1},1)"),
    DynCase(f"{TK}({T1},2)", way={(0, 1): "in the way"}),
    DynCase(f'{TK}({T1},"x")'),
    DynCase(f"{TK}({T1},1048577)"),
    DynCase(f"{DR}({T1},2)", readers=READ),
    DynCase(f"{DR}({T1},-2)"),
    DynCase(f"{DR}({T1},,1)"),
    DynCase(f"{DR}({T1},1,-1)"),
    DynCase(f"{DR}({T1},0)"),
    DynCase(f"{DR}({T1},4)"),
    DynCase(f"{DR}({T1},-4)"),
    DynCase(f"{DR}({T1},10)"),
    DynCase(f"{DR}({T1},0,3)"),
    DynCase(f"{DR}({T1},1.9)"),
    DynCase(f'{DR}({T1},"x")'),
    DynCase(f"{DR}(5,0)"),
    DynCase(f"{DR}(5,1)"),
    DynCase(f"{DR}({T1},{{0,1}})"),
    DynCase(f"{DR}({T1},)"),
    DynCase(f"{DR}({T1},#DIV/0!)"),
    DynCase(f"{EX}({{1,2;3,4}},3,3)", readers=READ),
    DynCase(f"{EX}({{1,2;3,4}},3,3,0)"),
    DynCase(f'{EX}({{1,2;3,4}},3,,"x")'),
    DynCase(f"{EX}({{1,2;3,4}},,3)"),
    DynCase(f"{EX}({{1,2;3,4}},1,2)"),
    DynCase(f"{EX}({{1,2;3,4}},2,2)"),
    DynCase(f'{EX}({T1},5,4,"")'),
    DynCase(f"{EX}({T1},5)"),
    DynCase(f"{EX}(1,2,2)"),
    DynCase(f"{EX}({{1,2}},2,2,{{9,8}})"),
    DynCase(f"{EX}({{1,2}},2,2,#DIV/0!)"),
    DynCase(f"{EX}({{1,2}},-1)"),
    DynCase(f"{EX}({{1,2}},1.5,3.9)"),
    DynCase(f"{EX}({{1,2}},2,2,BU2)"),
    DynCase(f"{EX}({{1,2}},1048577)"),
    DynCase(f"{EX}({{1,2}},2,,)"),
    DynCase(f"{EX}({{1,2}},2,0)"),
    DynCase(f'{EX}({{1,2}},"2")'),
    DynCase(f"{EX}({{1,2}},TRUE)"),
    DynCase(f"{EX}({{1,2}},#N/A)"),
    DynCase("EXPAND({1,2},2)"),
    # second set
    DynCase(f"{TK}({T2},2,-3)", readers=READ),
    DynCase(f"{TK}({T2},-1)"),
    DynCase(f"{DR}({T2},-1,1)"),
    DynCase(f"{DR}({T2},1,-2)"),
    DynCase(f'{EX}({T2},4,5,"-")'),
    DynCase(f"{EX}(BX1:BX3,3,2,TRUE)"),
    DynCase(f"{TK}(BY1:BY3,-2)"),
    DynCase(f"{DR}({T2},{{1;2}})"),
    # follow-up: plain formulas over whole columns, lifting, direct errors
    DynCase(f"{TK}({{1,2;3,4}},1)", legacy=True),
    DynCase(f"{TK}(BT:BT,2)", legacy=True),
    DynCase(f"{DR}({T1},0)", legacy=True),
    DynCase(f"{EX}({T1},4)", legacy=True),
    DynCase(f"{TK}(5,{{1,1}})"),
    DynCase(f"{TK}({{1,2;3,4}},{{1,2}})"),
    DynCase(f"{DR}(5,{{0,0}})"),
    DynCase(f"{EX}(1,{{1,2}})"),
    DynCase(f"{TK}(#N/A,1)"),
    DynCase(f"{EX}(#N/A,2)"),
    DynCase(f"{EX}({{1,2}},2,3,)"),
    DynCase(f"{DR}(BT:BT,1)", legacy=True),
    # follow-up: an empty array argument
    DynCase(f'{EX}(,2,2,"x")'),
    DynCase(f"{EX}(,1)"),
    DynCase(f"{TK}(,,)"),
    DynCase(f"{DR}(,0)"),
]

STACK = [
    DynCase(f"{VS}({{1,2}},{{3,4}})", readers=READ),
    DynCase(f"{VS}({{1,2,3}},{{4}})"),
    DynCase(f'{VS}(1,"a",TRUE)'),
    DynCase(f"{VS}({T1},BX1:BZ3)"),
    DynCase(f"{VS}(BT1:BT2,BU1:BU2)"),
    DynCase(f"{VS}(#N/A,1)"),
    DynCase(f"{VS}({{1,2}},{{3,4}})", way={(1, 1): 0.0}),
    DynCase(f"{VS}(BU2:BU2,)"),
    DynCase(f"{VS}({T1})"),
    DynCase(f"{VS}({{1;2}},{{3,4,5}})"),
    DynCase(f'{VS}("",BU4)'),
    DynCase(f"{VS}({{1,2}},{{3,4}})", legacy=True),
    DynCase("VSTACK(1,2)"),
    DynCase(f"SUM({VS}(1,2,3))"),
    DynCase(f'{VS}({SEQ}(2),{{"a","b"}})'),
    DynCase(f"{HS}({{1;2}},{{3;4}})", readers=READ),
    DynCase(f"{HS}({{1;2;3}},5)"),
    DynCase(f'{HS}(1,"",FALSE)'),
    DynCase(f"{HS}(BT1:BT3,BV1:BV3)"),
    DynCase(f"{HS}(BU2)"),
    DynCase(f"{HS}({{1,2;3,4}},{{5;6;7}})"),
    DynCase(f"{HS}({T1},{T1})"),
    DynCase(f'{HS}(#DIV/0!,"x")'),
    DynCase(f"{HS}({{1;2}},{{3;4}})", way={(1, 0): "in the way"}),
    DynCase("HSTACK(1,2)"),
    # second set
    DynCase(f"{HS}(BX1:BX3,BY1:BY3,CA1:CA3)", readers=READ),
    DynCase(f"{VS}(BX1:CA1,BX3:CA3)"),
    DynCase(f"{HS}(BZ1:BZ3,{{1;2}})"),
    DynCase(f"{VS}(BY1:BY3,BX2)"),
    # follow-up: plain formulas over whole columns, lifting, direct errors
    DynCase(f"{VS}({T1})", legacy=True),
    DynCase(f"{HS}({T1})", legacy=True),
    DynCase(f"{VS}(1,#N/A)"),
    DynCase(f"{HS}({{1,2}},#N/A)"),
    DynCase(f"{VS}(BY2,1)"),
    DynCase(f"{VS}({{1,#N/A}})"),
    DynCase(f"{HS}(1,)"),
    # follow-up: an empty array argument
    DynCase(f"{VS}(,)"),
    DynCase(f"{HS}(,{{1;2}})"),
]

# Formulas that make Excel refuse the file (error -50 on opening, each alone in a workbook):
# SORT and FILTER without their prefix (UNIQUE, SORTBY and SEQUENCE without theirs open), and
# argument counts the functions do not take. The reshaping functions without their prefix open.
REFUSED = (
    "SORT(AA1:AA8)",
    "FILTER(AE1:AE6,AE1:AE6>2)",
    f"{SBY}(AE1:AE6)",
    f"{SBY}(AE1:AE6,AD1:AD6,1,AE1:AE6)",
    f"{SBY}(AE1:AE6,AD1:AD6,1,AE1:AE6,1,AE1:AE6)",
    f"{FLT}(AE1:AE6)",
    f"{SRT}()",
    f"{UNQ}()",
    f"{SEQ}()",
    # the reshaping functions with too few or too many arguments (capture/refusal_probe.py)
    f"{CC}({T1})",
    f"{CR}({T1})",
    f"{DR}({T1})",
    f"{TK}({T1})",
    f"{EX}({T1})",
    f"{HS}()",
    f"{VS}()",
    f"{DR}({T1},1,1,1)",
    f"{TK}({T1},1,1,1)",
    f"{EX}({T1},5,5,0,0)",
)

SUITES: dict[str, list[DynCase]] = {
    "sequence": SEQUENCE,
    "mmult": MMULT,
    "unique": UNIQUE,
    "sort": SORT,
    "sortby": SORTBY,
    "filter": FILTER,
    "choose": CHOOSE,
    "droptake": DROPTAKE,
    "stack": STACK,
}


def _letter(n: int) -> str:
    s = ""
    while n:
        n, r = divmod(n - 1, 26)
        s = chr(65 + r) + s
    return s


def anchor_of(i: int) -> tuple[int, int]:
    """The anchor of case ``i``: (row, column), 1-based."""
    return 1 + i * BLOCK, 2


def layout(suite: str) -> list[tuple[str, str, str]]:
    """(kind, cell, text) for a suite: data and way constants as reprs, formulas in file form."""
    out: list[tuple[str, str, str]] = []
    for letters, values in DATA.items():
        for r, v in enumerate(values, 1):
            if v is None:
                continue
            where = f"{letters}{r}"
            if isinstance(v, str) and v.startswith("="):
                out.append(("legacy", where, v[1:]))
            else:
                out.append(("data", where, repr(v)))
    for i, case in enumerate(SUITES[suite]):
        row, col = anchor_of(i)
        anchor = f"{_letter(col)}{row}"
        for (dr, dc), v in case.way.items():
            out.append(("way", f"{_letter(col + dc)}{row + dr}", repr(v)))
        out.append(("legacy" if case.legacy else "dynamic", anchor, case.formula))
        for j, text in enumerate(case.readers):
            out.append(("reader", f"N{row + j}", text.replace("{a}", anchor)))
    return out


__all__ = ["BLOCK", "DATA", "REFUSED", "SUITES", "DynCase", "anchor_of", "layout"]
