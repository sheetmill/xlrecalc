"""Formulas covering every part of the grammar (ARCHITECTURE.md §5.2), without a leading ``=``.

Written for this project. Each must tokenize, parse and print back exactly.
"""

LITERALS = [
    "1",
    "0",
    "1.5",
    ".5",
    "1.",
    "1E3",
    "1e+3",
    "1.5E-10",
    "123456789012345678901234567890",
    '"text"',
    '""',
    '"say ""hi"""',
    '"Données, 東京, ünïcödé"',
    '"line\nbreak"',
    "TRUE",
    "FALSE",
    "true",
    "#NULL!",
    "#DIV/0!",
    "#VALUE!",
    "#REF!",
    "#NAME?",
    "#NUM!",
    "#N/A",
    "#GETTING_DATA",
    "#SPILL!",
    "#CALC!",
]

OPERATORS = [
    "1+2",
    "1-2",
    "2*3",
    "6/3",
    "2^3",
    '"a"&"b"',
    "1=1",
    "1<>2",
    "1<2",
    "1>2",
    "1<=2",
    "1>=2",
    "-1",
    "+1",
    "--1",
    "-+-1",
    "50%",
    "50%%",
    "-2^2",
    "2^-2",
    "2^3^2",
    "1+2*3",
    "(1+2)*3",
    "1+2&3",
    "1&2=12",
    "-A1%",
    "2^3%",
    "A1*-1",
    "A1-+-B1",
]

REFERENCES = [
    "A1",
    "$A$1",
    "A$1",
    "$A1",
    "a1",
    "XFD1048576",
    "A1:B2",
    "$A$1:$B$2",
    "A1:A1",
    "A:A",
    "A:C",
    "$A:$C",
    "1:1",
    "1:3",
    "$1:$3",
    "Sheet2!A1",
    "Sheet2!$A$1:$B$9",
    "Sheet2!A:A",
    "Sheet2!1:2",
    "'My Sheet'!A1",
    "'It''s'!A1",
    "'2024'!A1",
    "Données!A1",
    "'Sheet 1'!A1:B2",
    "Sheet1:Sheet3!A1",
    "'Sheet 1:Sheet 3'!A1:B2",
    "[1]Sheet1!A1",
    "'[1]Sheet 1'!$A$1",
    "[1]!Rate",
    "[1]!'NGH1,PRIM ACT 1'",
    "IF([3]!'NGJ1 LAST'=0,1,[3]!'NGJ1 LAST')",
    "[2]!'It''s'",
    "Sheet1!#REF!",
    "#REF!",
    "Rate",
    "_Rate",
    "my.name",
    "is_ok?",
    "\\backslash",
    "Sheet2!Rate",
    "A1:Rate",
    "XFE1",
    "ABCD1",
    "A1048577",
    "A1#",
    "Sheet2!A1#",
    "@A1:A10",
    "A1:B2 B1:C3",
    "A:A 1:1",
    "(A1,B1)",
    "(A1:A3,C1:C3,E1)",
    "SUM((A1,B1),C1)",
    "INDEX(A1:C3,1,1):C3",
    "A1:INDEX(A1:C3,2,2)",
    "OFFSET(A1,1,1):OFFSET(A1,2,2)",
]

STRUCTURED = [
    "Table1[Price]",
    "Table1[#All]",
    "Table1[#Data]",
    "Table1[#Headers]",
    "Table1[#Totals]",
    "Table1[]",
    "Table1[[#Headers],[Price]]",
    "Table1[[#This Row],[Price]]",
    "Table1[@Price]",
    "Table1[@[Unit Price]]",
    "Table1[[Price]:[Qty]]",
    "Table1[[#Data],[Price]:[Qty]]",
    "Table1[['#Hash]]",
    "Table1[[Col'[1']]]",
    "[@Price]",
    "[@[Unit Price]]*2",
    "SUM(Table1[Price])",
    "Table1[Price]*Table1[Qty]",
]

FUNCTIONS = [
    "SUM(A1:A10)",
    "SUM(A1,B1,C1)",
    "NOW()",
    "PI()",
    "TRUE()",
    "FALSE()",
    "IF(A1,,1)",
    "IF(A1,1,)",
    "IF(,,)",
    "SUM(,)",
    "CHOOSE(2,,,)",
    "IF(A1>0,IF(B1>0,1,2),3)",
    "LOG10(100)",
    "ATAN2(1,1)",
    "DAYS360(A1,B1)",
    "VLOOKUP(A1,Sheet2!$A$1:$C$100,3,FALSE)",
    "INDEX(A:A,MATCH(B1,C:C,0))",
    "SUMPRODUCT((A1:A10>0)*(B1:B10))",
    'COUNTIF(A1:A10,">"&B1)',
    '_xlfn.XLOOKUP(A1,B:B,C:C,"none")',
    "_xlfn.IFS(A1>1,1,TRUE,0)",
    "_xlfn._xlws.SORT(A1:A10)",
    "_xlfn._xlws.FILTER(A1:A10,B1:B10>0)",
    "_xlfn.UNIQUE(A1:A10)",
    "_xlfn.SEQUENCE(3)",
    "_xlfn.ANCHORARRAY(A1)",
    "_xlfn.SINGLE(A1:A10)",
    "_xludf.MYFUNCTION(A1)",
    "_xlfn.LET(_xlpm.x,1,_xlpm.x+1)",
    "_xlfn.LAMBDA(_xlpm.x,_xlpm.x+1)(2)",
    "_xlfn.LAMBDA(_xlpm.x,_xlpm.y,_xlpm.x*_xlpm.y)(2,3)",
    "Sheet2!MyFunc(1)",
    "SUM(Sheet2:Sheet3!A1)",
    "ROUND(SUM(A1:A3)/COUNT(A1:A3),2)",
    'IFERROR(1/0,"x")',
    'INDIRECT("A"&ROW())',
    "OFFSET(A1,1,1,2,2)",
    'TEXT(A1,"#,##0.00")',
    'TEXT(A1,"[Red]0;[Blue]-0")',
    'CONCATENATE(A1," ",B1)',
    'HYPERLINK("https://example.com","x")',
]

ARRAYS = [
    "{1,2,3}",
    "{1;2;3}",
    "{1,2;3,4}",
    "{-1,+2,-3.5E-2}",
    '{"a","b";"c","d"}',
    "{TRUE,FALSE}",
    "{#N/A,1}",
    "SUM({1,2,3}*A1:A3)",
    "{1 , 2 ; 3 , 4}",
]

WHITESPACE = [
    " 1",
    "1 ",
    " 1 + 2 ",
    "SUM( A1:A10 )",
    "SUM(A1 , B1)",
    "ROUND( 2.5 ,0)",
    "IF(A1>0,\n  1,\n  2)",
    "IF(A1>0,\r\n  1,\r\n  2)",
    "\n\nA1",
    "( 1 + 2 )",
    "-\t1",
    "SUM( )",
]

HOSTILE = [
    "\"'); import os; ('\"",
    "\"__import__('os').system('echo')\"",
    '"\\N{BULLET} \\x00 \\u0041"',
    '"""""\'\'\'"""',
    "'__import__(''os'')'!A1",
    "'); import os; ('!A1",
    "'K[0]'!A1+'F[0]'!B1",
    '"lambda: 1"&"exec"',
    "__import__",
    "exec",
    "_xlfn.EVAL(1)",
    "eval(1)",
    "Table1[[#This Row],[__class__]]",
    '{"\\n",";"}',
]

# Forms that real workbooks use (found by tools/formula_census.py), written for this project.
SHEET_NAMES = [
    "'Q1-Q2'!A1",
    "'R&D'!$B$2",
    "'Cash+Debt'!C3",
    "'Plan (v2)'!D4",
    "'v1.2'!E5",
    "'a,b'!F6",
    "'3-15-24'!A1",
    "'2024 Budget'!$A$1:$B$9",
    "'Résumé'!A1",
    "Résumé!A1",
    "'Ops ''25'!A1",
    "'East:West'!A1",
    "'Jan-24:Mar-24'!B2",
    "SUM('North (1):South (9)'!C3)",
    "'[3]Q1-Q2'!A1",
    "'[3]R&D Plan '!$B$2",
    "'[4]2024 Plan'!C3",
    "'[4]12-31'!D4*-1",
    "'[5]Año 1'!E5",
    "'[5]Lucy''s Sheet '!F6",
    "'[2]Ops ''25 (B)'!G7",
    "+'[1]1-2 Summary'!$A$1/100",
    "[2]!'.SPX,LAST,1'",
    "[2]!'.SPX,LAST,1'/100",
    "IF([2]!'TICK A'=0,1,[2]!'TICK A')",
    "+'[2]2024'!$A$1",
    "SUM('10 (a):20 (b)'!C1)",
    "IF(A1>='12 - Plan Notes '!B$2,1,0)",
    "-'[2]3-Año'!$A$1/24",
    "'[2]1-Año Fiscal '!$B$8",
]
NAMES_AND_PREFIXES = [
    "_xll.ADDIN(A1,B1)",
    "C9*_xll.ADDIN.RATE($B9,M$7)",
    "_xll.ADDIN()",
    "sum(a1:a3)",
    "Sum(A1:A3)",
]
NUMBERS = [
    "0.123456789012345678",
    "A1*0.12345678901234567",
    "1234567890123456789.5",
]
SPACING = [
    "IF(A1=0, 0,  1)",
    "IF(A1=0,  B1,   C1)",
    "SUM(A1)  +  SUM(B1)",
    '"a"  & "b"',
    "A1/(B1*C1  )",
    "(1+2   )",
    "      6:6",
    "IF(A1>0,1,2)\n* B1",
    "A1\n+B1\n+C1",
    "SUM(\n  A1,\n  B1\n)",
]


def _nested_if(depth: int) -> str:
    return "IF(A1>0," * depth + "1" + ",0)" * depth


def _arguments(count: int) -> str:
    return "SUM(" + ",".join(f"A{k}" for k in range(1, count + 1)) + ")"


def _long(length: int) -> str:
    terms, text = [], ""
    k = 1
    while len(text) < length:
        terms.append(f"IF($B{k}>0,C{k}*D{k},0)")
        text = "+".join(terms)
        k += 1
    return text


SIZES = [
    *(_nested_if(d) for d in (4, 7, 8, 15, 16, 31, 32, 64)),
    "((((((((((1))))))))))",
    "-(-(-(-(-(-(-(-1)))))))",
    *(_arguments(n) for n in (11, 30, 31, 100, 255)),
    *(_long(n) for n in (101, 255, 256, 1000, 1001, 4000, 8000)),
]

CORPUS = (
    LITERALS
    + OPERATORS
    + REFERENCES
    + STRUCTURED
    + FUNCTIONS
    + ARRAYS
    + WHITESPACE
    + HOSTILE
    + SHEET_NAMES
    + NAMES_AND_PREFIXES
    + NUMBERS
    + SPACING
    + SIZES
)

# Not formulas: each must raise a syntax or tokenize error.
INVALID = [
    "",
    "1+",
    "(1",
    "1)",
    "SUM(1",
    '"open',
    "'open!A1",
    "A1 !",
    "{1,A1}",
    '{-"a"}',
    "Sheet1! A1",
    "Table1[Price",
    "#BOGUS",
    "1 ~ 2",
    "SUM(1,2))",
]
