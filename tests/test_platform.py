"""``platform=`` and ``regional_settings=``: which Excel computes, and the viewer's settings.

Every entry point takes both (``load``, ``Loader``, ``recalculate``, ``evaluate``,
``Formula.parse``, ``ConditionalFormula.parse`` and ``cell_is``); ``None`` is the Mac and the
platform's en-US preset. ``TEXT`` passes both on to ``xlformat``. Excel for Windows is refused
until its captures are in the repository.
"""

from __future__ import annotations

import io

import pytest
import xlformat
import xlio

import xlrecalc
from tests.xlsx_files import package

DATA = package('<row r="1"><c r="A1"><v>2</v></c><c r="B1"><f>TEXT(A1,"0.00")</f></c></row>')


def test_the_mac_is_the_default() -> None:
    wb = xlrecalc.load(io.BytesIO(DATA))
    assert wb.platform == "mac"
    assert wb.regional_settings is xlformat.RegionalSettings.EN_US_MAC
    assert wb.value("Data!B1") == "2.00"


@pytest.mark.parametrize("platform", [None, "mac", xlformat.Platform.MAC])
def test_the_mac_by_name_or_as_xlformat_platform(platform: object) -> None:
    wb = xlrecalc.load(io.BytesIO(DATA), platform=platform)  # type: ignore[arg-type]
    assert wb.platform == "mac"
    with xlio.open(io.BytesIO(DATA)) as book:
        assert xlrecalc.Loader(book, platform=platform).finish().platform == "mac"  # type: ignore[arg-type]
    assert xlrecalc.evaluate('=TEXT(2,"0.0")', platform=platform) == "2.0"  # type: ignore[arg-type]


def _entry_points(tmp_path: object) -> dict[str, object]:
    missing = f"{tmp_path}/none.xlsx"  # refused before the file is opened
    return {
        "load": lambda **k: xlrecalc.load(missing, **k),
        "Loader": lambda **k: xlrecalc.Loader(xlio.open(io.BytesIO(DATA)), **k),
        "recalculate": lambda **k: xlrecalc.recalculate(missing, **k),
        "evaluate": lambda **k: xlrecalc.evaluate("=1+1", **k),
        "Formula.parse": lambda **k: xlrecalc.Formula.parse("=1+1", **k),
        "ConditionalFormula.parse": lambda **k: xlrecalc.ConditionalFormula.parse(
            "A1>1", sqref="A1", **k
        ),
        "ConditionalFormula.cell_is": lambda **k: xlrecalc.ConditionalFormula.cell_is(
            "greaterThan", ["1"], sqref="A1", **k
        ),
    }


@pytest.mark.parametrize(
    "entry",
    [
        "load",
        "Loader",
        "recalculate",
        "evaluate",
        "Formula.parse",
        "ConditionalFormula.parse",
        "ConditionalFormula.cell_is",
    ],
)
def test_each_entry_point_refuses_what_it_cannot_honor(entry: str, tmp_path: object) -> None:
    call = _entry_points(tmp_path)[entry]
    with pytest.raises(LookupError, match="Windows"):
        call(platform="windows")  # type: ignore[operator]
    with pytest.raises(LookupError, match="Windows"):
        call(platform=xlformat.Platform.WINDOWS)  # type: ignore[operator]
    with pytest.raises(ValueError, match="platform must be"):
        call(platform="linux")  # type: ignore[operator]
    with pytest.raises(TypeError, match="regional_settings"):
        call(regional_settings="en-US")  # type: ignore[operator]


def test_text_passes_the_platform_and_the_regional_settings_to_xlformat(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    seen: list[dict[str, object]] = []
    text_function = xlformat.text_function

    def recorded(*args: object, **kwargs: object) -> object:
        seen.append(kwargs)
        return text_function(*args, **kwargs)  # type: ignore[arg-type]

    monkeypatch.setattr(xlformat, "text_function", recorded)
    regional = xlformat.RegionalSettings.EN_US_WINDOWS
    wb = xlrecalc.load(io.BytesIO(DATA), regional_settings=regional)
    assert wb.regional_settings is regional
    assert wb.value("Data!B1") == "2.00"
    assert xlrecalc.evaluate('=TEXT(3,"0")', regional_settings=regional) == "3"
    assert [(k["platform"], k["regional_settings"]) for k in seen] == [("mac", regional)] * 2
