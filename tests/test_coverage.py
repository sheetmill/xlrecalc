"""The committed coverage matrix is current (``tools/coverage.py``, ARCHITECTURE.md §8.4)."""

from __future__ import annotations

import pytest
from tools.coverage import OUTPUT, matrix, platform_label, recorded_platform


@pytest.mark.excel
def test_coverage_matrix_is_current():
    """The page as generated now, where the platform is the one it was generated on.

    The split between exact and within-tolerance rows depends on the math library's last bits,
    which come with the operating system: elsewhere the comparison would fail for a reason
    unrelated to the page.
    """
    page = OUTPUT.read_text(encoding="utf-8")
    recorded, here = recorded_platform(page), platform_label()
    assert recorded is not None, "docs/coverage.md does not record its platform"
    if recorded != here:
        pytest.skip(f"docs/coverage.md was generated on {recorded}; this is {here}")
    assert page == matrix(), (
        "docs/coverage.md is out of date: run `uv run python tools/coverage.py --write`"
    )


def test_every_registered_function_appears():
    from xlrecalc._functions import SPECS

    text = OUTPUT.read_text(encoding="utf-8")
    for name in SPECS:
        bare = name.removeprefix("_XLFN.")
        assert f"| `{bare}` |" in text, name


def test_another_platform_skips_the_comparison(monkeypatch):
    import tests.test_coverage as module

    monkeypatch.setattr(module, "platform_label", lambda: "macOS 15, arm64")
    with pytest.raises(pytest.skip.Exception, match="this is macOS 15, arm64"):
        test_coverage_matrix_is_current()
