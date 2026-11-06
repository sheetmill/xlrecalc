import pytest

from tests.watchdog import watch_session


def pytest_sessionstart(session: pytest.Session) -> None:
    """A guard that regresses must cost a failed run, not the machine's memory."""
    watch_session()


def pytest_addoption(parser: pytest.Parser) -> None:
    parser.addoption(
        "--tier",
        choices=["fast", "full", "excel"],
        default="full",
        help="Capture tier: 'full' (default, every captured row), 'fast' (a sample), or 'excel'"
        " (full, plus the tests that drive Excel 16 for Mac).",
    )


def pytest_collection_modifyitems(config: pytest.Config, items: list[pytest.Item]) -> None:
    if config.getoption("--tier") == "excel":
        return
    skip = pytest.mark.skip(reason="drives Excel: run with --tier excel")
    for item in items:
        if "excel_app" in item.keywords:
            item.add_marker(skip)


@pytest.fixture(scope="session")
def tier(request: pytest.FixtureRequest) -> str:
    return str(request.config.getoption("--tier"))


def pytest_terminal_summary(terminalreporter: pytest.TerminalReporter) -> None:
    """Report the rows compared with Excel and the known deviations, on every run."""
    from tests.capture_report import REPORT

    if REPORT:
        terminalreporter.section("captures")
        for line in REPORT.values():
            terminalreporter.write_line(line)
