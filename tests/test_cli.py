from pathlib import Path
from typing import Callable, Generator, List

from click.testing import CliRunner
import pytest

from csv2docx import cli


@pytest.fixture
def runner() -> CliRunner:
    return CliRunner()


@pytest.fixture
def cleanoutputdir() -> Generator[None, None, None]:
    default_outpath = Path.cwd() / "output"
    yield
    for f in default_outpath.glob("*"):
        try:
            f.unlink()
        except OSError as e:
            print(f"Error: {f} : {e.strerror}")
    default_outpath.rmdir()


@pytest.fixture
def options_gen(tmp_path: Path) -> Callable:
    def _options_gen(**kwargs: str) -> List:
        default = {
            "template": "tests/data/example.docx",
            "data": "tests/data/example.csv",
            "name": "NAME",
            "delimiter": ";",
            "path": str(tmp_path.resolve()),
        }

        # override values if provided
        for key, value in kwargs.items():
            if key in default:
                default[key] = value

        # convert dict to sequential list, add '--' for CLI options (i.e. the keys)
        return [x for (key, value) in default.items() for x in ("--" + key, value)]

    return _options_gen


@pytest.mark.usefixtures("cleanoutputdir")
def test_defaults(runner: CliRunner, options_gen: Callable) -> None:

    result = runner.invoke(cli.main, options_gen()[:-2])

    assert result.exit_code == 0


def test_output_path(runner: CliRunner, options_gen: Callable) -> None:

    result = runner.invoke(cli.main, options_gen())

    assert result.exit_code == 0


def test_deep_output_path(
    runner: CliRunner, options_gen: Callable, tmp_path: Path
) -> None:

    result = runner.invoke(
        cli.main, options_gen(path=tmp_path / "deep" / "sub" / "folder")
    )

    assert result.exit_code == 0


def test_invalid_file(runner: CliRunner, options_gen: Callable) -> None:

    result = runner.invoke(cli.main, options_gen(data="125123ssdfv9a98a7e43"))

    assert result.exit_code == 0
    assert "could not be opened or found." in result.output


def test_output_name_not_found(runner: CliRunner, options_gen: Callable) -> None:

    result = runner.invoke(cli.main, options_gen(name="notNAME"))

    assert result.exit_code == 0
    assert "could not found in the .csv header" in result.output


def test_csv_header_not_found(runner: CliRunner, options_gen: Callable) -> None:

    result = runner.invoke(
        cli.main, options_gen(data="tests/data/example_missing_column.csv")
    )

    assert result.exit_code == 0
    assert "missing in the .csv header" in result.output
