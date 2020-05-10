from pathlib import Path
from typing import Generator, List

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
def options(tmp_path: Path) -> List[str]:
    return [
        "--template",
        "tests/data/example.docx",
        "--data",
        "tests/data/example.csv",
        "--name",
        "NAME",
        "--path",
        str(tmp_path.resolve()),
    ]


@pytest.mark.usefixtures("cleanoutputdir")
def test_basic_succeeds(runner: CliRunner, options: List) -> None:

    del options[-2:]

    result = runner.invoke(cli.main, options)

    assert result.exit_code == 0


def test_custom_output_path_succeeds(runner: CliRunner, options: List) -> None:
    """testing pathlib.Path object as output"""

    result = runner.invoke(cli.main, options)

    assert result.exit_code == 0


def test_custom_deep_output_path_succeeds(runner: CliRunner, options: List) -> None:
    """testing deep non-existent directory"""

    options[-1] = Path(options[-1]) / "deep" / "sub" / "folder"

    result = runner.invoke(cli.main, options)

    assert result.exit_code == 0


def test_custom_output_dir_succeeds(
    runner: CliRunner, options: List, tmpdir: Path
) -> None:
    """testing py.path.local object as output
    the `tmpdir` provides a temporary directory unique to this test invocation"""

    options[-1] = tmpdir

    result = runner.invoke(cli.main, options)

    assert result.exit_code == 0


def test_output_name_not_found(runner: CliRunner, options: List) -> None:

    options[5] = "WRONG_NAME"

    result = runner.invoke(cli.main, options)

    assert result.exit_code == 1


def test_csv_header_not_found(runner: CliRunner, options: List) -> None:

    options[3] = "tests/data/example_missing_column.csv"

    result = runner.invoke(cli.main, options)

    assert result.exit_code == 0
