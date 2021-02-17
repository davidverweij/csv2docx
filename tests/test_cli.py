from pathlib import Path
from typing import Callable

from click.testing import CliRunner
import pytest

from csv2docx import cli


@pytest.mark.usefixtures("cleanoutputdir")
def test_defaults(runner: CliRunner, options_gen_cli: Callable) -> None:

    result = runner.invoke(cli.main, options_gen_cli()[:-2])

    assert result.exit_code == 0


def test_output_path(runner: CliRunner, options_gen_cli: Callable) -> None:

    result = runner.invoke(cli.main, options_gen_cli())

    assert result.exit_code == 0


def test_deep_output_path(
    runner: CliRunner, options_gen_cli: Callable, tmp_path: Path
) -> None:

    result = runner.invoke(
        cli.main, options_gen_cli(path=tmp_path / "deep" / "sub" / "folder")
    )

    assert result.exit_code == 0


def test_invalid_file(runner: CliRunner, options_gen_cli: Callable) -> None:

    result = runner.invoke(cli.main, options_gen_cli(data="125123ssdfv9a98a7e43"))

    assert result.exit_code == 1
    assert result.exception


def test_output_name_not_found(runner: CliRunner, options_gen_cli: Callable) -> None:

    result = runner.invoke(cli.main, options_gen_cli(name="notNAME"))

    assert result.exit_code == 1
    assert result.exception


def test_csv_header_not_found(runner: CliRunner, options_gen_cli: Callable) -> None:

    result = runner.invoke(
        cli.main, options_gen_cli(data="tests/data/example_missing_column.csv")
    )

    assert result.exit_code == 1
    assert result.exception
