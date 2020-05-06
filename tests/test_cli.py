import click.testing
import pytest

from csv2docx import cli

@pytest.fixture
def runner():
    return click.testing.CliRunner()

def test_succeeds(runner):
    result = runner.invoke(cli.main, ["--template", "tests/data/example.docx",
                                      "--data", "tests/data/example.csv",
                                      "--name", "NAME",
                                      "--path", "output_folder",
                                      "--delimiter", ";"])

    print(result.output)    # only visible when running with `-s`, i.e. `poetry run pytest -s`
    assert result.exit_code == 0


def test_csv_not_found(runner):
    result = runner.invoke(cli.main, ["--template", "tests/data/example.docx",
                                      "--data", "tests/data/non_existing_example.csv",
                                      "--name", "NAME"])

    print(result.output)    # only visible when running with `-s`, i.e. `poetry run pytest -s`
    assert result.exit_code == 1

def test_docx_not_found(runner):
    result = runner.invoke(cli.main, ["--template", "tests/data/non_existing_example.docx",
                                      "--data", "tests/data/example.csv",
                                      "--name", "NAME"])

    print(result.output)    # only visible when running with `-s`, i.e. `poetry run pytest -s`
    assert result.exit_code == 1

def test_name_not_found(runner):
    result = runner.invoke(cli.main, ["--template", "tests/data/example.docx",
                                      "--data", "tests/data/example.csv",
                                      "--name", "WRONG_NAME"])

    print(result.output)    # only visible when running with `-s`, i.e. `poetry run pytest -s`
    assert result.exit_code == 1
