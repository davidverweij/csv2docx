from pathlib import Path

import click.testing
import pytest

from csv2docx import cli


default_outpath = Path.cwd() / "output"


@pytest.fixture
def runner():
    return click.testing.CliRunner()


# not using a tmp_path to test default output folder
def test_basic_succeeds(runner):
    result = runner.invoke(
        cli.main,
        [
            "--template",
            "tests/data/example.docx",
            "--data",
            "tests/data/example.csv",
            "--name",
            "NAME",
        ],
    )
    assert result.exit_code == 0

    # remove standard output files
    for f in default_outpath.glob("*.docx"):
        try:
            f.unlink()
        except OSError as e:
            print("Error: %s : %s" % (f, e.strerror))


def test_custom_output_succeeds(runner, tmp_path, tmpdir):

    """testing pathlib.Path object as output"""
    result_path = runner.invoke(
        cli.main,
        [
            "--template",
            "tests/data/example.docx",
            "--data",
            "tests/data/example.csv",
            "--name",
            "NAME",
            "--path",
            tmp_path,
        ],
    )
    assert result_path.exit_code == 0

    # testing deep non-existent directory
    result_unknown_path = runner.invoke(
        cli.main,
        [
            "--template",
            "tests/data/example.docx",
            "--data",
            "tests/data/example.csv",
            "--name",
            "NAME",
            "--path",
            tmp_path / "deep" / "sub" / "folder",
        ],
    )
    assert result_unknown_path.exit_code == 0

    # testing py.path.local object as output
    # the `tmpdir` provides a temporary directory unique to this test invocation
    result_dir = runner.invoke(
        cli.main,
        [
            "--template",
            "tests/data/example.docx",
            "--data",
            "tests/data/example.csv",
            "--name",
            "NAME",
            "--path",
            tmpdir,
        ],
    )
    assert result_dir.exit_code == 0


def test_output_name_not_found(runner, tmp_path):
    result = runner.invoke(
        cli.main,
        [
            "--template",
            "tests/data/example.docx",
            "--data",
            "tests/data/example.csv",
            "--name",
            "WRONG_NAME",
            "--path",
            tmp_path,
        ],
    )
    assert result.exit_code == 1


def test_csv_header_not_found(runner, tmp_path):
    result = runner.invoke(
        cli.main,
        [
            "--template",
            "tests/data/example.docx",
            "--data",
            "tests/data/example_missing_column.csv",
            "--name",
            "NAME",
            "--path",
            tmp_path,
        ],
    )
    assert result.exit_code == 0


"""
# Below are test setups for when a file is not found. But we need to exit this
# gracefully first for this to be proper tests


def test_csv_not_found(runner):
    result = runner.invoke(
        cli.main,
        [
            "--template",
            "tests/data/example.docx",
            "--data",
            "tests/data/non_existing_example.csv",
            "--name",
            "NAME",
        ],
    )

    print(
        result.output
    )  # only visible when running with `-s`, i.e. `poetry run pytest -s`
    assert result.exit_code == 1


def test_docx_not_found(runner):
    result = runner.invoke(
        cli.main,
        [
            "--template",
            "tests/data/non_existing_example.docx",
            "--data",
            "tests/data/example.csv",
            "--name",
            "NAME",
        ],
    )

    print(
        result.output
    )  # only visible when running with `-s`, i.e. `poetry run pytest -s`
    assert result.exit_code == 1
"""
