from pathlib import Path
from typing import Callable

from mailmerge import MailMerge
import pytest
from pytest_mock import MockerFixture

from csv2docx import csv2docx


def test_library_convert_output_dir(options_gen: Callable, tmpdir: Path) -> None:
    """testing py.path.local object as output"""
    t = options_gen(path=tmpdir)

    result = csv2docx.convert(
        t["data"], t["template"], t["name"], t["path"], t["delimiter"]
    )

    assert result


def test_library_convert_nested_output(options_gen: Callable, tmp_path: Path) -> None:
    t = options_gen(path=tmp_path / "nested" / "output" / "folder")

    result = csv2docx.convert(
        t["data"], t["template"], t["name"], t["path"], t["delimiter"]
    )

    assert result


@pytest.mark.xfail(raises=FileNotFoundError)
def test_library_convert_template_wrong(options_gen: Callable) -> None:
    t = options_gen(template="non-existing.docx")

    result = csv2docx.convert(
        t["data"], t["template"], t["name"], t["path"], t["delimiter"]
    )

    assert result


@pytest.mark.xfail(raises=FileNotFoundError)
def test_library_convert_csv_wrong(options_gen: Callable) -> None:
    t = options_gen(data="non-existing.csv")

    result = csv2docx.convert(
        t["data"], t["template"], t["name"], t["path"], t["delimiter"]
    )

    assert result


@pytest.mark.xfail(raises=ValueError)
def test_library_convert_name_not_found(options_gen: Callable) -> None:
    t = options_gen(name="notNAME")

    result = csv2docx.convert(
        t["data"], t["template"], t["name"], t["path"], t["delimiter"]
    )

    assert result


@pytest.mark.xfail(raises=ValueError)
def test_library_convert_wrong_delimiter(options_gen: Callable) -> None:
    t = options_gen(delimiter=",")

    with pytest.warns(UserWarning):
        result = csv2docx.convert(
            t["data"], t["template"], t["name"], t["path"], t["delimiter"]
        )

    assert result


def test_custom_output_path_succeeds(tmp_path: Path) -> None:
    """testing pathlib.Path object as output"""

    result = csv2docx.create_output_folder(str(tmp_path / "custom_output"))

    assert isinstance(result, Path)
    assert result.name == "custom_output"


def test_unique_naming(tmp_path: Path, mailmerge_docx: MailMerge) -> None:
    """testing filenaming scheme"""

    for _ in range(20):
        mailmerge_docx.write(csv2docx.create_unique_name("text", tmp_path))

    result = [f for f in Path(tmp_path).iterdir() if f.is_file()]

    assert len(result) == len(set(result))


def test_empty_csv_rows(tmp_path: Path, mocker: MockerFixture) -> None:
    """testing behaviour when .csv file has empty rows"""

    mocker.patch.object(MailMerge, "write")

    result = csv2docx.convert(
        "tests/data/example_emptyrow.csv",
        "tests/data/example.docx",
        "NAME",
        str(tmp_path),
        ";",
    )

    assert result
    MailMerge.write.not_called()
