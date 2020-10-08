from pathlib import Path
from typing import Generator

from mailmerge import MailMerge
import pytest
from pytest_mock import MockerFixture

from csv2docx import csv2docx


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
def mailmerge_docx() -> MailMerge:
    return MailMerge(Path.cwd() / "tests/data/example.docx")


@pytest.mark.parametrize(
    "data, template, name, path, delimiter",
    [
        pytest.param(
            "tests/data/example.csv",
            "tests/data/example.docx",
            "NAME",
            "output",
            ";",
            id="standard",
        ),
        pytest.param(
            "tests/data/example.csv",
            "tests/data/example.docx",
            "NAME",
            "nested/folder",
            ";",
            id="nested folder",
        ),
        pytest.param(
            "tests/data/example.csv",
            "tests/data/example.docx",
            "NAME",
            "/output",
            ";",
            id="leading slash folder",
        ),
        pytest.param(
            "tests/data/example.csv",
            "tests/data/example.docx",
            "notNAME",
            "output",
            ";",
            id="name not found",
            marks=pytest.mark.xfail(raises=ValueError),
        ),
        pytest.param(
            "tests/data/example.csv",
            "tests/data/example.docx",
            "NAME",
            "output",
            ",",
            id="wrong delimiter",
            marks=pytest.mark.xfail(raises=ValueError),  # and gives a UserWarning
        ),
    ],
)
def test_library_convert_success(
    data: str, template: str, name: str, path: str, delimiter: str, tmp_path: Path
) -> None:

    # prepend a temporary path to not affect the module folder
    result = csv2docx.convert(data, template, name, str(tmp_path / path), delimiter)

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
