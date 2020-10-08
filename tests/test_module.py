from pathlib import Path
from typing import Callable, Dict, Generator

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


@pytest.fixture
def options_gen(tmp_path: Path) -> Callable:
    def _options_gen(**kwargs: str) -> Dict:
        default = {
            "template": "tests/data/example.docx",
            "data": "tests/data/example.csv",
            "name": "NAME",
            "path": str(tmp_path.resolve()),
            "delimiter": ";",
        }

        # override values if provided
        for key, value in kwargs.items():
            if key in default:
                default[key] = value

        # convert dict to sequential list, add '--' for CLI options (i.e. the keys)
        return default

    return _options_gen


# @pytest.mark.usefixtures("cleanoutputdir")
# def test_library_convert(options_gen: Callable) -> None:
#     t = options_gen(path="/output")
#
#     result = csv2docx.convert(
#         t["data"], t["template"], t["name"], t["path"], t["delimiter"]
#     )
#
#     assert result


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
