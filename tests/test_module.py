from pathlib import Path
from typing import Generator

from mailmerge import MailMerge
import pytest

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
