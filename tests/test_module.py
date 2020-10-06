from pathlib import Path
from typing import Generator

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


def test_custom_output_path_succeeds(tmp_path: Path) -> None:
    """testing pathlib.Path object as output"""

    result = csv2docx.create_output_folder(str(tmp_path / "custom_output"))

    assert result.name == "custom_output"
