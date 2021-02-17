from pathlib import Path
from typing import Callable, Dict, Generator, List

from click.testing import CliRunner
from mailmerge import MailMerge
import pytest


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


@pytest.fixture
def options_gen_cli(tmp_path: Path) -> Callable:
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
