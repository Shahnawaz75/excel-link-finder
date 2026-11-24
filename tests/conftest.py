from __future__ import annotations

import shutil
import sys
from pathlib import Path

import pytest

sys.path.insert(0, str(Path(__file__).resolve().parent))

import builder  # noqa: E402

EXAMPLES = Path(__file__).resolve().parent.parent / "examples"
OLE2 = b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1"


@pytest.fixture
def kitchen(tmp_path: Path) -> Path:
    """Synthetic workbook; its link [2] (Source.xlsx) exists, link [3] does not."""
    builder.plain(tmp_path / "Source.xlsx")
    return builder.kitchen_sink(tmp_path / "kitchen.xlsx")


@pytest.fixture
def plain(tmp_path: Path) -> Path:
    return builder.plain(tmp_path / "plain.xlsx")


@pytest.fixture
def example(tmp_path: Path) -> Path:
    """Copy of the Excel-made example workbook (and its source) in a temporary folder."""
    for name in ("sales-report.xlsx", "prices.xlsm"):
        shutil.copyfile(EXAMPLES / name, tmp_path / name)
    return tmp_path / "sales-report.xlsx"
