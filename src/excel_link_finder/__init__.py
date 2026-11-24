"""excel-link-finder: find and remove hidden external links in Excel files, without Excel.

>>> from excel_link_finder import scan, clean
>>> result = scan("book.xlsx")
>>> for finding in result.links:
...     print(finding.type, finding.sheet, finding.ref, finding.text, finding.target)
>>> clean("book.xlsx", "book_clean.xlsx")
"""

__version__ = "0.1.0"

from .cleaner import CLEANABLE_TYPES, CleanResult, clean  # noqa: E402
from .ooxml import check_package  # noqa: E402
from .scanner import (  # noqa: E402
    ALL_TYPES,
    DATA_LINK_TYPES,
    INFO_TYPES,
    ExcelLinkFinderError,
    Finding,
    ScanResult,
    UnsupportedFileError,
    find_workbooks,
    scan,
)

__all__ = [
    "ALL_TYPES",
    "CLEANABLE_TYPES",
    "DATA_LINK_TYPES",
    "INFO_TYPES",
    "CleanResult",
    "ExcelLinkFinderError",
    "Finding",
    "ScanResult",
    "UnsupportedFileError",
    "__version__",
    "check_package",
    "clean",
    "find_workbooks",
    "scan",
]
