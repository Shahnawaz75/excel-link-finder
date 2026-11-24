"""Open workbooks in Microsoft Excel and report repairs and remaining external links.

    python tools/excel_check.py FILE [FILE ...]

Windows only; needs Microsoft Excel and ``pip install pywin32``. The script starts its own
hidden Excel instance (it never touches workbooks you have open), opens each file read-only
without updating links, and checks:

* that Excel opened it without repairing it (Excel writes a repair log ``error*.xml`` to the
  temp folder and adds "[Repaired]" to the window caption when it had to fix content),
* how many external links Excel itself still sees (``Workbook.LinkSources``).

Exit code 0 when every file opened cleanly, 1 otherwise.
"""

from __future__ import annotations

import sys
import tempfile
from pathlib import Path

XL_EXCEL_LINKS = 1
XL_NORMAL_LOAD = 0
FORCE_DISABLE_MACROS = 3


def check(paths: list[str]) -> int:
    import win32com.client  # type: ignore[import-not-found]

    temp = Path(tempfile.gettempdir())
    excel = win32com.client.DispatchEx("Excel.Application")  # a new, private Excel process
    excel.Visible = False
    excel.DisplayAlerts = False
    excel.AutomationSecurity = FORCE_DISABLE_MACROS
    failures = 0
    try:
        for name in paths:
            path = Path(name).resolve()
            before = set(temp.glob("error*.xml"))
            try:
                book = excel.Workbooks.Open(str(path), UpdateLinks=0, ReadOnly=True, CorruptLoad=XL_NORMAL_LOAD)
            except Exception as exc:  # noqa: BLE001 - report any COM error
                print(f"FAIL      {path.name}: Excel could not open it ({exc})")
                failures += 1
                continue
            try:
                links = book.LinkSources(XL_EXCEL_LINKS) or ()
                caption = book.Windows(1).Caption if book.Windows.Count else ""
                repaired = sorted(set(temp.glob("error*.xml")) - before)
                ok = not repaired and "repaired" not in str(caption).lower()
                failures += not ok
                state = "OK" if ok else "REPAIRED"
                names = ", ".join(Path(str(link)).name for link in links) or "none"
                print(f"{state:<9} {path.name}: Excel {excel.Version}, external links Excel sees: {names}")
                for log in repaired:
                    print(f"          repair log: {log.read_text('utf-8', 'replace')[:400]}")
            finally:
                book.Close(False)
    finally:
        excel.Quit()
    return 1 if failures else 0


if __name__ == "__main__":
    if len(sys.argv) < 2:
        sys.exit(__doc__)
    sys.exit(check(sys.argv[1:]))
