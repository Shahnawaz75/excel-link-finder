# Changelog

All notable changes to this project are documented here.
The format follows [Keep a Changelog](https://keepachangelog.com/en/1.1.0/) and the project uses
[Semantic Versioning](https://semver.org/).

## [0.1.0] - 2026-11-24

First release.

### Added

- `excel-link-finder scan`: lists every external link in .xlsx, .xlsm, .xltx and .xltm workbooks without
  Excel: external link parts (with broken-link check for local and network paths), cell formulas
  (including array and shared formulas, and formulas that use an external defined name), defined names
  (also hidden and sheet-level ones), table calculated columns, conditional formatting and data validation
  (classic and Excel 2010+ forms), sparklines, chart series / categories / titles, pivot cache sources,
  macros assigned to shapes and buttons, form control links, shape text links and linked OLE objects.
  Data connections, Power Query queries and external hyperlinks are reported separately.
- Table, `--json` and `--csv` output, folder scanning with `--recursive`, exit codes 0 / 1 / 2.
- `excel-link-finder clean`: writes a cleaned copy (never changes the original). Formulas keep their last
  value, names and rules are deleted, charts keep their cached data, macro assignments and control links
  are cleared, unused external link parts are removed with their relationships and content types, link
  numbers are renumbered when only some links are removed, and `calcChain.xml` is dropped when formulas
  change. Every cleaned copy passes a package consistency check and is scanned again.
- Python API: `scan()`, `clean()`, `find_workbooks()`, `check_package()`.
- Clear messages for unsupported files: .xls, .xlsb and password-protected workbooks.
- Example workbooks made by Microsoft Excel (`examples/make_examples.py`) and a helper that opens files in
  Excel to confirm they load without repair (`tools/excel_check.py`).

[0.1.0]: https://github.com/Shahnawaz75/excel-link-finder/releases/tag/v0.1.0
