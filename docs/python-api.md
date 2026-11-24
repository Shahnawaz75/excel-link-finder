---
title: Python API
description: Use excel-link-finder from Python - scan() lists external links in .xlsx/.xlsm workbooks, clean() writes a copy without them.
---

# Python API

```python
from excel_link_finder import clean, find_workbooks, scan
```

## `scan(path, *, check_paths=True) -> ScanResult`

Reads one workbook (.xlsx, .xlsm, .xltx, .xltm) and returns every place that refers to another file.
`check_paths=False` skips the broken-link check. Raises `UnsupportedFileError` for .xls, .xlsb, encrypted and
non-workbook files, and `ExcelLinkFinderError` for damaged files.

```python
result = scan("book.xlsx")
result.has_links   # True if there is at least one data link
result.counts()    # {"external-link": 2, "cell-formula": 15, ...}
result.links       # data links: every finding except connections and hyperlinks
result.broken      # external links (and pivot sources) whose file does not exist
result.findings    # everything, in workbook order
result.to_dict()   # what --json prints for one file
```

## `Finding`

| Field | Meaning |
| --- | --- |
| `type` | one of `external-link`, `defined-name`, `cell-formula`, `table-formula`, `conditional-format`, `data-validation`, `sparkline`, `chart`, `pivot-cache`, `shape-macro`, `object-link`, `connection`, `hyperlink` |
| `sheet` | sheet name (for defined names: the sheet a sheet-level name belongs to) |
| `object` | defined name, chart / shape / control name, table column, connection name, or `[N]` for a link |
| `ref` | cell or range, for example `B2`, `B2:B6` |
| `text` | the formula or text as stored in the file |
| `links` | external link numbers used, for example `[2]` |
| `target` | the linked file(s) or URL |
| `status` | for links and pivot sources: `found`, `missing`, `url`, `unknown` or `not-checked` |
| `note` | extra detail: `hidden`, `via name VAT`, `series 1 values`, `shared formula, 3 cells` ... |
| `is_data_link` | False for connections and hyperlinks |

## `clean(source, output=None, *, keep_hyperlinks=False, only=None, overwrite=False) -> CleanResult`

Writes a copy of `source` without its external links and returns what was done. The default output is
`NAME_clean.xlsx` next to the source. The source is never modified; `clean` refuses to overwrite it, refuses
to overwrite an existing output unless `overwrite=True`, and the output must keep the source's extension.
`only` is a list of types from `CLEANABLE_TYPES` to limit the cleaning.

```python
cleaned = clean("book.xlsx", "book_clean.xlsx")
cleaned.removed          # {"cell-formula": 15, "defined-name": 3, ..., "external-link": 2}
cleaned.removed_links    # targets of the deleted external link parts
cleaned.kept             # findings left in place (pivot sources, connections, types not selected)
cleaned.remaining        # ScanResult of the cleaned copy
cleaned.is_clean         # True when no data links are left
cleaned.empty_cells      # formulas that had no saved value and are now empty
cleaned.calc_chain_removed
```

## Other helpers

- `find_workbooks(paths, recursive=False) -> (workbooks, skipped)` expands files and folders the way the
  command line does; `skipped` lists .xls / .xlsb files found in folders.
- `check_package(path) -> list[str]` checks that every relationship target exists, every part has a content
  type and every `r:id` resolves. An empty list means the package is consistent.
- Constants: `DATA_LINK_TYPES`, `INFO_TYPES`, `ALL_TYPES`, `CLEANABLE_TYPES`, `__version__`.
