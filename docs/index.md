---
title: excel-link-finder
description: Find and remove hidden external links in Excel files (.xlsx, .xlsm) without Excel. Lists links in formulas, defined names, conditional formatting, data validation, charts, pivot tables and buttons, and writes a cleaned copy when Break Links is not working.
---

# excel-link-finder

**Find external links in Excel files and remove them safely, without Excel.**

Excel warns *"This workbook contains links to one or more external sources that could be unsafe"* or
*"We can't update some of the links in your workbook"*, and Data > Edit Links > **Break Link does not make the
link go away**. The link is usually hiding in a place Excel's dialog does not fix: a hidden defined name, a
conditional formatting rule, a data validation list, a chart series, a pivot table source, or a button that
runs a macro from another workbook.

excel-link-finder is a small, free, open-source Python tool that reads the .xlsx / .xlsm file directly:

- `excel-link-finder scan book.xlsx` lists **every** external link with its type, sheet, name or object,
  cell, formula and linked file, and flags **broken links** whose file no longer exists;
- `excel-link-finder clean book.xlsx -o book_clean.xlsx` writes a **cleaned copy** that opens in Excel
  without a repair prompt (the original is never changed);
- it scans whole folders, prints JSON or CSV, returns CI-friendly exit codes and needs no Excel and no
  dependencies, so it runs on Windows, macOS and Linux.

```bash
pip install git+https://github.com/Shahnawaz75/excel-link-finder
excel-link-finder scan book.xlsx
excel-link-finder clean book.xlsx -o book_clean.xlsx
```

## Pages

- [Where Excel hides external links](where-excel-hides-links.md): every location, how links get there, how to
  fix each one in Excel by hand, and what `clean` does.
- [Example output](example-output.md): real output on an Excel-made example workbook, including what Excel's own
  Break Link leaves behind.
- [Python API](python-api.md): `scan()`, `clean()` and the result objects.
- [README on GitHub](https://github.com/Shahnawaz75/excel-link-finder#readme): install, options, CI use,
  limitations and FAQ.

## Frequently asked

**Does it work without Excel, on Linux or macOS?** Yes. It only uses the Python standard library.

**Does it work with .xls files?** No. Open the file in Excel and save it as .xlsx first; the tool says so
when you give it an .xls file.

**Is my file changed or uploaded?** No. `scan` only reads; `clean` writes a new copy. Nothing is sent over the
network, and link targets that are URLs are never fetched.

The source code, issues and releases are on [GitHub](https://github.com/Shahnawaz75/excel-link-finder).
MIT licensed.
