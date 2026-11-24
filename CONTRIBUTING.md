# Contributing to excel-link-finder

Thanks for helping! Bug reports with a small workbook that shows the problem are the most useful
contribution: links hide in many corners of the file format, and every real example makes the tool better.

## Reporting a missed or wrong link

1. Run `excel-link-finder scan book.xlsx --json -o scan.json`.
2. Tell us where Excel shows the link (Data > Edit Links / Workbook Links, Name Manager, a rule, a chart ...).
3. Attach the workbook if you can. **Delete confidential data first**: replace values with dummy numbers,
   keep only the sheet that holds the link. The link usually survives this.

## Development setup

```bash
git clone https://github.com/Shahnawaz75/excel-link-finder
cd excel-link-finder
python -m venv .venv
. .venv/bin/activate          # Windows: .venv\Scripts\activate
python -m pip install -e ".[test]" ruff
python -m pytest
ruff check . && ruff format --check .
```

The package itself has no dependencies; `openpyxl` is only used by one test that loads a cleaned copy.

## How the code is organised

| File | What it does |
| --- | --- |
| `src/excel_link_finder/ooxml.py` | zip parts, relationships, a small XML tree that remembers byte offsets, the byte-level `Editor`, `check_package()` |
| `src/excel_link_finder/formula.py` | tokenizer that finds `[N]`, `[Book.xlsx]` and full-path references in formulas, and renumbers them |
| `src/excel_link_finder/scanner.py` | walks workbook, sheets, drawings, charts, tables, pivot caches, connections; produces `Finding`s |
| `src/excel_link_finder/cleaner.py` | turns findings into edits, removes unused link parts, writes the cleaned copy |
| `src/excel_link_finder/report.py`, `cli.py` | table / JSON / CSV output and the command line |
| `tests/builder.py` | writes small synthetic workbooks for the tests |
| `examples/make_examples.py` | rebuilds the example workbooks with Microsoft Excel (Windows) |
| `tools/excel_check.py` | opens files in Excel and reports repairs and the links Excel still sees (Windows) |

One rule matters more than any other: **never re-serialise a part with ElementTree**. ElementTree renames
namespace prefixes and drops declarations that are only used in `mc:Ignorable`, which makes Excel show its
"We found a problem with some content" repair prompt. All changes are spliced into the original bytes.

## Testing with Excel (Windows)

The CI runs on Linux and Windows without Excel. If you have Excel, please also check cleaned copies with it:

```bash
python -m pip install pywin32
excel-link-finder clean book.xlsx -o book_clean.xlsx
python tools/excel_check.py book.xlsx book_clean.xlsx
```

`excel_check.py` starts its own hidden Excel instance (it never touches workbooks you have open), opens each
file read-only without updating links, and fails if Excel had to repair the file.
`python examples/make_examples.py` rebuilds the example workbooks the same way.

## Pull requests

- Keep the core free of third-party dependencies.
- Add a test for every new link location (extend `tests/builder.py` or add a small example).
- Run `python -m pytest` and `ruff check . && ruff format --check .` before you push.
- Update `CHANGELOG.md` and the location table in `README.md` when behaviour changes.
