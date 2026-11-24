# excel-link-finder: find and remove hidden external links in Excel files

[![CI](https://github.com/Shahnawaz75/excel-link-finder/actions/workflows/ci.yml/badge.svg)](https://github.com/Shahnawaz75/excel-link-finder/actions/workflows/ci.yml)
[![License: MIT](https://img.shields.io/badge/license-MIT-blue.svg)](LICENSE)

A free Python command-line tool to **find external links in Excel** workbooks (.xlsx, .xlsm) and remove them
safely when Edit Links > **Break Links is not working** and Excel keeps saying *"This workbook contains links
to external sources"*, with **no Excel needed**, so it also runs on Linux, macOS, servers and CI.

```text
$ excel-link-finder scan examples/sales-report.xlsx
examples\sales-report.xlsx
  TYPE                SHEET / NAME / OBJECT                 CELL/RANGE  FORMULA / TEXT                        LINKED FILE
  external-link       [1] (workbook link, last seen at ...  -           externalLink1.xml                     old-budget.xlsx (MISSING)
  external-link       [2] (workbook link)                   -           externalLink2.xml                     prices.xlsm (found)
  defined-name        OldTotal (hidden)                     -           [1]Budget!$B$10                       [1] old-budget.xlsx
  defined-name        TaxRate                               -           [2]Prices!$F$1                        [2] prices.xlsm
  ...
  cell-formula        Report (via name VAT)                 B2          [2]Prices!B2*(1+VAT)                  [2] prices.xlsm
  ...
  conditional-format  Report (expression rule)              B2:B6       B2>[prices.xlsm]Prices!#REF!          prices.xlsm
  data-validation     Report (list rule)                    D2:D6       'C:\work\examples\[prices.xlsm]Pr...  C:\work\examples\prices.xlsm
  object-link         Report > Drop Down 1 (form contro...  -           [2]Prices!$A$2:$A$6                   [2] prices.xlsm
  shape-macro         Report > Button 2 (form control m...  -           [2]!UpdateReport                      [2] prices.xlsm
  chart               Report > Price chart (chart title)    -           [2]Prices!$B$1                        [2] prices.xlsm
  ...
  pivot-cache         Report > StockPivot                   A1:C6       Prices!A1:C6                          prices.xlsm (found)
  ...

  Found: external-link 2 (1 broken), defined-name 3, cell-formula 15, table-formula 1, conditional-format 1,
    data-validation 1, sparkline 1, chart 4, pivot-cache 1, shape-macro 2, object-link 2
  Also found (not data links): connection 2, hyperlink 1
```

(`...` marks rows left out here. The [full output](docs/example-output.md) lists all 37 findings of this
example workbook, which was made by Microsoft Excel.)

## Why

External links (references to other workbooks) hide in many places Excel's **Data > Edit Links** dialog does
not fix: defined names, conditional formatting, data validation lists, chart series and titles, pivot caches,
buttons and shapes with a macro from another workbook, and form controls. Microsoft's own advice is to hunt
for them by hand, one place at a time.

We tried it on the example workbook: after running Excel's own **Break Link** on every source (through
automation, Excel 365), Excel *still* reported a link to `prices.xlsm`. `excel-link-finder scan` showed why:
a conditional formatting rule, a drop-down list's input range and a pivot table source still pointed at it.
After `excel-link-finder clean`, Excel reported no external links at all.
([Details and output](docs/example-output.md#what-excels-break-link-leaves-behind).)

excel-link-finder reads the .xlsx file directly (it is a zip of XML files), so it

- lists **every** link location in one go: type, sheet, name or object, cell or range, formula and linked file,
- flags **broken links** (linked files that no longer exist on this computer or network),
- writes a **cleaned copy** that opens in Excel without the "We found a problem with some content" prompt,
- scans **whole folders** in batch and works in CI, on Linux and macOS, without Excel or any dependency.

## Install

```bash
pip install git+https://github.com/Shahnawaz75/excel-link-finder
```

Python 3.10 or newer, no other dependencies. A PyPI release is coming soon.

## Quick start

```bash
excel-link-finder scan book.xlsx                      # table of every external link
excel-link-finder scan book.xlsx --json               # the same as JSON (or --csv)
excel-link-finder clean book.xlsx -o book_clean.xlsx  # cleaned copy; book.xlsx is not changed
excel-link-finder scan book_clean.xlsx                # check the result
```

`python -m excel_link_finder` works too.

## Where Excel hides external links

| Location (`scan` type) | Example from a real file | How it usually gets there | What `clean` does |
| --- | --- | --- | --- |
| External link part (`external-link`) | `[2]` = `prices.xlsm`, status *found* / *MISSING* | every link has one; it is what Edit Links lists | deleted once nothing uses it, with its relationships and content type |
| Cell formulas (`cell-formula`) | `SUM([2]Prices!C2:C6)` | typed, copied or moved sheets | formula removed, last calculated value kept |
| Formulas that use an external name | `[2]Prices!B2*(1+VAT)` (via name `VAT`) | names copied along with a sheet | value kept |
| Defined names (`defined-name`), also hidden ones | `OldTotal (hidden) = [1]Budget!$B$10` | copied sheets, old add-ins, Name Manager | name deleted |
| Table calculated columns (`table-formula`) | `Orders[Net] = Orders[[#This Row],[Qty]]*[2]Prices!$B$2` | table copied from another file | column formula deleted, values kept |
| Conditional formatting (`conditional-format`) | `B2>[prices.xlsm]Prices!#REF!` | files from older Excel versions or other tools (Excel 365 refuses to create them) | rule deleted |
| Data validation lists (`data-validation`) | `'C:\work\examples\[prices.xlsm]Prices'!#REF!` | files from older Excel versions or other tools (Excel 365 refuses to create them) | rule deleted |
| Sparklines (`sparkline`) | `[2]Prices!B2:B6` | copied sheets | sparkline deleted |
| Chart series, categories, titles (`chart`) | `[2]Prices!$B$2:$B$6` (series 1 values) | charts copied from another workbook | chart keeps the data it showed, link removed |
| Pivot table sources (`pivot-cache`) | `Prices!A1:C6` in `prices.xlsm` | pivot built on another workbook | reported; change it in Excel |
| Macros on buttons and shapes (`shape-macro`) | `[2]!UpdateReport` | buttons copied from a macro workbook | macro assignment cleared |
| Form controls and text boxes (`object-link`) | drop-down input range `[2]Prices!$A$2:$A$6` | copied controls, linked text boxes | link cleared |
| Linked OLE objects (`object-link`) | linked Word or Excel object | Paste Special > Paste Link | reported; delete or re-insert it in Excel |
| Data connections and Power Query (`connection`) | `File.Contents("C:\Data\sales.csv")` | Get Data, text imports, ODBC / OLE DB | reported only, not a workbook link |
| External hyperlinks (`hyperlink`) | `https://example.com/pricing-policy` | Insert > Link | deleted (keep them with `--keep-hyperlinks`) |

More detail, and how to fix each one by hand in Excel: [docs/where-excel-hides-links.md](docs/where-excel-hides-links.md).

## Remove links safely (clean)

```text
$ excel-link-finder clean examples/sales-report.xlsx -o sales-report_clean.xlsx
Cleaned copy written to sales-report_clean.xlsx
Removed:
  defined-name          3  names deleted
  cell-formula         15  formula cells now hold their last calculated value
  hyperlink             1  external hyperlinks deleted
  conditional-format    1  conditional formatting rules deleted
  data-validation       1  data validation rules deleted
  sparkline             1  sparklines deleted
  object-link           2  shape and form control links cleared
  shape-macro           2  macro assignments cleared
  chart                 4  chart references replaced by the data the chart already showed
  table-formula         1  table column formulas deleted (the cells keep their values)
  external-link         2  external link parts deleted (old-budget.xlsx, prices.xlsm)
  calcChain.xml            removed, Excel rebuilds it on the next save
Not removed:
  pivot-cache        Report > StockPivot                           change it in Excel: PivotTable Analyze > Change Data Source
  connection         sales (text file)                             not a link: manage it in Data > Queries & Connections
  connection         SalesCsv (Power Query)                        not a link: manage it in Data > Queries & Connections
Check: the cleaned copy still has 1 data link location(s); run 'excel-link-finder scan' on it for details.
```

- The original file is **never modified**; `clean` refuses to write over it. Without `-o` the copy is
  `NAME_clean.xlsx` next to the input. The copy keeps the input's extension (.xlsx stays .xlsx, .xlsm stays .xlsm).
- `--only TYPE ...` cleans only some types, for example `--only cell-formula defined-name`. Link parts that are
  still used stay, and the remaining `[N]` link numbers are renumbered everywhere so the file stays consistent.
- `--keep-hyperlinks` keeps external hyperlinks (they do not trigger Excel's link warning).
- The cleaned copy is checked before `clean` finishes: every relationship target exists, every part has a
  content type, every `r:id` resolves, and a new scan lists what is left. Exit code 0 means no data links are left.

Cleaned copies of the example workbooks were opened with Microsoft Excel 365 (16.0, Windows) without any
repair prompt, and Excel reported no external links for them (see [How it works](#how-it-works)).

## Scan a whole folder

```bash
excel-link-finder scan Finance/2025 --recursive
excel-link-finder scan Finance/2025 -r --csv -o links.csv   # one CSV row per finding, opens in Excel
```

```powershell
excel-link-finder scan \\server\share\Reports --recursive   # network shares work too (Windows)
```

Folders are searched for .xlsx, .xlsm, .xltx and .xltm files; Office lock files (`~$...`) are ignored and old
.xls / .xlsb files are listed as skipped. `--quiet` prints one summary per workbook.

## Use in CI / batch

Exit codes: **0** = no external links, **1** = external links found, **2** = a file could not be read.
Connections and hyperlinks do not count unless you add `--strict`. `--no-check` skips the broken-link check.

```yaml
# GitHub Actions: fail the build when a committed workbook links to another file
- run: pip install git+https://github.com/Shahnawaz75/excel-link-finder
- run: excel-link-finder scan templates/ --recursive
```

`--json` gives one document with a `files` list (`file`, `error`, `has_links`, `summary`, `findings`) and
`totals`; each finding has `type`, `sheet`, `object`, `ref`, `text`, `links`, `target`, `status`, `note`.
`--csv` writes one row per finding with the same fields plus `file` (UTF-8 with BOM when written with `-o`,
so Excel opens it correctly).

## Python API

```python
from excel_link_finder import clean, scan

result = scan("book.xlsx")
print(result.has_links, result.counts())
for finding in result.links:          # data links (no connections / hyperlinks)
    print(finding.type, finding.sheet, finding.object, finding.ref, finding.text, finding.target)
for finding in result.broken:
    print("missing:", finding.target)

cleaned = clean("book.xlsx", "book_clean.xlsx", keep_hyperlinks=False, only=None)
print(cleaned.removed, cleaned.is_clean)
for finding in cleaned.kept:          # things to fix in Excel, e.g. pivot table sources
    print(finding.type, finding.object)
```

See [docs/python-api.md](docs/python-api.md) for all fields and functions.

## How it works

An .xlsx / .xlsm file is a zip package of XML parts. excel-link-finder opens it with Python's `zipfile`,
follows the relationship (`.rels`) files from the workbook to every sheet, drawing, chart, table, pivot cache
and control, and parses each part with the standard library's expat parser. A small tokenizer reads the
formulas: Excel stores an external reference as `[N]Sheet!A1`, where `N` points at
`xl/externalLinks/externalLinkN.xml`; full paths such as `'C:\Data\[Book.xlsx]Sheet1'!A1`, `[0]` (the
workbook itself) and table references such as `Orders[[#This Row],[Qty]]` are told apart.

The cleaner never re-writes XML with a serializer, because that renames namespace prefixes and loses
declarations Excel needs, which is a classic cause of the repair prompt. It splices small changes into the
original bytes instead, removes containers that would be left empty, fixes `count` attributes, renumbers
`[N]` references in every part when only some links are removed, deletes `calcChain.xml` when formulas
change, and removes link parts together with their relationships and `[Content_Types].xml` entries.

Validation: the tests run on Linux and Windows without Excel (structure checks, re-scans, openpyxl). In
addition, the example workbooks were generated by Microsoft Excel 365 itself ([examples/make_examples.py](examples/make_examples.py)),
and cleaned copies of them, of the synthetic test workbook and of Strict Open XML and .xlsm variants were
opened in Excel with [tools/excel_check.py](tools/excel_check.py), which fails when Excel repairs a file.
Deliberately damaged files were rejected by that check, so it does detect broken output.

## Limitations

- **.xls, .xlsb and password-protected workbooks are not supported.** Save the file as .xlsx first.
- **Pivot table sources, data connections, Power Query and linked OLE objects are reported, not removed.**
  Changing them needs Excel (PivotTable Analyze > Change Data Source; Data > Queries & Connections).
- **VBA code is not inspected.** Macros that open other workbooks are not listed.
- **Charts become static** after `clean`: they keep showing the values they had, but no longer update.
- Formulas without a saved value (files written by some libraries) become empty cells; `clean` says how many.
- When a defined name is deleted, formulas that use it are converted to values too (they are listed as
  `cell-formula (via name ...)`). With `--only defined-name` those formulas stay and will show `#NAME?`.
- The broken-link check looks at local and network paths only; URLs (SharePoint, OneDrive, web) are never
  fetched, and Windows paths cannot be checked on Linux or macOS ("not checked").
- DDE links are detected; they were not tested with real DDE workbooks.
- Each sheet is read into memory, so very large workbooks need memory in proportion to their largest sheet.
- Excel validation was done with Excel 365 on Windows. Older Excel versions, LibreOffice and Google Sheets
  were not tested; the CI does not run Excel.

## FAQ

### How do I find external links in Excel?

In Excel you would check Data > Edit Links (Workbook Links in Excel 365), search formulas for `[` or `.xl`,
then go through Name Manager, the Conditional Formatting Rules Manager, every data validation rule, every
chart, every pivot table and every button. `excel-link-finder scan book.xlsx` lists all of these at once,
with the sheet and cell of each one.

### Why can't I break links in Excel?

Break Link converts formulas to values, but links stored in other places are left behind: in our test with
Excel 365 a conditional formatting rule, a drop-down list's input range and a pivot table source. Edit Links
then keeps showing the source (Break Link is also unavailable while a sheet is protected). Run `scan` to see
what is left, then `clean`.

### How do I remove the "links to external sources" security warning?

Excel shows *"This workbook contains links to one or more external sources that could be unsafe"* because the
workbook has external links. Remove them with `excel-link-finder clean book.xlsx -o book_clean.xlsx`. For the
cleaned example workbooks Excel itself reported no external links, so there is nothing left to warn about.

### What does "We can't update some of the links in your workbook" mean?

A linked file was moved, renamed or deleted. `scan` marks those links as **MISSING** and shows the path Excel
last saw. Put the file back, or remove the links with `clean`.

### Does it work without Excel or on Linux/Mac?

Yes. It is pure Python (standard library only) and reads the file directly, so it runs on Windows, macOS,
Linux, servers, Docker and CI without Excel or LibreOffice.

### Does it work with .xls files?

No. Old binary .xls files (Excel 97-2003) use a different format. Open the file in Excel, choose
File > Save As > Excel Workbook (*.xlsx), then scan the new file. The tool tells you this when you give it
an .xls file. .xlsb (binary workbook) is not supported either.

### Does it change or upload my workbook?

No. `scan` only reads the file and `clean` writes a new copy. The tool never connects to the network: link
targets are only checked on the local disk or network share.

## Contributing

Bug reports with a small sample workbook are very welcome (remove confidential data first), as are new link
locations and fixes. See [CONTRIBUTING.md](CONTRIBUTING.md) for the development setup and how to test with Excel.

## License

[MIT](LICENSE). See also the author's other tools: [dxf-diff](https://github.com/Shahnawaz75/dxf-diff)
(compare DXF drawings) and [pyinstaller-doctor](https://github.com/Shahnawaz75/pyinstaller-doctor).

If excel-link-finder saved you a few hours of link hunting, a star on GitHub is a nice way to say thanks.
