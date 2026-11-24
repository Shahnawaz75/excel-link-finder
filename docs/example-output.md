---
title: Example output
description: Real output of excel-link-finder scan and clean on an example workbook made by Microsoft Excel, and what Excel's own Break Link leaves behind.
---

# Example output

Everything on this page is the real output of excel-link-finder 0.1.0 (run on Windows, so paths are shown
with backslashes) on the workbooks in [`examples/`](https://github.com/Shahnawaz75/excel-link-finder/tree/main/examples).
They were made by Microsoft Excel 365 with [`examples/make_examples.py`](https://github.com/Shahnawaz75/excel-link-finder/blob/main/examples/make_examples.py):

- `prices.xlsm` is the source workbook,
- `sales-report.xlsx` links to `prices.xlsm` and to `old-budget.xlsx`, which was deleted on purpose (a broken link).

`C:\work\examples` is the neutral folder name the example script writes into the files instead of the real
temporary folder.

## Scan one workbook

```text
$ excel-link-finder scan examples/sales-report.xlsx
examples\sales-report.xlsx
  TYPE                SHEET / NAME / OBJECT                 CELL/RANGE  FORMULA / TEXT                        LINKED FILE
  external-link       [1] (workbook link, last seen at ...  -           externalLink1.xml                     old-budget.xlsx (MISSING)
  external-link       [2] (workbook link)                   -           externalLink2.xml                     prices.xlsm (found)
  defined-name        OldTotal (hidden)                     -           [1]Budget!$B$10                       [1] old-budget.xlsx
  defined-name        TaxRate                               -           [2]Prices!$F$1                        [2] prices.xlsm
  defined-name        VAT                                   -           [2]Prices!$F$1                        [2] prices.xlsm
  cell-formula        Summary                               B1          [1]Budget!$B$10                       [1] old-budget.xlsx
  hyperlink           Summary                               A3          -                                     https://example.com/pricing-policy
  cell-formula        Report                                A2          [2]Prices!A2                          [2] prices.xlsm
  cell-formula        Report (via name VAT)                 B2          [2]Prices!B2*(1+VAT)                  [2] prices.xlsm
  cell-formula        Report                                C2          SUM([2]Prices!C2:C6)                  [2] prices.xlsm
  cell-formula        Report                                I2          Orders[[#This Row],[Qty]]*[2]Pric...  [2] prices.xlsm
  cell-formula        Report                                A3          [2]Prices!A3                          [2] prices.xlsm
  cell-formula        Report (via name VAT)                 B3          [2]Prices!B3*(1+VAT)                  [2] prices.xlsm
  cell-formula        Report                                I3          Orders[[#This Row],[Qty]]*[2]Pric...  [2] prices.xlsm
  cell-formula        Report                                A4          [2]Prices!A4                          [2] prices.xlsm
  cell-formula        Report (via name VAT)                 B4          [2]Prices!B4*(1+VAT)                  [2] prices.xlsm
  cell-formula        Report                                I4          Orders[[#This Row],[Qty]]*[2]Pric...  [2] prices.xlsm
  cell-formula        Report                                A5          [2]Prices!A5                          [2] prices.xlsm
  cell-formula        Report (via name VAT)                 B5          [2]Prices!B5*(1+VAT)                  [2] prices.xlsm
  cell-formula        Report                                A6          [2]Prices!A6                          [2] prices.xlsm
  cell-formula        Report (via name VAT)                 B6          [2]Prices!B6*(1+VAT)                  [2] prices.xlsm
  conditional-format  Report (expression rule)              B2:B6       B2>[prices.xlsm]Prices!#REF!          prices.xlsm
  data-validation     Report (list rule)                    D2:D6       'C:\work\examples\[prices.xlsm]Pr...  C:\work\examples\prices.xlsm
  sparkline           Report                                E2          [2]Prices!B2:B6                       [2] prices.xlsm
  object-link         Report > Drop Down 1 (form contro...  -           [2]Prices!$A$2:$A$6                   [2] prices.xlsm
  shape-macro         Report > Button 2 (form control m...  -           [2]!UpdateReport                      [2] prices.xlsm
  chart               Report > Price chart (chart title)    -           [2]Prices!$B$1                        [2] prices.xlsm
  chart               Report > Price chart (series 1 name)  -           [2]Prices!$B$1                        [2] prices.xlsm
  chart               Report > Price chart (series 1 ca...  -           [2]Prices!$A$2:$A$6                   [2] prices.xlsm
  chart               Report > Price chart (series 1 va...  -           [2]Prices!$B$2:$B$6                   [2] prices.xlsm
  object-link         Report > VAT box (shape text link...  -           [2]Prices!$F$1                        [2] prices.xlsm
  shape-macro         Report > Refresh shape (macro ass...  -           [2]!RefreshPrices                     [2] prices.xlsm
  table-formula       Report > Orders[Net] (calculated ...  H1:I4       Orders[[#This Row],[Qty]]*[2]Pric...  [2] prices.xlsm
  pivot-cache         Report > StockPivot                   A1:C6       Prices!A1:C6                          prices.xlsm (found)
  connection          sales (text file)                     -           C:\work\examples\sales.csv            C:\work\examples\sales.csv
  connection          SalesCsv (Power Query)                -           File.Contents("C:\Data\sales.csv")    C:\Data\sales.csv

  Found: external-link 2 (1 broken), defined-name 3, cell-formula 15, table-formula 1, conditional-format 1,
    data-validation 1, sparkline 1, chart 4, pivot-cache 1, shape-macro 2, object-link 2
  Also found (not data links): connection 2, hyperlink 1
[exit code 1]
```

Things to notice:

- `[1]` and `[2]` are the external link parts. Formulas refer to them by number (`[2]Prices!A2`).
- `old-budget.xlsx` is **MISSING**: that is the "We can't update some of the links in your workbook" case.
- `OldTotal` is a **hidden** name that Excel's Name Manager does not show.
- The `B2:B6` formulas also use the name `VAT`, which points to the other workbook ("via name VAT").
- The conditional formatting and data validation rules point at the source by file name and full path,
  with `#REF!`: this is how Excel 365 rewrites such rules when it saves them. (Excel 365 refuses to create
  these two rules, so the example script adds them in the Excel 2010 file format and then lets Excel open
  and save the file.)
- The drop-down list, the button macro, the text box and the pivot table all depend on `prices.xlsm` too.

## Scan a folder

```text
$ excel-link-finder scan examples
examples\prices.xlsm
  No external links found.

examples\sales-report.xlsx
  Found: external-link 2 (1 broken), defined-name 3, cell-formula 15, table-formula 1, conditional-format 1,
    data-validation 1, sparkline 1, chart 4, pivot-cache 1, shape-macro 2, object-link 2
  Also found (not data links): connection 2, hyperlink 1

Scanned 2 workbooks: 1 with external links, 1 without, 0 could not be read.
[exit code 1]
```

## Clean

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
[exit code 1]
```

The exit code is 1 because the pivot table source still points to `prices.xlsm`; that has to be changed in
Excel. Data connections are listed for information: they are not workbook links.

```text
$ excel-link-finder scan sales-report_clean.xlsx
sales-report_clean.xlsx
  TYPE         SHEET / NAME / OBJECT   CELL/RANGE  FORMULA / TEXT                      LINKED FILE
  pivot-cache  Report > StockPivot     A1:C6       Prices!A1:C6                        prices.xlsm (MISSING)
  connection   sales (text file)       -           C:\work\examples\sales.csv          C:\work\examples\sales.csv
  connection   SalesCsv (Power Query)  -           File.Contents("C:\Data\sales.csv")  C:\Data\sales.csv

  Found: pivot-cache 1 (1 broken)
  Also found (not data links): connection 2
[exit code 1]
```

The pivot cache source is now **MISSING** only because the cleaned copy was written to another folder than
`prices.xlsm`.

## What Excel's Break Link leaves behind

To see why "Break Links is not working" is such a common complaint, Excel's own Break Link was run on every
link source of `sales-report.xlsx` through COM automation (the same as Data > Edit Links / Workbook Links >
Break Link), and the result was saved as `after-break-link.xlsx`:

```text
links before: ['old-budget.xlsx', 'prices.xlsm', 'prices.xlsm']
broke old-budget.xlsx
broke prices.xlsm
broke prices.xlsm
links Excel still sees after Break Link: ['prices.xlsm']
```

Excel still sees a link to `prices.xlsm`. The scan shows where it is:

```text
$ excel-link-finder scan after-break-link.xlsx
after-break-link.xlsx
  TYPE                SHEET / NAME / OBJECT                 CELL/RANGE  FORMULA / TEXT                      LINKED FILE
  external-link       [1] (workbook link)                   -           externalLink1.xml                   prices.xlsm (found)
  hyperlink           Summary                               A3          -                                   https://example.com/pricing-policy
  conditional-format  Report (expression rule)              B2:B6       B2>[prices.xlsm]Prices!#REF!        prices.xlsm
  object-link         Report > Drop Down 1 (form contro...  -           [1]Prices!$A$2:$A$6                 [1] prices.xlsm
  pivot-cache         Report > StockPivot                   A1:C6       Prices!A1:C6                        prices.xlsm (found)
  connection          sales (text file)                     -           C:\work\examples\sales.csv          C:\work\examples\sales.csv
  connection          SalesCsv (Power Query)                -           File.Contents("C:\Data\sales.csv")  C:\Data\sales.csv

  Found: external-link 1, conditional-format 1, pivot-cache 1, object-link 1
  Also found (not data links): connection 2, hyperlink 1
[exit code 1]
```

`clean` removes the rest:

```text
$ excel-link-finder clean after-break-link.xlsx -o after-break-link_clean.xlsx
Cleaned copy written to after-break-link_clean.xlsx
Removed:
  hyperlink             1  external hyperlinks deleted
  conditional-format    1  conditional formatting rules deleted
  object-link           1  shape and form control links cleared
  external-link         1  external link parts deleted (prices.xlsm)
Not removed:
  pivot-cache        Report > StockPivot                           change it in Excel: PivotTable Analyze > Change Data Source
  connection         sales (text file)                             not a link: manage it in Data > Queries & Connections
  connection         SalesCsv (Power Query)                        not a link: manage it in Data > Queries & Connections
Check: the cleaned copy still has 1 data link location(s); run 'excel-link-finder scan' on it for details.
[exit code 1]
```

## Checked in Excel

[`tools/excel_check.py`](https://github.com/Shahnawaz75/excel-link-finder/blob/main/tools/excel_check.py)
opens each file in a private, hidden Excel instance (read-only, without updating links), fails if Excel had to
repair it, and prints the links Excel itself still sees (`Workbook.LinkSources`):

```text
$ python tools/excel_check.py examples/sales-report.xlsx sales-report_clean.xlsx after-break-link.xlsx after-break-link_clean.xlsx
OK        sales-report.xlsx: Excel 16.0, external links Excel sees: old-budget.xlsx, prices.xlsm, prices.xlsm
OK        sales-report_clean.xlsx: Excel 16.0, external links Excel sees: none
OK        after-break-link.xlsx: Excel 16.0, external links Excel sees: prices.xlsm
OK        after-break-link_clean.xlsx: Excel 16.0, external links Excel sees: none
[exit code 0]
```

## JSON (first lines)

```json
$ excel-link-finder scan --json examples/sales-report.xlsx
{
  "tool": "excel-link-finder",
  "version": "0.1.0",
  "files": [
    {
      "file": "examples\\sales-report.xlsx",
      "error": null,
      "has_links": true,
      "summary": {
        "external-link": 2,
        "defined-name": 3,
        "cell-formula": 15,
        "table-formula": 1,
        "conditional-format": 1,
        "data-validation": 1,
        "sparkline": 1,
        "chart": 4,
        "pivot-cache": 1,
        "shape-macro": 2,
        "object-link": 2,
        "connection": 2,
        "hyperlink": 1,
        "broken": 1
      },
      "findings": [
        {
          "type": "external-link",
          "sheet": "",
          "object": "[1]",
          "ref": "",
          "text": "externalLink1.xml",
          "links": [
            1
          ],
          "target": "old-budget.xlsx",
          "status": "missing",
          "note": "workbook link, last seen at C:\\work\\examples\\old-budget.xlsx"
        },
        {
...
```

## CSV (first lines)

```text
$ excel-link-finder scan --csv examples/sales-report.xlsx
file,type,sheet,object,ref,text,links,target,status,note
examples\sales-report.xlsx,external-link,,[1],,externalLink1.xml,1,old-budget.xlsx,missing,"workbook link, last seen at C:\work\examples\old-budget.xlsx"
examples\sales-report.xlsx,external-link,,[2],,externalLink2.xml,2,prices.xlsm,found,workbook link
examples\sales-report.xlsx,defined-name,,OldTotal,,[1]Budget!$B$10,1,old-budget.xlsx,,hidden
examples\sales-report.xlsx,defined-name,,TaxRate,,[2]Prices!$F$1,2,prices.xlsm,,
examples\sales-report.xlsx,defined-name,,VAT,,[2]Prices!$F$1,2,prices.xlsm,,
examples\sales-report.xlsx,cell-formula,Summary,,B1,[1]Budget!$B$10,1,old-budget.xlsx,,
...
```
