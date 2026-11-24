---
title: Where Excel hides external links
description: Every place an .xlsx or .xlsm workbook can keep a link to another file - formulas, hidden names, conditional formatting, data validation, charts, pivot tables, buttons, controls - and how to find and remove each one.
---

# Where Excel hides external links

An external link is a reference from one workbook to another file. In the .xlsx file, each linked
workbook gets a part `xl/externalLinks/externalLinkN.xml` with a cached copy of the values used, and
formulas refer to it by number: `[2]Prices!B2` means "cell B2 of sheet Prices in external link 2". As long as
anything still refers to a link, Excel keeps it, shows it in **Data > Edit Links** (called **Workbook Links**
in Excel 365) and warns about it when the file is opened.

The table lists each place, with the `excel-link-finder scan` type, how to find it by hand in Excel and what
`excel-link-finder clean` does with it.

| Type | Where in Excel | Find it by hand | `clean` |
| --- | --- | --- | --- |
| `external-link` | the link itself | Data > Edit Links / Workbook Links | part, relationships and content type deleted when nothing uses the link |
| `cell-formula` | cell formulas, including array formulas | Ctrl+F, look in *Formulas*, search for `[` or `.xl` | formula removed, last value kept |
| `defined-name` | Formulas > Name Manager | Name Manager; **hidden names are not shown there** | name deleted |
| `table-formula` | calculated columns and totals rows of tables | click a cell of each table column | column formula deleted, values kept |
| `conditional-format` | Home > Conditional Formatting > Manage Rules | rules manager, for *This Worksheet*, sheet by sheet | rule deleted |
| `data-validation` | Data > Data Validation | select cells with validation (F5 > Special > Data validation) | rule deleted |
| `sparkline` | sparkline groups | Sparkline > Edit Data | sparkline deleted |
| `chart` | series values, categories, series names, chart and axis titles | select each series and look at the `=SERIES(...)` formula | chart keeps its cached data; the link is removed |
| `pivot-cache` | PivotTable source data | PivotTable Analyze > Change Data Source | reported; change it in Excel |
| `shape-macro` | macros assigned to buttons, shapes, pictures, charts | right-click > Assign Macro, for every object | assignment cleared |
| `object-link` | form control input ranges and cell links, text boxes linked to a cell, linked OLE objects | Format Control; the formula bar of each text box | link cleared (linked OLE objects are only reported) |
| `connection` | Data > Queries & Connections (text files, databases, web, Power Query) | Queries & Connections pane | reported only |
| `hyperlink` | links to web pages or files | right-click > Edit Hyperlink | deleted (keep them with `--keep-hyperlinks`) |

## How the links get there

Most hidden links are created by **copying or moving a sheet from another workbook**. Excel then rewrites
every reference on the copied sheet so that it still points to the original file: formulas, chart series,
table formulas, sparklines, form controls and the pivot table source. Names used by those formulas are copied
too, so the new workbook gets names such as `VAT = [2]Prices!$F$1`. Buttons copied from a macro workbook keep
running the macro in that workbook (`[2]!UpdateReport`).

Conditional formatting and data validation rules that point to another workbook come from older files and other
tools; Excel 365 itself refuses to create them. When Excel 365 saves such a rule it rewrites it with the file
name and `#REF!`, for example `B2>[prices.xlsm]Prices!#REF!`, which still counts as a link.

## Why Break Link does not remove them

Edit Links > Break Link converts formulas to values, but in our test (Excel 365, see the
[example output](example-output.md#what-excels-break-link-leaves-behind)) it left a conditional formatting rule,
a drop-down list's input range and the pivot table source, so Excel kept reporting the link. Break Link is also
unavailable while a sheet is protected.

## Broken links

A link is **broken** when the linked file is not there any more: it was moved, renamed or deleted, or a network
drive is not connected. Excel then says *"We can't update some of the links in your workbook"*.
`excel-link-finder scan` checks local and network paths and marks those links **MISSING**, with the path Excel
last saw. Links to web addresses (SharePoint, OneDrive) are shown as URLs and never fetched.

## Not covered

- VBA code that opens other workbooks (the VBA project is binary and is not inspected).
- Old .xls and .xlsb files: save them as .xlsx first.
