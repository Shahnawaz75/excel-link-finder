r"""Build the example workbooks in this folder with Microsoft Excel (Windows + ``pip install pywin32``).

    python examples/make_examples.py

Creates:

* ``prices.xlsm``: the source workbook (sheet "Prices" and the name VAT). It contains no VBA
  code; the buttons in the report only *refer* to macro names in it, as copied buttons do.
* ``sales-report.xlsx``: a report that links to ``prices.xlsm`` and to ``old-budget.xlsx``.
  ``old-budget.xlsx`` is deleted on purpose, so that link is broken.

Everything is made by Excel through COM automation in a private, hidden Excel instance. The
report sheet is built inside ``prices.xlsm`` and then copied to a new workbook, which is how
most hidden links are born: Excel rewrites every reference of the copied sheet (formulas,
chart series, sparklines, table formulas, form controls, pivot table) to point back at the
source file. Two things are done to the saved files afterwards:

1. Excel 365 refuses to create conditional formatting or data validation rules that point to
   another workbook, but files from older versions and from copied sheets contain them. One
   rule of each kind is added in the format Excel 2010+ writes (x14 extensions), then Excel
   opens and saves the file again, so the final file is still written by Excel.
2. The temporary folder name is replaced by ``C:\work\examples`` inside the files, so no
   personal paths end up in the repository.
"""

from __future__ import annotations

import re
import shutil
import sys
import tempfile
import zipfile
from pathlib import Path

HERE = Path(__file__).resolve().parent
NEUTRAL = r"C:\work\examples"
XLSX, XLSM = 51, 52
X14 = "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main"
XM = "http://schemas.microsoft.com/office/excel/2006/main"
RULES = (
    f'<ext uri="{{78C0D931-6437-407d-A8EE-F0AAD7539E65}}" xmlns:x14="{X14}"><x14:conditionalFormattings>'
    f'<x14:conditionalFormatting xmlns:xm="{XM}"><x14:cfRule type="expression" priority="1" '
    'id="{6E9B3C1A-5D2F-4A8B-9C7E-1F2A3B4C5D6E}"><xm:f>B2&gt;[1]Prices!$B$2</xm:f><x14:dxf><fill><patternFill>'
    '<bgColor rgb="FFFFC7CE"/></patternFill></fill></x14:dxf></x14:cfRule><xm:sqref>B2:B6</xm:sqref>'
    "</x14:conditionalFormatting></x14:conditionalFormattings></ext>"
    f'<ext uri="{{CCE6A557-97BC-4b89-ADB6-D9C93CAAB3DF}}" xmlns:x14="{X14}"><x14:dataValidations count="1" '
    f'xmlns:xm="{XM}"><x14:dataValidation type="list" allowBlank="1" showInputMessage="1" showErrorMessage="1">'
    "<x14:formula1><xm:f>[1]Prices!$A$2:$A$6</xm:f></x14:formula1><xm:sqref>D2:D6</xm:sqref>"
    "</x14:dataValidation></x14:dataValidations></ext>"
)


def new_book(excel, sheet_name: str):
    book = excel.Workbooks.Add()
    while book.Worksheets.Count > 1:
        book.Worksheets(book.Worksheets.Count).Delete()
    book.Worksheets(1).Name = sheet_name
    book.RemovePersonalInformation = True  # no author names in the saved files
    return book


def build_report_sheet(source, prices, missing):
    """Add a "Report" sheet to the source workbook that refers to its "Prices" sheet."""
    report = source.Worksheets.Add(missing, prices)
    report.Name = "Report"
    report.Range("A1").Value = "Price report"
    report.Range("A2").Formula = "=Prices!A2"
    report.Range("A2:A6").FillDown()
    report.Range("B2").Formula = "=Prices!B2*(1+VAT)"
    report.Range("B2:B6").FillDown()
    report.Range("C2").Formula = "=SUM(Prices!C2:C6)"
    chart = report.Shapes.AddChart2(-1, 51, 330, 20, 360, 220)  # clustered column chart
    chart.Name = "Price chart"
    chart.Chart.SetSourceData(prices.Range("A1:B6"))
    chart.Chart.HasTitle = True
    chart.Chart.ChartTitle.Formula = "=Prices!$B$1"
    report.Range("E2").SparklineGroups.Add(1, "Prices!B2:B6")
    box = report.Shapes.AddTextbox(1, 330, 260, 120, 30)
    box.Name = "VAT box"
    box.DrawingObject.Formula = "=Prices!$F$1"
    drop = report.DropDowns().Add(20, 300, 120, 20)
    drop.ListFillRange = "Prices!$A$2:$A$6"
    drop.LinkedCell = "$F$2"
    for row, qty in enumerate((20, 30, 40), start=2):
        report.Cells(row, 8).Value = qty
    report.Range("H1").Value, report.Range("I1").Value = "Qty", "Net"
    table = report.ListObjects.Add(1, report.Range("H1:I4"), None, 1)
    table.Name = "Orders"
    report.Range("I2").Formula = "=[@Qty]*Prices!$B$2"  # becomes a calculated column
    cache = source.PivotCaches().Create(1, "Prices!R1C1:R6C3")
    pivot = cache.CreatePivotTable(report.Range("K1"), "StockPivot")
    pivot.PivotFields("Product").Orientation = 1
    pivot.AddDataField(pivot.PivotFields("Stock"))
    return report


def build(work: Path) -> None:
    import pythoncom  # type: ignore[import-not-found]
    import win32com.client  # type: ignore[import-not-found]

    excel = win32com.client.DispatchEx("Excel.Application")  # a new, private Excel process
    excel.Visible = False
    excel.DisplayAlerts = False
    try:
        source = new_book(excel, "Prices")
        prices = source.Worksheets("Prices")
        rows = [
            ("Product", "Price", "Stock"),
            ("Bolt M8", 0.25, 1200),
            ("Nut M8", 0.1, 3000),
            ("Washer M8", 0.05, 5000),
            ("Bracket", 3.8, 150),
            ("Hinge", 2.4, 320),
        ]
        for r, values in enumerate(rows, start=1):
            for c, value in enumerate(values, start=1):
                prices.Cells(r, c).Value = value
        prices.Range("E1").Value, prices.Range("F1").Value = "VAT", 0.2
        source.Names.Add("VAT", "=Prices!$F$1")
        source.SaveAs(str(work / "prices.xlsm"), XLSM)  # saved before the Report sheet is added

        build_report_sheet(source, prices, pythoncom.Missing).Copy()  # copy to a new workbook: references become links
        book = excel.ActiveWorkbook
        book.RemovePersonalInformation = True
        report = book.Worksheets("Report")
        shape = report.Shapes.AddShape(1, 20, 260, 120, 30)
        shape.Name = "Refresh shape"
        shape.TextFrame2.TextRange.Text = "Refresh"
        shape.OnAction = "'prices.xlsm'!RefreshPrices"
        button = report.Buttons().Add(160, 260, 120, 30)
        button.Caption = "Update"
        button.OnAction = "'prices.xlsm'!UpdateReport"

        old = new_book(excel, "Budget")
        old.Worksheets(1).Range("A10").Value, old.Worksheets(1).Range("B10").Value = "Total", 12345
        old.SaveAs(str(work / "old-budget.xlsx"), XLSX)
        summary = book.Worksheets.Add(pythoncom.Missing, report)
        summary.Name = "Summary"
        summary.Range("A1").Value = "Budget last year"
        summary.Range("B1").Formula = "='[old-budget.xlsx]Budget'!$B$10"
        book.Names.Add("TaxRate", "=[prices.xlsm]Prices!$F$1")
        book.Names.Add("OldTotal", "='[old-budget.xlsx]Budget'!$B$10").Visible = False
        summary.Hyperlinks.Add(summary.Range("A3"), "https://example.com/pricing-policy", "", "", "Pricing policy")
        (work / "sales.csv").write_text("Month,Sales\nJan,100\nFeb,120\n", encoding="utf-8")
        query = summary.QueryTables.Add(f"TEXT;{work / 'sales.csv'}", summary.Range("D1"))
        query.TextFileCommaDelimiter = True
        query.Refresh(False)
        book.Queries.Add(
            "SalesCsv",
            'let\n    Source = Csv.Document(File.Contents("C:\\Data\\sales.csv"), '
            '[Delimiter=",", Encoding=65001])\nin\n    Source',
        )
        book.SaveAs(str(work / "sales-report.xlsx"), XLSX)
        for workbook in (book, old, source):
            workbook.Close(False)
        (work / "old-budget.xlsx").unlink()  # this link is broken on purpose

        add_rules(work / "sales-report.xlsx")
        book = excel.Workbooks.Open(str(work / "sales-report.xlsx"), UpdateLinks=0)
        book.Save()  # let Excel write the injected rules itself
        book.Close(False)
    finally:
        excel.Quit()


def rewrite(path: Path, change) -> None:
    """Rewrite some parts of a zip package in place; ``change(name, data)`` returns new bytes."""
    with zipfile.ZipFile(path) as zin:
        items = [(info, change(info.filename, zin.read(info))) for info in zin.infolist()]
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as zout:
        for info, data in items:
            zout.writestr(info, data)


def add_rules(path: Path) -> None:
    """Add an external conditional format and data validation to the Report sheet."""
    with zipfile.ZipFile(path) as zf:
        workbook = zf.read("xl/workbook.xml").decode()
        rels = zf.read("xl/_rels/workbook.xml.rels").decode()
    rid = re.search(r'<sheet name="Report"[^>]*r:id="([^"]+)"', workbook).group(1)
    part = "xl/" + re.search(rf'Id="{rid}"[^>]*Target="([^"]+)"|Target="([^"]+)"[^>]*Id="{rid}"', rels).group(1)

    def change(name: str, data: bytes) -> bytes:
        if name != part:
            return data
        text = data.decode()
        text = (
            text.replace("<extLst>", "<extLst>" + RULES, 1)
            if "<extLst>" in text
            else text.replace("</worksheet>", f"<extLst>{RULES}</extLst></worksheet>")
        )
        return text.encode()

    rewrite(path, change)


def sanitize(path: Path, work: Path) -> None:
    """Replace the temporary folder with a neutral one and make sure no user path is left."""
    names = {str(work): NEUTRAL, str(work).replace("\\", "/"): NEUTRAL.replace("\\", "/")}

    def change(_name: str, data: bytes) -> bytes:
        for old, new in names.items():
            data = re.sub(re.escape(old).encode(), new.replace("\\", "\\\\").encode(), data, flags=re.I)
        return data

    rewrite(path, change)
    with zipfile.ZipFile(path) as zf:
        for name in zf.namelist():
            if re.search(rb"[\\/]users[\\/]|%5cusers%5c", zf.read(name), re.I):
                raise SystemExit(f"{path.name}:{name} still contains a user folder path")


def main() -> None:
    if sys.platform != "win32":
        sys.exit("make_examples.py needs Windows with Microsoft Excel and pywin32")
    work = Path(tempfile.mkdtemp(prefix="elf-"))  # short path: Excel rejects paths over 218 characters
    try:
        build(work)
        for name in ("prices.xlsm", "sales-report.xlsx"):
            sanitize(work / name, work)
            shutil.copyfile(work / name, HERE / name)
            print(f"wrote {HERE.name}/{name}")
    finally:
        shutil.rmtree(work, ignore_errors=True)


if __name__ == "__main__":
    main()
