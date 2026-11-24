"""Write small, spec-correct .xlsx packages for the tests, so the test suite does not need Excel.

``kitchen_sink`` builds a workbook with external links in formulas (plain, quoted, full path,
shared, array, dynamic array), defined names (hidden, sheet-level, chained), conditional
formatting and data validation (classic and Excel 2010 ``x14`` forms), a sparkline, hyperlinks,
data connections and an optional Power Query mashup. Charts, shapes, form controls, tables and
pivot tables are covered by the Excel-made workbook in ``examples/``.
"""

from __future__ import annotations

import base64
import io
import re
import struct
import zipfile
from pathlib import Path
from xml.sax.saxutils import escape

MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
X14 = "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main"
XM = "http://schemas.microsoft.com/office/excel/2006/main"
CT = "application/vnd.openxmlformats-officedocument.spreadsheetml"
LINKS = ("https://contoso.sharepoint.com/sites/fin/Rates.xlsx", "Source.xlsx", "missing/Old%20Budget.xlsx")

STYLES = (
    f'<styleSheet xmlns="{MAIN}"><fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts>'
    '<fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills>'
    '<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>'
    '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>'
    '<cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/></cellXfs>'
    '<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>'
    '<dxfs count="1"><dxf><fill><patternFill><bgColor rgb="FFFFC7CE"/></patternFill></fill></dxf></dxfs></styleSheet>'
)
METADATA = (
    f'<metadata xmlns="{MAIN}" xmlns:xda="http://schemas.microsoft.com/office/spreadsheetml/2017/dynamicarray">'
    '<metadataTypes count="1"><metadataType name="XLDAPR" minSupportedVersion="120000" copy="1" pasteAll="1" '
    'pasteValues="1" merge="1" splitFirst="1" rowColShift="1" clearFormats="1" clearComments="1" assign="1" '
    'coerce="1" cellMeta="1"/></metadataTypes><futureMetadata name="XLDAPR" count="1"><bk><extLst>'
    '<ext uri="{bdbb8cdc-fa1e-496e-a857-3c3f30c029c3}"><xda:dynamicArrayProperties fDynamic="1" fCollapsed="0"/>'
    '</ext></extLst></bk></futureMetadata><cellMetadata count="1"><bk><rc t="1" v="0"/></bk></cellMetadata></metadata>'
)


def cell(ref: str, formula: str | None = None, value: object = None, t: str = "", f_attrs: str = "", attrs: str = "") -> str:
    """One ``<c>`` element."""
    f = f"<f{f_attrs}>{escape(formula)}</f>" if formula else (f"<f{f_attrs}/>" if f_attrs else "")
    v = f"<v>{escape(str(value))}</v>" if value is not None else ""
    t_attr = f' t="{t}"' if t else ""
    return f'<c r="{ref}"{t_attr}{attrs}>{f}{v}</c>'


def sheet_xml(cells: list[str], after: str = "", ext: str = "") -> str:
    """A worksheet; ``after`` goes between sheetData and pageMargins, ``ext`` into extLst."""
    rows: dict[int, list[str]] = {}
    for c in cells:
        rows.setdefault(int(re.search(r'r="[A-Z]+(\d+)"', c).group(1)), []).append(c)  # type: ignore[union-attr]
    data = "".join(f'<row r="{n}">{"".join(rows[n])}</row>' for n in sorted(rows))
    return (
        f'<worksheet xmlns="{MAIN}" xmlns:r="{R}" xmlns:x14="{X14}" xmlns:xm="{XM}"><sheetData>{data}</sheetData>'
        f'{after}<pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>'
        f"{f'<extLst>{ext}</extLst>' if ext else ''}</worksheet>"
    )


REPORT_CELLS = [
    cell("A1", "[2]Data!A1", 10),
    cell("B2", "TaxRate*A1", 2, f_attrs=' t="shared" ref="B2:B4" si="0"'),  # Excel rejects [N] in shared formulas
    cell("B3", None, 3, f_attrs=' t="shared" si="0"'),
    cell("B4", None, 4, f_attrs=' t="shared" si="0"'),
    cell("C1", "[2]Data!C1", "hello & bye", t="str"),
    cell("C2", "'[3]My Sheet'!A1", 5),
    cell("C3", "'C:\\Finance\\[Budget 2025.xlsx]Plan'!$B$3", 7),
    cell("D1", "SUM(A1:A3)", 10),
    cell("D2", '"[2]Data!A1"&"x"', "[2]Data!A1x", t="str"),
    cell("D4", "TaxRate*2", 0.4),
    cell("E1", "[2]Data!A1:A2*2", 20, f_attrs=' t="array" ref="E1:E2"'),
    cell("E2", None, 22),
    cell("F1", "_xlfn.SORT([2]Data!A1:A2)", 1, f_attrs=' t="array" ref="F1:F2"', attrs=' cm="1"'),
    cell("F2", None, 2),
    cell("G1", None, 3),
]
REPORT_AFTER = (
    '<conditionalFormatting sqref="A1:A5"><cfRule type="expression" dxfId="0" priority="2">'
    "<formula>A1&gt;[2]Data!$A$1</formula></cfRule>"
    '<cfRule type="cellIs" dxfId="0" priority="1" operator="greaterThan"><formula>100</formula></cfRule></conditionalFormatting>'
    '<conditionalFormatting sqref="G1"><cfRule type="expression" dxfId="0" priority="3"><formula>[3]Sheet1!$A$1=1</formula>'
    "</cfRule></conditionalFormatting>"
    '<dataValidations count="2"><dataValidation type="list" allowBlank="1" sqref="H1:H5"><formula1>[2]Data!$A$1:$A$3</formula1>'
    '</dataValidation><dataValidation type="whole" sqref="H6"><formula1>1</formula1><formula2>10</formula2></dataValidation>'
    "</dataValidations>"
    '<hyperlinks><hyperlink ref="A10" r:id="rId1"/>'
    '<hyperlink ref="A11" location="\'Data Sheet\'!A1" display="Back"/></hyperlinks>'
)
REPORT_EXT = (
    f'<ext uri="{{78C0D931-6437-407d-A8EE-F0AAD7539E65}}" xmlns:x14="{X14}"><x14:conditionalFormattings>'
    f'<x14:conditionalFormatting xmlns:xm="{XM}"><x14:cfRule type="expression" priority="4" '
    'id="{11111111-2222-3333-4444-555555555555}"><xm:f>[2]Data!$B$1&gt;0</xm:f><x14:dxf><fill><patternFill>'
    '<bgColor rgb="FFFFC7CE"/></patternFill></fill></x14:dxf></x14:cfRule><xm:sqref>C1:C5</xm:sqref>'
    "</x14:conditionalFormatting></x14:conditionalFormattings></ext>"
    f'<ext uri="{{CCE6A557-97BC-4b89-ADB6-D9C93CAAB3DF}}" xmlns:x14="{X14}"><x14:dataValidations count="1" xmlns:xm="{XM}">'
    '<x14:dataValidation type="list" allowBlank="1"><x14:formula1><xm:f>[2]Data!$C$1:$C$3</xm:f></x14:formula1>'
    "<xm:sqref>I1:I5</xm:sqref></x14:dataValidation></x14:dataValidations></ext>"
    f'<ext uri="{{05C60535-1F16-4fd2-B633-F4F36F0B64E0}}" xmlns:x14="{X14}"><x14:sparklineGroups xmlns:xm="{XM}">'
    '<x14:sparklineGroup displayEmptyCellsAs="gap"><x14:colorSeries rgb="FF376092"/><x14:colorNegative rgb="FFD00000"/>'
    '<x14:colorAxis rgb="FF000000"/><x14:colorMarkers rgb="FFD00000"/><x14:colorFirst rgb="FFD00000"/>'
    '<x14:colorLast rgb="FFD00000"/><x14:colorHigh rgb="FFD00000"/><x14:colorLow rgb="FFD00000"/><x14:sparklines>'
    "<x14:sparkline><xm:f>[2]Data!A1:A3</xm:f><xm:sqref>J1</xm:sqref></x14:sparkline></x14:sparklines>"
    "</x14:sparklineGroup></x14:sparklineGroups></ext>"
)
NAMES = (
    '<definedName name="_xlnm.Print_Area" localSheetId="1">Report!$A$1:$D$10</definedName>'
    '<definedName name="Chain">TaxRate*2</definedName>'
    '<definedName name="Internal">Report!$A$1</definedName>'
    '<definedName name="LocalRate" localSheetId="1">[2]Data!$E$2</definedName>'
    '<definedName name="Secret" hidden="1">[3]Sheet1!$A$1</definedName>'
    '<definedName name="TaxRate">[2]Data!$E$1</definedName>'
)
CONNECTIONS = (
    f'<connections xmlns="{MAIN}"><connection id="1" name="Warehouse" type="5" refreshedVersion="6" background="1">'
    '<dbPr connection="Provider=SQLOLEDB.1;Password=example-only;User ID=sa;Data Source=db01;Initial Catalog=Sales" '
    'command="SELECT 1" commandType="2"/></connection><connection id="2" name="sales csv" type="6" refreshedVersion="6">'
    '<textPr sourceFile="C:\\Data\\sales.csv" delimited="1" comma="1"><textFields><textField/></textFields></textPr>'
    "</connection></connections>"
)


def mashup(code: str) -> bytes:
    """A Power Query DataMashup customXml part holding ``code`` as Section1.m."""
    inner = io.BytesIO()
    with zipfile.ZipFile(inner, "w") as zf:
        zf.writestr("Formulas/Section1.m", code)
    parts = inner.getvalue()
    blob = struct.pack("<II", 0, len(parts)) + parts + struct.pack("<IIII", 0, 0, 0, 0)
    return (
        f'<?xml version="1.0" encoding="utf-8"?><DataMashup xmlns="http://schemas.microsoft.com/DataMashup">'
        f"{base64.b64encode(blob).decode()}</DataMashup>"
    ).encode()


def build(
    path: Path,
    sheets: dict[str, str],
    links: tuple[str, ...] = (),
    names: str = "",
    sheet_rels: dict[str, str] | None = None,
    extra: dict[str, tuple[str, str, bytes]] | None = None,
) -> Path:
    """Write a workbook. ``extra`` maps part name -> (workbook rel type, content type, bytes)."""
    sheet_rels = sheet_rels or {}
    extra = extra or {}
    overrides = {"/xl/workbook.xml": f"{CT}.sheet.main+xml", "/xl/styles.xml": f"{CT}.styles+xml"}
    wb_rels = [("rIdS", "styles", "styles.xml", "")]
    parts: dict[str, str | bytes] = {"xl/styles.xml": STYLES}
    sheet_list = []
    for i, (name, xml) in enumerate(sheets.items(), start=1):
        parts[f"xl/worksheets/sheet{i}.xml"] = xml
        overrides[f"/xl/worksheets/sheet{i}.xml"] = f"{CT}.worksheet+xml"
        wb_rels.append((f"rId{i}", "worksheet", f"worksheets/sheet{i}.xml", ""))
        sheet_list.append(f'<sheet name="{escape(name)}" sheetId="{i}" r:id="rId{i}"/>')
        if name in sheet_rels:
            parts[f"xl/worksheets/_rels/sheet{i}.xml.rels"] = rels_xml([("rId1", "hyperlink", sheet_rels[name], "External")])
    refs = []
    for n, target in enumerate(links, start=1):
        parts[f"xl/externalLinks/externalLink{n}.xml"] = (
            f'<externalLink xmlns="{MAIN}"><externalBook xmlns:r="{R}" r:id="rId1"><sheetNames><sheetName val="Data"/>'
            '<sheetName val="Sheet1"/><sheetName val="My Sheet"/></sheetNames><sheetDataSet><sheetData sheetId="0">'
            '<row r="1"><cell r="A1"><v>10</v></cell></row></sheetData></sheetDataSet></externalBook></externalLink>'
        )
        parts[f"xl/externalLinks/_rels/externalLink{n}.xml.rels"] = rels_xml([("rId1", "externalLinkPath", target, "External")])
        overrides[f"/xl/externalLinks/externalLink{n}.xml"] = f"{CT}.externalLink+xml"
        wb_rels.append((f"rIdL{n}", "externalLink", f"externalLinks/externalLink{n}.xml", ""))
        refs.append(f'<externalReference r:id="rIdL{n}"/>')
    for part, (rel_type, content_type, data) in extra.items():
        parts[part] = data
        overrides[f"/{part}"] = content_type
        if rel_type:
            wb_rels.append((f"rIdX{len(wb_rels)}", rel_type, "../" + part if part.startswith("customXml") else part[3:], ""))
    references = f"<externalReferences>{' '.join(refs)}</externalReferences>" if refs else ""
    defined = f"<definedNames>{names}</definedNames>" if names else ""
    parts["xl/workbook.xml"] = (
        f'<workbook xmlns="{MAIN}" xmlns:r="{R}"><workbookPr/><bookViews><workbookView/></bookViews>'
        f'<sheets>{"".join(sheet_list)}</sheets>{references}{defined}<calcPr calcId="191029"/></workbook>'
    )
    parts["xl/_rels/workbook.xml.rels"] = rels_xml(wb_rels)
    parts["_rels/.rels"] = rels_xml([("rId1", "officeDocument", "xl/workbook.xml", "")])
    types = "".join(f'<Override PartName="{p}" ContentType="{t}"/>' for p, t in overrides.items())
    parts["[Content_Types].xml"] = (
        '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
        f'<Default Extension="xml" ContentType="application/xml"/>{types}</Types>'
    )
    order = ["[Content_Types].xml", "_rels/.rels"] + [p for p in parts if p not in ("[Content_Types].xml", "_rels/.rels")]
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as zf:
        for name in order:
            data = parts[name]
            zf.writestr(
                name, data if isinstance(data, bytes) else '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' + data
            )
    return path


def rels_xml(rels: list[tuple[str, str, str, str]]) -> str:
    items = "".join(
        f'<Relationship Id="{rid}" Type="{R}/{kind}" Target="{escape(target)}"' + (f' TargetMode="{mode}"' if mode else "") + "/>"
        for rid, kind, target, mode in rels
    )
    return f'<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">{items}</Relationships>'


def kitchen_sink(path: Path, with_mashup: bool = False) -> Path:
    """The main test workbook: links [1] (unused URL), [2] Source.xlsx, [3] missing/Old Budget.xlsx."""
    calc = f'<calcChain xmlns="{MAIN}"><c r="A1" i="2"/><c r="B2" i="2"/><c r="D1" i="2"/></calcChain>'
    extra = {
        "xl/calcChain.xml": ("calcChain", f"{CT}.calcChain+xml", calc.encode()),
        "xl/metadata.xml": ("sheetMetadata", f"{CT}.sheetMetadata+xml", METADATA.encode()),
        "xl/connections.xml": ("connections", f"{CT}.connections+xml", CONNECTIONS.encode()),
    }
    if with_mashup:
        code = (
            'section Section1;\r\nshared Sales = let\r\n    Source = Csv.Document(File.Contents("C:\\Data\\sales.csv"))\r\n'
            'in\r\n    Source;\r\nshared #"Web Rates" = Web.Contents("https://example.com/rates.json");'
        )
        extra["customXml/item1.xml"] = ("customXml", "application/xml", mashup(code))
    data = sheet_xml([cell("A1", None, 1), cell("A2", None, 2), cell("B1", "Report!A1+1", 11)])
    return build(
        path,
        {"Data Sheet": data, "Report": sheet_xml(REPORT_CELLS, REPORT_AFTER, REPORT_EXT)},
        LINKS,
        NAMES,
        {"Report": "https://example.com/budget-guide"},
        extra,
    )


def plain(path: Path) -> Path:
    """A workbook without any external link."""
    return build(path, {"Sheet1": sheet_xml([cell("A1", None, 1), cell("A2", "A1*2", 2)])})
