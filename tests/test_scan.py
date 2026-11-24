import os

import builder
import pytest
from conftest import OLE2

from excel_link_finder import ExcelLinkFinderError, UnsupportedFileError, find_workbooks, scan
from excel_link_finder.scanner import link_status


def of_type(result, kind):
    return [f for f in result.findings if f.type == kind]


def test_external_link_parts_targets_and_status(kitchen):
    links = of_type(scan(kitchen), "external-link")
    assert [(f.object, f.text, f.status) for f in links] == [
        ("[1]", "externalLink1.xml", "url"),
        ("[2]", "externalLink2.xml", "found"),
        ("[3]", "externalLink3.xml", "missing"),
    ]
    assert links[2].target == "missing/Old Budget.xlsx"
    assert scan(kitchen, check_paths=False).findings[0].status == "not-checked"


def test_cell_formulas_with_sheet_and_cell(kitchen):
    cells = {f.ref: f for f in of_type(scan(kitchen), "cell-formula")}
    assert set(cells) == {"A1", "B2:B4", "C1", "C2", "C3", "D4", "E1", "F1"}
    assert (cells["A1"].sheet, cells["A1"].text, cells["A1"].links, cells["A1"].target) == (
        "Report",
        "[2]Data!A1",
        [2],
        "Source.xlsx",
    )
    assert cells["C2"].links == [3]
    assert cells["C3"].links == [] and cells["C3"].target == "C:\\Finance\\Budget 2025.xlsx"
    assert cells["B2:B4"].note == "shared formula, 3 cells, via name TaxRate"
    assert cells["E1"].note == "array formula E1:E2"


def test_internal_formulas_are_not_reported(kitchen):
    texts = {f.text for f in scan(kitchen).findings}
    assert not texts & {"SUM(A1:A3)", '"[2]Data!A1"&"x"', "Report!A1+1", "Report!$A$1"}


def test_defined_names_hidden_local_and_chained(kitchen):
    names = {f.object: f for f in of_type(scan(kitchen), "defined-name")}
    assert set(names) == {"TaxRate", "Secret", "LocalRate", "Chain"}
    assert (names["Secret"].note, names["Secret"].links) == ("hidden", [3])
    assert names["LocalRate"].sheet == "Report"
    assert (names["Chain"].note, names["Chain"].links) == ("via name TaxRate", [2])


def test_conditional_formatting_validation_and_sparklines(kitchen):
    result = scan(kitchen)
    rules = {f.ref: f.text for f in of_type(result, "conditional-format")}
    assert rules == {"A1:A5": "A1>[2]Data!$A$1", "G1": "[3]Sheet1!$A$1=1", "C1:C5": "[2]Data!$B$1>0"}  # not the "100" rule
    lists = {f.ref: f.text for f in of_type(result, "data-validation")}
    assert lists == {"H1:H5": "[2]Data!$A$1:$A$3", "I1:I5": "[2]Data!$C$1:$C$3"}
    assert [(f.ref, f.text) for f in of_type(result, "sparkline")] == [("J1", "[2]Data!A1:A3")]


def test_hyperlinks_and_connections_are_reported_separately(kitchen):
    result = scan(kitchen)
    (link,) = of_type(result, "hyperlink")
    assert (link.sheet, link.ref, link.target) == ("Report", "A10", "https://example.com/budget-guide")
    connections = {f.object: f for f in of_type(result, "connection")}
    assert "Password=***" in connections["Warehouse"].text and "example-only" not in connections["Warehouse"].text
    assert connections["sales csv"].target == "C:\\Data\\sales.csv"
    assert not any(f.is_data_link for f in [link, *connections.values()])
    assert result.counts() == {
        "external-link": 3,
        "defined-name": 4,
        "cell-formula": 8,
        "conditional-format": 3,
        "data-validation": 2,
        "sparkline": 1,
        "connection": 2,
        "hyperlink": 1,
    }


def test_power_query_sources(tmp_path):
    result = scan(builder.kitchen_sink(tmp_path / "pq.xlsx", with_mashup=True))
    queries = {f.object: f.target for f in of_type(result, "connection") if f.note == "Power Query"}
    assert queries == {"Sales": "C:\\Data\\sales.csv", "Web Rates": "https://example.com/rates.json"}


def test_excel_made_example(example):
    result = scan(example)
    assert result.counts() == {
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
    }
    links = [(f.target, f.status) for f in of_type(result, "external-link")]
    assert links == [("old-budget.xlsx", "missing"), ("prices.xlsm", "found")]
    assert {f.note for f in of_type(result, "chart")} == {
        "chart title",
        "series 1 name",
        "series 1 categories",
        "series 1 values",
    }
    assert {(f.sheet, f.object, f.text) for f in of_type(result, "shape-macro")} == {
        ("Report", "Refresh shape", "[2]!RefreshPrices"),
        ("Report", "Button 2", "[2]!UpdateReport"),
    }
    assert {f.object for f in of_type(result, "object-link")} == {"VAT box", "Drop Down 1"}
    (pivot,) = of_type(result, "pivot-cache")
    assert (pivot.object, pivot.target, pivot.status) == ("StockPivot", "prices.xlsm", "found")
    assert [f.object for f in of_type(result, "table-formula")] == ["Orders[Net]"]
    assert [f.object for f in of_type(result, "defined-name") if f.note == "hidden"] == ["OldTotal"]
    assert [f.target for f in of_type(result, "data-validation")] == ["C:\\work\\examples\\prices.xlsm"]


def test_link_status_checks_local_paths_only(tmp_path):
    (tmp_path / "here.xlsx").write_bytes(b"x")
    assert link_status("here.xlsx", tmp_path) == "found"
    assert link_status("not-here.xlsx", tmp_path) == "missing"
    assert link_status("https://example.com/book.xlsx", tmp_path) == "url"
    expected = "missing" if os.name == "nt" else "unknown"
    assert link_status("file:///Z:\\no\\such\\folder\\book.xlsx", tmp_path) == expected


@pytest.mark.parametrize(
    ("name", "data", "message"),
    [
        ("old.xls", OLE2 + bytes(504), "save it as .xlsx"),
        ("locked.xlsx", OLE2 + bytes(504), "encrypted"),
        ("binary.xlsb", b"PK\x03\x04" + bytes(30), "binary workbook"),
        ("notes.xlsx", b"just text", "not an Excel workbook"),
    ],
)
def test_unsupported_files_get_a_clear_message(tmp_path, name, data, message):
    path = tmp_path / name
    path.write_bytes(data)
    with pytest.raises(UnsupportedFileError, match=message):
        scan(path)


def test_damaged_zip_is_an_error(tmp_path):
    path = tmp_path / "damaged.xlsx"
    path.write_bytes(b"PK\x03\x04" + bytes(100))
    with pytest.raises(ExcelLinkFinderError):
        scan(path)


def test_find_workbooks_in_folders(tmp_path):
    (tmp_path / "sub").mkdir()
    builder.plain(tmp_path / "a.xlsx")
    builder.plain(tmp_path / "sub" / "b.xlsm")
    (tmp_path / "~$a.xlsx").write_bytes(b"Office lock file")
    (tmp_path / "old.xls").write_bytes(OLE2)
    (tmp_path / "notes.txt").write_text("not a workbook")
    found, skipped = find_workbooks([tmp_path])
    assert [p.name for p in found] == ["a.xlsx"] and [p.name for p in skipped] == ["old.xls"]
    found, _ = find_workbooks([tmp_path], recursive=True)
    assert sorted(p.name for p in found) == ["a.xlsx", "b.xlsm"]
