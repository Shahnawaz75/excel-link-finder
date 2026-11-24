import warnings
import zipfile

import pytest

from excel_link_finder import ExcelLinkFinderError, check_package, clean, scan


def part(path, name):
    with zipfile.ZipFile(path) as zf:
        return zf.read(name).decode("utf-8")


def names(path):
    with zipfile.ZipFile(path) as zf:
        return zf.namelist()


def test_clean_removes_link_parts_rels_and_content_types(kitchen, tmp_path):
    result = clean(kitchen, tmp_path / "out.xlsx")
    out = result.output
    assert result.is_clean
    assert result.removed_links == [
        "https://contoso.sharepoint.com/sites/fin/Rates.xlsx",
        "Source.xlsx",
        "missing/Old Budget.xlsx",
    ]
    assert not [n for n in names(out) if "externalLink" in n]
    assert "externalReferences" not in part(out, "xl/workbook.xml")
    assert "externalLink" not in part(out, "[Content_Types].xml") + part(out, "xl/_rels/workbook.xml.rels")
    assert check_package(out) == []
    assert {f.type for f in scan(out).findings} == {"connection"}  # connections are not links and are kept


def test_formulas_are_replaced_by_their_cached_values(kitchen, tmp_path):
    sheet = part(clean(kitchen, tmp_path / "out.xlsx").output, "xl/worksheets/sheet2.xml")
    assert '<c r="A1"><v>10</v></c>' in sheet
    assert '<c r="C1" t="inlineStr"><is><t xml:space="preserve">hello &amp; bye</t></is></c>' in sheet
    for ref, value in (("B2", 2), ("B3", 3), ("B4", 4)):  # every cell of the shared formula
        assert f'<c r="{ref}"><v>{value}</v></c>' in sheet
    assert '<c r="E1"><v>20</v></c>' in sheet and '<c r="F1"><v>1</v></c>' in sheet  # array + dynamic array (cm dropped)
    assert "<f>SUM(A1:A3)</f>" in sheet  # formulas inside the workbook stay


def test_calc_chain_is_dropped_when_formulas_change(kitchen, tmp_path):
    out = clean(kitchen, tmp_path / "out.xlsx").output
    assert "xl/calcChain.xml" not in names(out)
    assert "calcChain" not in part(out, "xl/_rels/workbook.xml.rels") + part(out, "[Content_Types].xml")


def test_rules_and_names_are_removed_internal_ones_kept(kitchen, tmp_path):
    out = clean(kitchen, tmp_path / "out.xlsx").output
    sheet = part(out, "xl/worksheets/sheet2.xml")
    assert sheet.count("<conditionalFormatting") == 1 and "<formula>100</formula>" in sheet
    assert '<dataValidations count="1">' in sheet and 'sqref="H6"' in sheet
    assert "<extLst>" not in sheet  # the x14 rule, list and sparkline were all external
    assert '<hyperlink ref="A11"' in sheet and 'ref="A10"' not in sheet
    assert "budget-guide" not in part(out, "xl/worksheets/_rels/sheet2.xml.rels")
    workbook = part(out, "xl/workbook.xml")
    assert 'name="Internal"' in workbook and "_xlnm.Print_Area" in workbook
    assert "TaxRate" not in workbook and "Secret" not in workbook


def test_keep_hyperlinks(kitchen, tmp_path):
    out = clean(kitchen, tmp_path / "out.xlsx", keep_hyperlinks=True).output
    assert [f.target for f in scan(out).findings if f.type == "hyperlink"] == ["https://example.com/budget-guide"]


def test_partial_clean_renumbers_the_links_that_stay(kitchen, tmp_path):
    result = clean(kitchen, tmp_path / "out.xlsx", only=["cell-formula"])
    assert result.removed_links == ["https://contoso.sharepoint.com/sites/fin/Rates.xlsx"]  # [1] was never used
    assert not result.is_clean
    after = scan(result.output)
    assert {f.object: f.text for f in after.findings if f.type == "defined-name"}["Secret"] == "[2]Sheet1!$A$1"  # was [3]
    assert {f.text for f in after.findings if f.type in ("sparkline", "data-validation")} == {
        "[1]Data!A1:A3",
        "[1]Data!$A$1:$A$3",
        "[1]Data!$C$1:$C$3",
    }  # were [2]
    assert [f.target for f in after.findings if f.type == "external-link"] == ["Source.xlsx", "missing/Old Budget.xlsx"]
    assert not [f for f in after.findings if f.type == "cell-formula"]
    assert check_package(result.output) == []


def test_cleaned_copy_opens_in_openpyxl(kitchen, tmp_path):
    openpyxl = pytest.importorskip("openpyxl")
    out = clean(kitchen, tmp_path / "out.xlsx").output
    with warnings.catch_warnings():
        warnings.simplefilter("ignore")
        sheet = openpyxl.load_workbook(out)["Report"]
    assert [sheet[c].value for c in ("A1", "B3", "C1", "E1", "C3")] == [10, 3, "hello & bye", 20, 7]
    assert sheet["D1"].value == "=SUM(A1:A3)"


def test_clean_the_excel_made_example(example, tmp_path):
    result = clean(example, tmp_path / "clean.xlsx")
    assert result.removed == {
        "defined-name": 3,
        "cell-formula": 15,
        "hyperlink": 1,
        "conditional-format": 1,
        "data-validation": 1,
        "sparkline": 1,
        "object-link": 2,
        "shape-macro": 2,
        "chart": 4,
        "table-formula": 1,
        "external-link": 2,
    }
    assert [f.type for f in result.remaining.links] == ["pivot-cache"]  # has to be changed in Excel
    assert check_package(result.output) == []
    chart = part(result.output, "xl/charts/chart1.xml")
    assert "<c:f>" not in chart and "<c:numLit>" in chart and "<c:strLit>" in chart and "<c:rich" in chart
    assert "FmlaMacro" not in part(result.output, "xl/drawings/vmlDrawing1.vml")
    assert "[2]!" not in part(result.output, "xl/drawings/drawing1.xml")


def test_the_original_is_never_modified(kitchen, tmp_path):
    before = kitchen.read_bytes()
    clean(kitchen, tmp_path / "out.xlsx")
    assert kitchen.read_bytes() == before
    with pytest.raises(ExcelLinkFinderError, match="input"):
        clean(kitchen, kitchen)
    with pytest.raises(ExcelLinkFinderError, match="already exists"):
        clean(kitchen, tmp_path / "out.xlsx")
    assert clean(kitchen, tmp_path / "out.xlsx", overwrite=True).is_clean
    with pytest.raises(ExcelLinkFinderError, match="extension"):
        clean(kitchen, tmp_path / "out.xlsm")
    with pytest.raises(ExcelLinkFinderError, match="unknown type"):
        clean(kitchen, tmp_path / "x.xlsx", only=["pivot-cache"])


def test_workbook_without_links_is_copied_unchanged(plain, tmp_path):
    result = clean(plain, tmp_path / "copy.xlsx")
    assert result.removed == {} and result.is_clean
    assert (tmp_path / "copy.xlsx").read_bytes() == plain.read_bytes()


def test_check_package_finds_broken_structure(kitchen, tmp_path):
    broken = tmp_path / "broken.xlsx"
    with zipfile.ZipFile(kitchen) as zin, zipfile.ZipFile(broken, "w") as zout:
        for info in zin.infolist():
            if info.filename != "xl/externalLinks/externalLink2.xml":
                zout.writestr(info, zin.read(info))
    problems = check_package(broken)
    assert any("points to missing part xl/externalLinks/externalLink2.xml" in p for p in problems)
    assert any("override for missing part" in p for p in problems)
