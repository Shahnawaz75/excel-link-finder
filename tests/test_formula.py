import pytest

from excel_link_finder.formula import external_refs, name_refs, renumber


def refs(formula: str) -> list[tuple[int | None, str]]:
    return [(r.index, r.target) for r in external_refs(formula)]


def test_numbered_references():
    assert refs("[1]Prices!A2*2") == [(1, "")]
    assert refs("SUM([2]Data!C2:C6)+'[3]Bob''s Data'!A1") == [(2, ""), (3, "")]
    assert refs("[1]!RefreshPrices") == [(1, "")]


def test_full_path_and_book_name_references():
    assert refs("'C:\\Finance\\[Budget 2025.xlsx]Plan'!$B$3") == [(None, "C:\\Finance\\Budget 2025.xlsx")]
    assert refs("[Budget.xlsx]Plan!A1") == [(None, "Budget.xlsx")]
    assert refs("Budget.xlsx!TaxRate") == [(None, "Budget.xlsx")]
    assert refs("'C:\\Data\\Book.xlsm'!Total") == [(None, "C:\\Data\\Book.xlsm")]
    assert refs("B2>[prices.xlsm]Prices!#REF!") == [(None, "prices.xlsm")]
    assert refs("'https://contoso.sharepoint.com/sites/x/[Rates.xlsx]Rates'!A1") == [
        (None, "https://contoso.sharepoint.com/sites/x/Rates.xlsx")
    ]


@pytest.mark.parametrize(
    "formula",
    [
        '"[1]Data!A1"&"x"',  # text, not a reference
        "Orders[[#This Row],[Qty]]*2",  # table (structured) reference
        "Table1[1]",  # table column called "1"
        "[0]!LocalName",  # [0] is the workbook itself
        "SUM(Sheet1!A1:B2)+'My Sheet'!A1",
        "#REF!+1",
        'IFERROR(A1,"[2]")',
    ],
)
def test_things_that_are_not_external(formula):
    assert refs(formula) == []


def test_table_reference_next_to_an_external_reference():
    assert refs("Orders[[#This Row],[Qty]]*[2]Prices!$B$2") == [(2, "")]


def test_renumber_only_touches_link_numbers():
    assert renumber("[2]A!A1+'[3]B C'!A1+[1]D!A1", {1: 1, 2: 1, 3: 2}) == "[1]A!A1+'[2]B C'!A1+[1]D!A1"
    assert renumber('"[2]"&[2]A!A1+Orders[[#All],[2]]', {2: 1}) == '"[2]"&[1]A!A1+Orders[[#All],[2]]'


def test_name_refs():
    formula = "TaxRate*A1+SUM(Rates)+Report!LocalRate+[1]!Ext+[1]Data!Ext2+'My Sheet'!Other+TRUE+$B$2"
    assert name_refs(formula) == [(None, "taxrate"), (None, "rates"), ("report", "localrate"), ("my sheet", "other")]
