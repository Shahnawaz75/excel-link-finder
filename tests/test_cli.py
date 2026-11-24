import csv
import json
import re
import subprocess
import sys
from pathlib import Path

import builder
import pytest
from conftest import OLE2

from excel_link_finder import __version__, scan
from excel_link_finder.cli import main
from excel_link_finder.report import CSV_FIELDS


def test_scan_exit_codes_and_table(kitchen, plain, capsys):
    assert main(["scan", str(plain)]) == 0
    assert "No external links found." in capsys.readouterr().out
    assert main(["scan", str(kitchen)]) == 1
    out = capsys.readouterr().out
    assert "Found: external-link 3 (1 broken), defined-name 4, cell-formula 8" in out
    assert "missing/Old Budget.xlsx (MISSING)" in out and "Also found (not data links): connection 2, hyperlink 1" in out


def test_scan_xls_explains_what_to_do(tmp_path, capsys):
    old = tmp_path / "old.xls"
    old.write_bytes(OLE2 + bytes(504))
    assert main(["scan", str(old)]) == 2
    assert "save it as .xlsx" in capsys.readouterr().out


def test_json_report(kitchen, capsys):
    assert main(["scan", "--json", str(kitchen)]) == 1
    doc = json.loads(capsys.readouterr().out)
    assert (doc["tool"], doc["version"], doc["totals"]) == (
        "excel-link-finder",
        __version__,
        {"files": 1, "with_links": 1, "errors": 0},
    )
    (entry,) = doc["files"]
    assert set(entry) == {"file", "error", "has_links", "summary", "findings"}
    assert all(
        set(f) == {"type", "sheet", "object", "ref", "text", "links", "target", "status", "note"} for f in entry["findings"]
    )
    assert entry["summary"]["cell-formula"] == 8 and entry["summary"]["broken"] == 1


def test_csv_report_file(kitchen, tmp_path, capsys):
    report = tmp_path / "links.csv"
    assert main(["scan", "--csv", "-o", str(report), str(kitchen)]) == 1
    rows = list(csv.DictReader(report.read_text(encoding="utf-8-sig").splitlines()))
    assert tuple(rows[0]) == CSV_FIELDS
    assert {"type": "cell-formula", "sheet": "Report", "ref": "A1", "links": "2"}.items() <= next(
        r for r in rows if r["ref"] == "A1"
    ).items()
    assert len(rows) == len(scan(kitchen).findings)


def test_folder_scan_recursive(tmp_path, capsys):
    (tmp_path / "deep").mkdir()
    builder.plain(tmp_path / "plain.xlsx")
    builder.kitchen_sink(tmp_path / "deep" / "linked.xlsx")
    assert main(["scan", str(tmp_path)]) == 0
    assert main(["scan", "--recursive", "--quiet", str(tmp_path)]) == 1
    out = capsys.readouterr().out
    assert "Scanned 2 workbooks: 1 with external links, 1 without, 0 could not be read." in out
    assert "TYPE" not in out  # --quiet prints the summaries only


def test_strict_mode_counts_hyperlinks(tmp_path):
    sheet = builder.sheet_xml([builder.cell("A1", None, 1)], '<hyperlinks><hyperlink ref="A1" r:id="rId1"/></hyperlinks>')
    book = builder.build(tmp_path / "web.xlsx", {"S": sheet}, sheet_rels={"S": "https://example.com/"})
    assert main(["scan", str(book)]) == 0
    assert main(["scan", "--strict", str(book)]) == 1


def test_clean_command_default_output(kitchen, capsys):
    assert main(["clean", str(kitchen)]) == 0
    assert (kitchen.parent / "kitchen_clean.xlsx").exists()
    assert "Check: the cleaned copy has no external links." in capsys.readouterr().out


def test_clean_command_reports_what_is_left(example, tmp_path, capsys):
    assert main(["clean", str(example), "-o", str(tmp_path / "c.xlsx")]) == 1  # the pivot cache stays
    out = capsys.readouterr().out
    assert "PivotTable Analyze > Change Data Source" in out and re.search(r"external-link\s+2", out)
    assert main(["clean", str(example), "-o", str(tmp_path / "j.xlsx"), "--json"]) == 1
    assert json.loads(capsys.readouterr().out)["remaining_links"] == 1


def test_clean_command_refuses_to_overwrite_the_input(kitchen, capsys):
    assert main(["clean", str(kitchen), "-o", str(kitchen)]) == 2
    assert "refusing to overwrite the input" in capsys.readouterr().err


def test_version_and_module_entry_point(capsys):
    with pytest.raises(SystemExit) as exc:
        main(["--version"])
    assert exc.value.code == 0 and __version__ in capsys.readouterr().out
    run = subprocess.run([sys.executable, "-m", "excel_link_finder", "--version"], capture_output=True, text=True, check=True)
    assert __version__ in run.stdout
    pyproject = (Path(__file__).resolve().parent.parent / "pyproject.toml").read_text(encoding="utf-8")
    assert f'version = "{__version__}"' in pyproject
