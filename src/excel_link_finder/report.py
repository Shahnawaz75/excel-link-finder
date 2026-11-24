"""Render scan and clean results as a text table, JSON or CSV."""

from __future__ import annotations

import csv
import io
import json
import textwrap

from .cleaner import CleanResult
from .scanner import DATA_LINK_TYPES, INFO_TYPES, Finding, ScanResult

CSV_FIELDS = ("file", "type", "sheet", "object", "ref", "text", "links", "target", "status", "note")
_COLUMNS = (("TYPE", 18), ("SHEET / NAME / OBJECT", 36), ("CELL/RANGE", 11), ("FORMULA / TEXT", 36), ("LINKED FILE", 44))
_STATUS = {
    "found": "found",
    "missing": "MISSING",
    "url": "URL, not checked",
    "unknown": "not checked",
    "not-checked": "not checked",
}
_REMOVED = {
    "cell-formula": "formula cells now hold their last calculated value",
    "defined-name": "names deleted",
    "table-formula": "table column formulas deleted (the cells keep their values)",
    "conditional-format": "conditional formatting rules deleted",
    "data-validation": "data validation rules deleted",
    "sparkline": "sparklines deleted",
    "chart": "chart references replaced by the data the chart already showed",
    "shape-macro": "macro assignments cleared",
    "object-link": "shape and form control links cleared",
    "hyperlink": "external hyperlinks deleted",
    "external-link": "external link parts deleted",
}
_ADVICE = {
    "pivot-cache": "change it in Excel: PivotTable Analyze > Change Data Source",
    "connection": "not a link: manage it in Data > Queries & Connections",
    "object-link": "linked OLE object: delete or re-insert it in Excel",
    "chart": "edit this chart reference in Excel",
    "external-link": "still used by the items listed here",
}


def _cut(text: object, width: int) -> str:
    flat = " ".join(str(text).split())
    return flat if len(flat) <= width else flat[: width - 3] + "..."


def _basename(target: str) -> str:
    return target.replace("\\", "/").rstrip("/").rsplit("/", 1)[-1] or target


def where(finding: Finding) -> str:
    """Sheet / name / object column, with the note in brackets."""
    place = " > ".join(x for x in (finding.sheet, finding.object) if x)
    if finding.note and place:
        return f"{place} ({finding.note})"
    return place or finding.note


def _linked(finding: Finding, labels: dict[int, str]) -> str:
    if finding.type in ("external-link", "pivot-cache") and finding.status:
        return f"{finding.target} ({_STATUS.get(finding.status, finding.status)})"
    if finding.links:
        return ", ".join(f"[{i}] {labels.get(i, '?')}" for i in finding.links)
    return finding.target


def summary(result: ScanResult) -> list[str]:
    """One or two summary lines for a workbook."""
    counts = result.counts()
    lines = []
    if result.has_links:
        parts = []
        for name in DATA_LINK_TYPES:
            if name in counts:
                broken = sum(f.type == name for f in result.broken)
                parts.append(f"{name} {counts[name]}" + (f" ({broken} broken)" if broken else ""))
        lines.append("Found: " + ", ".join(parts))
    else:
        lines.append("No external links found.")
    info = [f"{name} {counts[name]}" for name in INFO_TYPES if name in counts]
    if info:
        lines.append("Also found (not data links): " + ", ".join(info))
    return lines


def render_text(results: list[ScanResult], quiet: bool = False) -> str:
    """Human-readable report: one table per workbook plus summary lines."""
    out: list[str] = []
    for result in results:
        out.append(str(result.path))
        if result.error:
            out += [f"  error: {result.error}", ""]
            continue
        if result.findings and not quiet:
            labels = {f.links[0]: _basename(f.target) for f in result.findings if f.type == "external-link" and f.links}
            rows = [[f.type, where(f), f.ref, f.text, _linked(f, labels)] for f in result.findings]
            cells = [[_cut(value or "-", width) for value, (_, width) in zip(row, _COLUMNS, strict=True)] for row in rows]
            widths = [max(len(head), *(len(r[i]) for r in cells)) for i, (head, _) in enumerate(_COLUMNS)]
            for row in [[head for head, _ in _COLUMNS], *cells]:
                out.append("  " + "  ".join(v.ljust(w) for v, w in zip(row, widths, strict=True)).rstrip())
            out.append("")
        for line in summary(result):
            out += textwrap.wrap(line, 110, initial_indent="  ", subsequent_indent="    ", break_on_hyphens=False)
        out.append("")
    if len(results) > 1:
        linked = sum(r.has_links for r in results)
        errors = sum(r.error is not None for r in results)
        out.append(
            f"Scanned {len(results)} workbooks: {linked} with external links, "
            f"{len(results) - linked - errors} without, {errors} could not be read."
        )
    return "\n".join(out).rstrip() + "\n"


def render_json(results: list[ScanResult], version: str) -> str:
    totals = {
        "files": len(results),
        "with_links": sum(r.has_links for r in results),
        "errors": sum(r.error is not None for r in results),
    }
    doc = {"tool": "excel-link-finder", "version": version, "files": [r.to_dict() for r in results], "totals": totals}
    return json.dumps(doc, indent=2) + "\n"


def render_csv(results: list[ScanResult]) -> str:
    buffer = io.StringIO()
    writer = csv.writer(buffer, lineterminator="\n")
    writer.writerow(CSV_FIELDS)
    for result in results:
        if result.error:
            writer.writerow([str(result.path), "error", "", "", "", result.error, "", "", "", ""])
        for finding in result.findings:
            row = finding.to_dict()
            row["links"] = " ".join(map(str, finding.links))
            writer.writerow([str(result.path), *(row[k] for k in CSV_FIELDS[1:])])
    return buffer.getvalue()


def render_clean(result: CleanResult) -> str:
    """Summary printed by ``excel-link-finder clean``."""
    out = [f"Cleaned copy written to {result.output}"]
    if result.removed:
        out.append("Removed:")
        for name, count in result.removed.items():
            extra = f" ({', '.join(_basename(t) for t in result.removed_links)})" if name == "external-link" else ""
            out.append(f"  {name:<19}{count:>4}  {_REMOVED.get(name, '')}{extra}")
        if result.calc_chain_removed:
            out.append(f"  {'calcChain.xml':<23}  removed, Excel rebuilds it on the next save")
    else:
        out.append("Nothing needed to be removed: the copy is identical to the original.")
    if result.empty_cells:
        out.append(f"  Note: {result.empty_cells} formula cell(s) had no saved value and are now empty.")
    if result.kept:
        out.append("Not removed:")
        for finding in result.kept:
            if finding.type == "external-link" or not finding.cleanable:
                advice = _ADVICE.get(finding.type, "fix it in Excel")
            else:
                advice = "not selected with --only"
            out.append(f"  {finding.type:<19}{_cut(where(finding) or finding.text, 44):<46}{advice}")
    remaining = result.remaining
    if remaining is not None:
        left = len(remaining.links)
        out.append(
            "Check: the cleaned copy has no external links."
            if not left
            else f"Check: the cleaned copy still has {left} data link location(s); "
            "run 'excel-link-finder scan' on it for details."
        )
    return "\n".join(out) + "\n"
