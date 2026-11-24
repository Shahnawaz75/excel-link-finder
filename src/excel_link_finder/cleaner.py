"""Write a cleaned copy of a workbook with its external links removed (the original is never changed)."""

from __future__ import annotations

import os
import shutil
import zipfile
from collections import defaultdict
from collections.abc import Iterable
from dataclasses import dataclass, field
from html import unescape
from pathlib import Path

from .formula import renumber
from .ooxml import Editor, Node, Package, _rel_attr_values, check_package, escape_text, inner_xml, parse_xml, prefix_of, rels_name
from .scanner import ExcelLinkFinderError, Finding, ScanResult, _Scanner, check_format, scan

CLEANABLE_TYPES = (
    "cell-formula",
    "defined-name",
    "table-formula",
    "conditional-format",
    "data-validation",
    "sparkline",
    "chart",
    "shape-macro",
    "object-link",
    "hyperlink",
)

# A container is removed with its last item, because Excel rejects some empty containers.
_CASCADE: dict[str, tuple[str, ...] | None] = {
    "conditionalFormatting": ("cfRule",),
    "conditionalFormattings": ("conditionalFormatting",),
    "dataValidations": ("dataValidation",),
    "definedNames": ("definedName",),
    "hyperlinks": ("hyperlink",),
    "sparklines": ("sparkline",),
    "sparklineGroup": ("sparklines",),
    "sparklineGroups": ("sparklineGroup",),
    "externalReferences": ("externalReference",),
    "extLst": ("ext",),
    "ext": None,  # any child
}
_DRAWINGML = {
    "http://schemas.openxmlformats.org/drawingml/2006/chart": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "http://purl.oclc.org/ooxml/drawingml/chart": "http://purl.oclc.org/ooxml/drawingml/main",
}


@dataclass
class CleanResult:
    """What :func:`clean` changed."""

    source: Path
    output: Path
    removed: dict[str, int] = field(default_factory=dict)
    removed_links: list[str] = field(default_factory=list)
    kept: list[Finding] = field(default_factory=list)
    remaining: ScanResult | None = None
    empty_cells: int = 0
    calc_chain_removed: bool = False

    @property
    def is_clean(self) -> bool:
        """True when the cleaned copy has no data links left."""
        return self.remaining is not None and not self.remaining.has_links

    def to_dict(self) -> dict[str, object]:
        return {
            "source": str(self.source),
            "output": str(self.output),
            "removed": self.removed,
            "removed_links": self.removed_links,
            "kept": [f.to_dict() for f in self.kept],
            "empty_cells": self.empty_cells,
            "calc_chain_removed": self.calc_chain_removed,
            "remaining_links": len(self.remaining.links) if self.remaining else None,
        }


def clean(
    source: str | Path,
    output: str | Path | None = None,
    *,
    keep_hyperlinks: bool = False,
    only: Iterable[str] | None = None,
    overwrite: bool = False,
) -> CleanResult:
    """Write a copy of ``source`` without its external links and return what was removed.

    * external cell formulas are replaced by their cached values,
    * defined names, conditional formats, data validations, sparklines and table formulas
      that point to other workbooks are deleted,
    * charts keep their cached data but lose the link,
    * macro assignments and control links that point to other workbooks are cleared,
    * external hyperlinks are removed unless ``keep_hyperlinks`` is set,
    * external link parts that nothing uses any more are removed, with their relationships
      and content types, and ``calcChain.xml`` is dropped when formulas change.

    ``only`` limits the cleaning to some finding types (see :data:`CLEANABLE_TYPES`).
    Pivot caches, connections and linked OLE objects are reported in ``kept``.
    """
    src, dst = Path(source), Path(output) if output else Path(source).with_name(f"{Path(source).stem}_clean{Path(source).suffix}")
    check_format(src)
    if dst.suffix.lower() != src.suffix.lower():
        raise ExcelLinkFinderError(f"the output file must have the same extension as the input ({src.suffix})")
    if dst.exists():
        if os.path.samefile(dst, src):
            raise ExcelLinkFinderError("refusing to overwrite the input file: choose another output name")
        if not overwrite:
            raise ExcelLinkFinderError(f"{dst} already exists (use --force to overwrite it)")
    selected = set(only) if only else set(CLEANABLE_TYPES)
    if unknown := selected - set(CLEANABLE_TYPES):
        raise ExcelLinkFinderError(f"unknown type(s) for --only: {', '.join(sorted(unknown))}")
    if keep_hyperlinks:
        selected.discard("hyperlink")
    result = CleanResult(src, dst)
    with Package(src) as pkg:
        scanner = _Scanner(pkg, src, check_paths=False)
        cleaner = _Cleaner(pkg, scanner.run(), scanner.wb_part, selected, result)
        parts, deleted = cleaner.run()
        if parts or deleted:
            _write_zip(pkg, dst, parts, deleted)
        else:
            shutil.copyfile(src, dst)
    problems = check_package(dst)
    if problems:  # only count problems the original did not have already
        problems = sorted(set(problems) - set(check_package(src)))
    if problems:
        dst.unlink()
        raise ExcelLinkFinderError(
            "internal error, the cleaned copy failed the package check: " + "; ".join(sorted(problems)[:3])
        )
    result.remaining = scan(dst, check_paths=False)
    return result


def _remove_nodes(editor: Editor, nodes: list[Node]) -> None:
    """Remove ``nodes``, plus containers that would be left empty; fix ``count`` attributes."""
    gone = {id(n): n for n in nodes}
    queue = list(nodes)
    while queue:
        parent = queue.pop().parent
        if parent is None or id(parent) in gone or parent.local not in _CASCADE:
            continue
        items = _CASCADE[parent.local]
        members = [c for c in parent.children if items is None or c.local in items]
        if members and all(id(c) in gone for c in members):
            gone[id(parent)] = parent
            queue.append(parent)
    for node in gone.values():
        editor.remove(node)
    parents = {id(n.parent): n.parent for n in gone.values() if n.parent is not None and id(n.parent) not in gone}
    for parent in parents.values():
        if parent is not None and parent.get("count") is not None and parent.local in _CASCADE:
            items = _CASCADE[parent.local]
            left = [c for c in parent.children if (items is None or c.local in items) and id(c) not in gone]
            editor.set_attr(parent, "count", str(len(left)))


class _Cleaner:
    """Turns the findings of one package into byte-level edits."""

    def __init__(self, pkg: Package, findings: list[Finding], wb_part: str, selected: set[str], result: CleanResult) -> None:
        self.pkg = pkg
        self.findings = findings
        self.wb_part = wb_part
        self.selected = selected
        self.result = result
        self.editors: dict[str, Editor] = {}
        self.removals: dict[str, list[Node]] = defaultdict(list)
        self.rel_drops: dict[str, set[str]] = defaultdict(set)
        self.deleted: set[str] = set()
        self.cells_changed = False

    def editor(self, part: str) -> Editor:
        if part not in self.editors:
            self.editors[part] = Editor(self.pkg.read(part))
        return self.editors[part]

    def run(self) -> tuple[dict[str, bytes], set[str]]:
        links = [f for f in self.findings if f.type == "external-link"]
        remove = [f for f in self.findings if f.type in self.selected and f.cleanable]
        removing = {id(f) for f in remove}
        kept = [f for f in self.findings if f.type != "external-link" and id(f) not in removing]
        used = {i for f in kept for i in f.links}
        drop = [f for f in links if f.links[0] not in used and f.cleanable and f.edits[0].part]
        dropping = {id(f) for f in drop}
        survivors = [f for f in links if id(f) not in dropping]
        mapping = {f.links[0]: new for new, f in enumerate(survivors, start=1)}
        for finding in remove:
            self._apply(finding)
            count = sum(e.action == "cell" for e in finding.edits) if finding.type == "cell-formula" else 1
            self.result.removed[finding.type] = self.result.removed.get(finding.type, 0) + count
        if any(old != new for old, new in mapping.items()):
            for finding in kept:
                self._renumber(finding, mapping)
        for finding in drop:
            self._drop_link(finding)
        if drop:
            self.result.removed["external-link"] = len(drop)
        if self.cells_changed:
            for rel in self.pkg.rels(self.wb_part).values():
                if rel.kind == "calcChain" and self.pkg.has(rel.part):  # Excel rebuilds it on the next save
                    self.deleted.add(rel.part or "")
                    self.rel_drops[self.wb_part].add(rel.id)
                    self.result.calc_chain_removed = True
        self.result.kept = [f for f in kept + survivors if f.type != "hyperlink" or "hyperlink" in self.selected]
        for part, nodes in self.removals.items():
            _remove_nodes(self.editor(part), nodes)
        self._drop_rels()
        if self.deleted:
            types = "[Content_Types].xml"
            gone = {d.lower() for d in self.deleted}
            for node in parse_xml(self.pkg.read(types)).findall("Override"):
                if (node.get("PartName") or "").lstrip("/").lower() in gone:
                    self.editor(types).remove(node)
        parts = {p: e.apply() for p, e in self.editors.items() if e.changed and p not in self.deleted}
        return parts, self.deleted

    def _apply(self, finding: Finding) -> None:
        for edit in finding.edits:
            node = edit.node
            if edit.action == "remove" and node is not None:
                self.removals[edit.part].append(node)
            elif edit.action == "cell" and node is not None:
                self._clear_cell(edit.part, node)
            elif edit.action == "chart" and node is not None:
                self._chart_literal(edit.part, node)
            elif edit.action == "blank" and node is not None:
                self.editor(edit.part).set_attr(node, edit.attr, "")
            elif edit.action == "drop-attr" and node is not None:
                self.editor(edit.part).set_attr(node, edit.attr, None)
            elif edit.action == "span":
                self.editor(edit.part).replace(*edit.span, "")
            elif edit.action == "rel":
                self.rel_drops[edit.part].add(edit.attr)

    def _clear_cell(self, part: str, cell: Node) -> None:
        """Keep the cached value of a formula cell and drop the formula."""
        editor = self.editor(part)
        formula, value = cell.find("f"), cell.find("v")
        if formula is not None:
            editor.remove(formula)
        if cell.get("cm") is not None:  # dynamic-array metadata only applies to formulas
            editor.set_attr(cell, "cm", None)
        if value is None:
            self.result.empty_cells += 1
        if cell.get("t") == "str":  # "str" means "formula result"; plain text needs another type
            if value is None:
                editor.set_attr(cell, "t", None)
            else:
                p = prefix_of(editor.data, value)
                editor.set_attr(cell, "t", "inlineStr")
                editor.replace(
                    value.start, value.end, f'<{p}is><{p}t xml:space="preserve">{escape_text(value.text)}</{p}t></{p}is>'
                )
        self.cells_changed = True

    def _chart_literal(self, part: str, ref: Node) -> None:
        """Replace a chart ``numRef``/``strRef`` with literal data taken from its cache."""
        editor = self.editor(part)
        data, p = editor.data, prefix_of(editor.data, ref)
        cache = next((c for c in ref.children if c.local.endswith("Cache")), None)
        holder = ref.parent
        owner = holder.parent if holder is not None else None
        if ref.local == "numRef":
            body = inner_xml(data, cache) or f'<{p}ptCount val="0"/>'
            text = f"<{p}numLit>{body}</{p}numLit>"
        elif holder is not None and holder.local == "tx":
            points = cache.findall("pt") if cache is not None else []
            words = escape_text("".join(v.text for pt in points if (v := pt.find("v")) is not None))
            if owner is not None and owner.local == "ser":
                text = f"<{p}v>{words}</{p}v>"
            else:
                dml = _DRAWINGML.get(ref.ns, next(iter(_DRAWINGML.values())))
                text = f'<{p}rich xmlns:a="{dml}"><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>{words}</a:t></a:r></a:p></{p}rich>'
        elif ref.local == "multiLvlStrRef":
            count = cache.find("ptCount") if cache is not None else None
            head = data[count.start : count.end].decode("utf-8") if count is not None else ""
            text = f"<{p}strLit>{head}{inner_xml(data, cache.find('lvl') if cache is not None else None)}</{p}strLit>"
        else:
            text = f"<{p}strLit>{inner_xml(data, cache)}</{p}strLit>"
        editor.replace(ref.start, ref.end, text)

    def _renumber(self, finding: Finding, mapping: dict[int, int]) -> None:
        """Rewrite [N] link numbers in a finding that stays, after earlier links were removed."""
        if all(mapping.get(i, i) == i for i in finding.links):
            return
        for source in finding.sources:
            editor = self.editor(source.part)
            if source.node is None:
                old = unescape(editor.data[source.span[0] : source.span[1]].decode("utf-8", "replace"))
                editor.replace(source.span[0], source.span[1], escape_text(renumber(old, mapping)))
            elif source.attr:
                editor.set_attr(source.node, source.attr, renumber(source.node.get(source.attr) or "", mapping))
            else:
                editor.set_text(source.node, renumber(source.node.text, mapping))

    def _drop_link(self, finding: Finding) -> None:
        edit = finding.edits[0]
        if edit.node is not None:
            self.removals[self.wb_part].append(edit.node)
            self.rel_drops[self.wb_part].add(edit.node.rid() or "")
        self.deleted |= {p for p in (edit.part, rels_name(edit.part)) if self.pkg.has(p)}
        self.result.removed_links.append(finding.target or finding.text)

    def _drop_rels(self) -> None:
        """Delete requested relationships that no remaining element refers to."""
        for part, ids in self.rel_drops.items():
            rels_part = rels_name(part)
            if part in self.deleted or not self.pkg.has(rels_part):
                continue
            current = self.editors[part].apply() if part in self.editors else self.pkg.read(part)
            still_used, _ = _rel_attr_values(current)
            for rel in parse_xml(self.pkg.read(rels_part)).findall("Relationship"):
                if rel.get("Id") in ids and rel.get("Id") not in still_used:
                    self.editor(rels_part).remove(rel)


def _write_zip(pkg: Package, dst: Path, parts: dict[str, bytes], deleted: set[str]) -> None:
    """Copy the package to ``dst``, replacing edited parts and leaving out deleted ones."""
    gone = {d.lower() for d in deleted}
    new = {k.lower(): v for k, v in parts.items()}
    tmp = dst.with_name(dst.name + ".part")
    try:
        with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as out:
            for info in pkg.zip.infolist():
                low = info.filename.lower()
                if low in gone:
                    continue
                item = zipfile.ZipInfo(info.filename, info.date_time)
                item.compress_type = zipfile.ZIP_DEFLATED
                out.writestr(item, new[low] if low in new else pkg.zip.read(info))
        os.replace(tmp, dst)
    finally:
        if tmp.exists():
            tmp.unlink()
