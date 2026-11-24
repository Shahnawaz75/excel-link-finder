"""Scan .xlsx / .xlsm / .xltx / .xltm workbooks for external links, without Excel."""

from __future__ import annotations

import base64
import io
import os
import posixpath
import re
import struct
import threading
import zipfile
from collections.abc import Iterable
from dataclasses import dataclass, field
from html import unescape
from pathlib import Path
from urllib.parse import unquote
from xml.parsers import expat

from .formula import external_refs, name_refs
from .ooxml import Node, Package, Rel

DATA_LINK_TYPES = (
    "external-link",
    "defined-name",
    "cell-formula",
    "table-formula",
    "conditional-format",
    "data-validation",
    "sparkline",
    "chart",
    "pivot-cache",
    "shape-macro",
    "object-link",
)
INFO_TYPES = ("connection", "hyperlink")
ALL_TYPES = DATA_LINK_TYPES + INFO_TYPES
SUPPORTED_SUFFIXES = (".xlsx", ".xlsm", ".xltx", ".xltm")

X14 = "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main"
CHARTEX = "http://schemas.microsoft.com/office/drawing/2014/chartex"
_OLE2 = b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1"
_URL = re.compile(r"^[a-z][a-z0-9+.-]*://", re.I)
_WIN_ABS = re.compile(r"^(?:[A-Za-z]:[\\/]|\\\\|//)")
_SECRET = re.compile(r"(?i)\b(password|pwd)\s*=\s*[^;]*")
_CONNECTION_KINDS = {"1": "ODBC", "2": "DAO", "3": "file", "4": "web query", "5": "OLE DB", "6": "text file", "7": "ADO"}
_CONTROL_NOTES = {
    "fmlaLink": "form control cell link",
    "fmlaRange": "form control input range",
    "fmlaTxbx": "form control text link",
    "fmlaGroup": "form control group link",
}
_VML_SHAPE = re.compile(rb"<v:shape\b[^>]*?\bid=\"([^\"]*)\"[^>]*>(.*?)</v:shape>", re.S)
_VML_FMLA = re.compile(rb"<((?:\w+:)?)(Fmla(?:Macro|Link|Range|Txbx|Group))>([^<]*)</\1\2>")
_M_SHARED = re.compile(r'shared\s+(#"(?:[^"]|"")*"|[^\s=]+)\s*=')
_M_SOURCE = re.compile(
    r"\b((?:File|Folder|Web|Sql|Odbc|OleDb|OData|SharePoint|Access|Oracle|MySQL|PostgreSQL|"
    r"AnalysisServices|Exchange|Salesforce|Snowflake|AzureStorage)\.\w+)\s*\(\s*\"((?:[^\"]|\"\")*)\""
)
_CHART_ROLES = {"cat": "categories", "val": "values", "xVal": "x values", "yVal": "y values", "bubbleSize": "bubble sizes"}


class ExcelLinkFinderError(Exception):
    """Base class for errors raised by excel-link-finder."""


class UnsupportedFileError(ExcelLinkFinderError):
    """The file is not a workbook this tool can read (for example an old .xls file)."""


@dataclass
class Edit:
    """A change the cleaner makes to remove a finding (internal)."""

    action: str  # remove | cell | chart | blank | drop-attr | span | rel | link
    part: str
    node: Node | None = None
    attr: str = ""
    span: tuple[int, int] = (0, 0)


@dataclass
class Source:
    """Where the formula text of a finding is stored, so link numbers can be rewritten (internal)."""

    part: str
    node: Node | None = None  # None: a text span in a VML part
    attr: str = ""  # attribute name; "" means the element text
    span: tuple[int, int] = (0, 0)


@dataclass
class Finding:
    """One place in a workbook that refers to another file."""

    type: str
    sheet: str = ""
    object: str = ""
    ref: str = ""
    text: str = ""
    links: list[int] = field(default_factory=list)
    target: str = ""
    status: str = ""
    note: str = ""
    cleanable: bool = True
    edits: list[Edit] = field(default_factory=list, repr=False, compare=False)
    sources: list[Source] = field(default_factory=list, repr=False, compare=False)

    @property
    def is_data_link(self) -> bool:
        """True for links to other workbooks (False for connections and hyperlinks)."""
        return self.type in DATA_LINK_TYPES

    def to_dict(self) -> dict[str, object]:
        """Plain dictionary for JSON output."""
        keys = ("type", "sheet", "object", "ref", "text", "links", "target", "status", "note")
        return {key: getattr(self, key) for key in keys}


@dataclass
class ScanResult:
    """Everything found in one workbook."""

    path: Path
    findings: list[Finding] = field(default_factory=list)
    error: str | None = None

    @property
    def links(self) -> list[Finding]:
        """Findings that are data links (everything except connections and hyperlinks)."""
        return [f for f in self.findings if f.is_data_link]

    @property
    def has_links(self) -> bool:
        return bool(self.links)

    @property
    def broken(self) -> list[Finding]:
        """External links whose target file does not exist on this computer."""
        return [f for f in self.findings if f.type in ("external-link", "pivot-cache") and f.status == "missing"]

    def counts(self) -> dict[str, int]:
        """Number of findings per type, in a stable order (types with no findings are left out)."""
        return {t: n for t in ALL_TYPES if (n := sum(f.type == t for f in self.findings))}

    def to_dict(self) -> dict[str, object]:
        summary: dict[str, int] = {**self.counts(), "broken": len(self.broken)}
        return {
            "file": str(self.path),
            "error": self.error,
            "has_links": self.has_links,
            "summary": summary,
            "findings": [f.to_dict() for f in self.findings],
        }


@dataclass
class _Hit:
    links: set[int]
    files: list[str]
    via: list[str]


def check_format(path: Path) -> None:
    """Raise :class:`UnsupportedFileError` unless ``path`` looks like an .xlsx-family workbook."""
    if not path.is_file():
        raise ExcelLinkFinderError(f"{path}: no such file")
    with path.open("rb") as handle:
        head = handle.read(8)
    suffix = path.suffix.lower()
    if head == _OLE2:
        if suffix in SUPPORTED_SUFFIXES:
            raise UnsupportedFileError(
                f"{path.name} is encrypted (password to open). Remove the password in Excel and try again."
            )
        raise UnsupportedFileError(
            f"{path.name} is an old binary Excel file (.xls, Excel 97-2003). It is not supported: "
            "open it in Excel and save it as .xlsx (File > Save As > Excel Workbook) first."
        )
    if suffix == ".xlsb":
        raise UnsupportedFileError(f"{path.name} is a binary workbook (.xlsb). It is not supported: save it as .xlsx first.")
    if not head.startswith(b"PK"):
        raise UnsupportedFileError(f"{path.name} is not an Excel workbook (.xlsx, .xlsm, .xltx or .xltm).")


def scan(path: str | Path, *, check_paths: bool = True) -> ScanResult:
    """Scan one workbook and return every external link location in it.

    ``check_paths`` looks up local and network paths of linked files to flag broken links
    (URLs are never fetched). Raises :class:`ExcelLinkFinderError` for unreadable files.
    """
    path = Path(path)
    check_format(path)
    try:
        with Package(path) as pkg:
            return ScanResult(path, _Scanner(pkg, path, check_paths).run())
    except (zipfile.BadZipFile, expat.ExpatError, KeyError, ValueError) as exc:
        raise ExcelLinkFinderError(f"{path.name} could not be read: {exc}") from exc


def find_workbooks(paths: Iterable[str | Path], recursive: bool = False) -> tuple[list[Path], list[Path]]:
    """Expand files and folders into workbook paths.

    Returns ``(workbooks, skipped)``. Files given directly are always returned; inside
    folders only .xlsx/.xlsm/.xltx/.xltm files are picked up, Office lock files (``~$*``)
    are ignored and old .xls/.xlsb files are returned in ``skipped``.
    """
    found: list[Path] = []
    skipped: list[Path] = []
    for path in map(Path, paths):
        if not path.is_dir():
            found.append(path)
            continue
        for child in sorted(path.rglob("*") if recursive else path.glob("*")):
            if not child.is_file() or child.name.startswith("~$"):
                continue
            suffix = child.suffix.lower()
            if suffix in SUPPORTED_SUFFIXES:
                found.append(child)
            elif suffix in (".xls", ".xlsb"):
                skipped.append(child)
    return found, skipped


def _display_target(raw: str) -> str:
    """Readable form of a link target from a .rels file."""
    low = raw.lower()
    text = raw[8:] if low.startswith("file:///") else "//" + raw[7:] if low.startswith("file://") else raw
    return unquote(text) if "%" in text and not _URL.match(text) else text


def _exists(path: Path, timeout: float | None) -> str:
    box: list[bool] = []

    def probe() -> None:
        try:
            box.append(path.exists())
        except (OSError, ValueError):
            box.append(False)

    if timeout is None:
        probe()
    else:  # network paths can hang for a long time when the server is gone
        thread = threading.Thread(target=probe, daemon=True)
        thread.start()
        thread.join(timeout)
    return ("found" if box[0] else "missing") if box else "unknown"


def link_status(raw: str, base: Path) -> str:
    """``found``, ``missing``, ``url`` (never fetched) or ``unknown`` for a link target."""
    text = _display_target(raw)
    if _URL.match(text):
        return "url"
    if os.name == "nt":
        if text.startswith("/") and not text.startswith("//"):
            return "unknown"  # a macOS / Linux absolute path
        path = Path(text) if _WIN_ABS.match(text) else base / text
    else:
        if _WIN_ABS.match(text):
            return "unknown"  # a Windows path cannot be checked on this system
        path = Path(text.replace("\\", "/"))
        path = path if path.is_absolute() else base / path
    return _exists(path, 3.0 if str(path).startswith(("\\\\", "//")) else None)


def _source_text(source: Source) -> str:
    if source.node is None:
        return ""
    return (source.node.get(source.attr) or "") if source.attr else source.node.text


def _shape_name(node: Node | None) -> str:
    """Name of the drawing object that owns ``node`` (taken from its ``cNvPr`` element)."""
    while node is not None:
        if node.local == "cNvPr":
            return node.get("name") or ""
        for child in node.children:
            if child.local.startswith("nv") and child.local.endswith("Pr") and child.find("cNvPr") is not None:
                return child.find("cNvPr").get("name") or ""  # type: ignore[union-attr]
        node = node.parent
    return ""


def _chart_role(formula: Node) -> tuple[str, Node | None, str]:
    """What a chart formula feeds, which element the cleaner replaces, and how."""
    ref = formula.parent
    if formula.ns == CHARTEX:
        return "chart data", formula, "remove"
    holder = ref.parent if ref is not None else None
    if ref is not None and holder is not None and ref.local in ("numRef", "strRef", "multiLvlStrRef"):
        series = ref.ancestor("ser")
        idx = series.find("idx") if series is not None else None
        number = f"series {int(idx.get('val') or 0) + 1} " if idx is not None else ""
        if holder.local == "tx":
            owner = holder.parent
            if owner is not None and owner.local == "ser":
                return f"{number}name", ref, "chart"
            if owner is not None and owner.local == "title":
                parent = owner.parent
                return ("chart title" if parent is not None and parent.local == "chart" else "axis title"), ref, "chart"
            return "label", ref, "chart"
        return f"{number}{_CHART_ROLES.get(holder.local, holder.local)}", ref, "chart"
    ext = formula.ancestor("ext")
    if ext is not None:
        return "chart extension (data label range)", ext, "remove"
    return "chart reference", None, ""


def _mashup_queries(blob: bytes) -> list[tuple[str, list[tuple[str, str]]]]:
    """Query names and data sources from a Power Query ``DataMashup`` blob (MS-QDEFF format)."""
    length = struct.unpack_from("<I", blob, 4)[0]
    with zipfile.ZipFile(io.BytesIO(blob[8 : 8 + length])) as package:
        code = package.read("Formulas/Section1.m").decode("utf-8-sig")
    starts = list(_M_SHARED.finditer(code))
    queries = []
    for i, match in enumerate(starts):
        body = code[match.end() : starts[i + 1].start() if i + 1 < len(starts) else len(code)]
        name = match.group(1)
        if name.startswith('#"'):
            name = name[2:-1].replace('""', '"')
        queries.append((name, [(m.group(1), m.group(2).replace('""', '"')) for m in _M_SOURCE.finditer(body)]))
    return queries


class _Scanner:
    """Walks the parts of one package and collects :class:`Finding` objects."""

    def __init__(self, pkg: Package, path: Path, check_paths: bool) -> None:
        self.pkg = pkg
        self.base = path.resolve().parent
        self.check_paths = check_paths
        self.findings: list[Finding] = []
        self.wb_part = ""
        self.sheets: list[str] = []
        self.link_targets: dict[int, str] = {}
        self.all_names: set[tuple[str | None, str]] = set()
        self.ext_names: dict[tuple[str | None, str], tuple[str, set[int]]] = {}
        self.pivots: dict[str, list[tuple[str, str]]] = {}

    def run(self) -> list[Finding]:
        office = [r for r in self.pkg.rels("").values() if r.kind == "officeDocument" and self.pkg.has(r.part)]
        if not office:
            raise UnsupportedFileError("this file has no workbook part, so it is not an Excel workbook")
        self.wb_part = office[0].part or ""
        if "binary" in self.pkg.content_type(self.wb_part).lower():
            raise UnsupportedFileError("binary workbooks (.xlsb) are not supported: save the file as .xlsx first")
        workbook = self.pkg.xml(self.wb_part)
        rels = self.pkg.rels(self.wb_part)
        self._external_links(workbook, rels)
        sheets = [(s.get("name") or "", rels.get(s.rid() or "")) for s in workbook.iter("sheet")]
        self.sheets = [name for name, _ in sheets]
        self._defined_names(workbook)
        for name, rel in sheets:
            if rel is not None and self.pkg.has(rel.part):
                self._sheet(name, rel.part or "", chartsheet=rel.kind == "chartsheet")
        self._pivot_caches(workbook, rels)
        for rel in rels.values():
            if rel.kind == "connections" and self.pkg.has(rel.part):
                self._connections(rel.part or "")
        self._power_query()
        return self.findings

    # -- helpers ---------------------------------------------------------------------------

    def _analyze(self, text: str | None, sheet: str | None) -> _Hit | None:
        """Return the links used by a formula, or None if it stays inside the workbook."""
        if not text:
            return None
        refs = external_refs(text)
        hit = _Hit({r.index for r in refs if r.index}, [r.target for r in refs if r.index is None], [])
        if self.ext_names:
            for scope, name in name_refs(text):
                if scope is None and sheet and (sheet.lower(), name) in self.all_names:
                    scope = sheet.lower()  # a sheet-level name hides a workbook-level one
                found = self.ext_names.get((scope, name))
                if found:
                    hit.via.append(found[0])
                    hit.links |= found[1]
        return hit if hit.links or hit.files or hit.via else None

    def _finding(self, ftype: str, hits: list[_Hit], **fields: object) -> Finding:
        links = sorted({i for h in hits for i in h.links})
        targets = [self.link_targets.get(i, f"[{i}]") for i in links] + [f for h in hits for f in h.files]
        via = list(dict.fromkeys(v for h in hits for v in h.via))
        note = ", ".join(x for x in (str(fields.pop("note", "")), "via name " + ", ".join(via) if via else "") if x)
        finding = Finding(ftype, links=links, target="; ".join(dict.fromkeys(targets)), note=note, **fields)  # type: ignore[arg-type]
        self.findings.append(finding)
        return finding

    def _rule(self, ftype: str, sheet: str, part: str, nodes: list[Node], sources: list[Source], ref: str, note: str) -> None:
        """Report a CF / DV / sparkline rule if any of its formulas is external."""
        texts = [(s, _source_text(s)) for s in sources]
        hits = [(s, h) for s, t in texts if (h := self._analyze(t, sheet))]
        if hits:
            text = " ; ".join(t for _, t in texts if t)
            finding = self._finding(ftype, [h for _, h in hits], sheet=sheet, ref=ref, text=text, note=note)
            finding.edits += [Edit("remove", part, n) for n in nodes]
            finding.sources += [s for s, _ in hits]

    def _status(self, raw: str) -> str:
        return link_status(raw, self.base) if self.check_paths else "not-checked"

    # -- workbook level --------------------------------------------------------------------

    def _external_links(self, workbook: Node, rels: dict[str, Rel]) -> None:
        for index, ref in enumerate(workbook.iter("externalReference"), start=1):
            rel = rels.get(ref.rid() or "")
            part = rel.part if rel is not None else None
            finding = Finding("external-link", object=f"[{index}]", links=[index], note="workbook link")
            finding.edits.append(Edit("link", part or "", ref))
            self.findings.append(finding)
            if part is None or not self.pkg.has(part):
                finding.text, finding.status = "(link part missing)", "unknown"
                continue
            finding.text = posixpath.basename(part)
            link_rels = self.pkg.rels(part)
            root = self.pkg.xml(part, keep=lambda n: n.local != "row")  # skip the cached cell values
            body = next((c for c in root.children if c.local in ("externalBook", "ddeLink", "oleLink")), None)
            target = link_rels.get(body.rid() or "") if body is not None else None
            if body is not None and body.local == "ddeLink":
                finding.note, finding.target = "DDE link", f"{body.get('ddeService')}|{body.get('ddeTopic')}"
            else:
                if body is not None and body.local == "oleLink":
                    finding.note, finding.cleanable = f"OLE link ({body.get('progId')})", False
                raw = target.target if target is not None else ""
                finding.target = _display_target(raw)
                alt_nodes = body.iter("absoluteUrl", "relativeUrl") if body is not None else iter(())
                alternates = [link_rels[n.rid() or ""].target for n in alt_nodes if (n.rid() or "") in link_rels]
                finding.status = self._status(raw) if raw else "missing"
                if finding.status == "missing" and any(self._status(a) == "found" for a in alternates):
                    finding.status = "found"
                others = [t for t in map(_display_target, alternates) if t != finding.target]
                if others and finding.status == "missing":
                    finding.note += ", last seen at " + ", ".join(others)
            self.link_targets[index] = finding.target

    def _defined_names(self, workbook: Node) -> None:
        names = []
        for node in workbook.iter("definedName"):
            sid = node.get("localSheetId") or ""
            scope = self.sheets[int(sid)] if sid.isdigit() and int(sid) < len(self.sheets) else None
            key = (scope.lower() if scope else None, (node.get("name") or "").lower())
            self.all_names.add(key)
            names.append((key, scope, node))
        hits: dict[int, _Hit] = {}
        while True:  # repeat so that names pointing at external names are found too
            before = len(hits)
            for i, (key, scope, node) in enumerate(names):
                if i not in hits and (hit := self._analyze(node.text, scope)):
                    hits[i] = hit
                    self.ext_names[key] = (node.get("name") or "", hit.links)
            if len(hits) == before:
                break
        for i, (_key, scope, node) in enumerate(names):
            if i in hits:
                hidden = "hidden" if node.get("hidden") in ("1", "true") else ""
                finding = self._finding(
                    "defined-name", [hits[i]], sheet=scope or "", object=node.get("name") or "", text=node.text, note=hidden
                )
                finding.edits.append(Edit("remove", self.wb_part, node))
                finding.sources.append(Source(self.wb_part, node))

    def _pivot_caches(self, workbook: Node, rels: dict[str, Rel]) -> None:
        for cache in workbook.iter("pivotCache"):
            rel = rels.get(cache.rid() or "")
            if rel is None or not self.pkg.has(rel.part):
                continue
            part = rel.part or ""
            source = self.pkg.xml(part).find("cacheSource")
            cache_rels = self.pkg.rels(part)
            for node in source.iter("worksheetSource", "rangeSet") if source is not None else []:
                link = cache_rels.get(node.rid() or "")
                hit = self._analyze(node.get("name"), None)
                if not hit and not (link is not None and link.external):
                    continue
                where = self.pivots.get(part.lower(), [])
                text = node.get("name") or "!".join(x for x in (node.get("sheet"), node.get("ref")) if x)
                finding = self._finding(
                    "pivot-cache",
                    [hit] if hit else [],
                    sheet=", ".join(dict.fromkeys(s for s, _ in where)),
                    object=", ".join(n for _, n in where) or posixpath.basename(part),
                    ref=node.get("ref") or "",
                    text=text,
                    cleanable=False,
                )
                if link is not None and link.external:
                    finding.target = _display_target(link.target)
                    finding.status = self._status(link.target)

    def _connections(self, part: str) -> None:
        for conn in self.pkg.xml(part).iter("connection"):
            kind = _CONNECTION_KINDS.get(conn.get("type") or "", "connection")
            db, text_pr, web = conn.find("dbPr"), conn.find("textPr"), conn.find("webPr")
            text = _SECRET.sub(r"\1=***", db.get("connection") or "") if db is not None else ""
            target = ""
            if text_pr is not None:
                target = text_pr.get("sourceFile") or ""
            elif web is not None:
                target = web.get("url") or ""
            elif match := re.search(r"(?i)(?:^|;)\s*Data Source=([^;]*)", text):
                target = match.group(1)
            if "Microsoft.Mashup" in text:
                location = re.search(r"Location=([^;]*)", text)
                kind, target = "Power Query", f"query {location.group(1)}" if location else target
            elif target.lower() == "$embedded$":
                kind, target = "data model", "this workbook"
            self.findings.append(
                Finding(
                    "connection",
                    object=conn.get("name") or "",
                    text=text or target,
                    target=target or conn.get("odcFile") or "",
                    note=kind,
                    cleanable=False,
                )
            )

    def _power_query(self) -> None:
        for name in self.pkg.names():
            if not re.fullmatch(r"customxml/item\d+\.xml", name.lower()):
                continue
            try:
                root = self.pkg.xml(name)
                if root.local != "DataMashup":
                    continue
                queries = _mashup_queries(base64.b64decode(root.text))
            except (expat.ExpatError, ValueError, KeyError, struct.error, zipfile.BadZipFile):
                self.findings.append(
                    Finding("connection", object="Power Query", note="Power Query (could not be decoded)", cleanable=False)
                )
                continue
            for query, sources in queries:
                text = "; ".join(f'{fn}("{arg}")' for fn, arg in sources)
                self.findings.append(
                    Finding(
                        "connection",
                        object=query,
                        text=text,
                        target="; ".join(a for _, a in sources),
                        note="Power Query",
                        cleanable=False,
                    )
                )

    # -- sheet level -----------------------------------------------------------------------

    def _sheet(self, sheet: str, part: str, chartsheet: bool = False) -> None:
        rels = self.pkg.rels(part)
        shared: dict[str, Finding | None] = {}

        def keep(node: Node) -> bool:  # drop cells without external formulas while parsing
            if node.local == "c":
                formula = node.find("f")
                return formula is not None and self._cell(sheet, part, node, formula, shared)
            return node.local != "row" or bool(node.children)

        root = self.pkg.xml(part, keep=keep)
        for group in shared.values():
            if group is not None and len(group.edits) > 1:
                group.note = ", ".join(x for x in (f"shared formula, {len(group.edits)} cells", group.note) if x)
        controls: dict[str, dict] = {}
        if not chartsheet:
            self._conditional_formats(sheet, part, root)
            self._validations(sheet, part, root)
            self._sparklines(sheet, part, root)
            self._hyperlinks(sheet, part, root, rels)
            controls = self._controls(sheet, part, root, rels)
            self._ole_objects(sheet, part, root)
        for rel in rels.values():
            if not self.pkg.has(rel.part):
                continue
            target = rel.part or ""
            if rel.kind == "drawing":
                self._drawing(sheet, target)
            elif rel.kind == "vmlDrawing":
                self._vml(sheet, target, controls)
            elif rel.kind == "table":
                self._table(sheet, target)
            elif rel.kind == "pivotTable":
                table = self.pkg.xml(target)
                for cache in self.pkg.rels(target).values():
                    if cache.kind == "pivotCacheDefinition" and cache.part:
                        self.pivots.setdefault(cache.part.lower(), []).append((sheet, table.get("name") or ""))

    def _cell(self, sheet: str, part: str, cell: Node, formula: Node, shared: dict[str, Finding | None]) -> bool:
        kind, si, text = formula.get("t"), formula.get("si"), formula.text
        if kind == "shared" and si is not None and not text.strip():  # follower of a shared formula
            group = shared.get(si)
            if group is not None:
                group.edits.append(Edit("cell", part, cell))
            return group is not None
        hit = self._analyze(text, sheet)
        if kind == "shared" and si is not None:
            shared[si] = None
        if hit is None:
            return False
        ref = (formula.get("ref") if kind == "shared" else None) or cell.get("r") or ""
        note = f"array formula {formula.get('ref') or ''}".strip() if kind == "array" else ""
        finding = self._finding("cell-formula", [hit], sheet=sheet, ref=ref, text=text, note=note)
        finding.edits.append(Edit("cell", part, cell))
        finding.sources.append(Source(part, formula))
        if kind == "shared" and si is not None:
            shared[si] = finding
        return True

    def _conditional_formats(self, sheet: str, part: str, root: Node) -> None:
        x14_rules = {r.get("id"): r for r in root.iter("cfRule") if r.ns == X14 and r.get("id")}
        partner: dict[int, Node] = {}  # main rule <-> x14 rule (data bars and the like use both)
        for rule in root.iter("cfRule"):
            if rule.ns != X14:
                for ident in rule.iter("id"):
                    other = x14_rules.get(ident.text.strip())
                    if ident.ns == X14 and other is not None:
                        partner[id(rule)], partner[id(other)] = other, rule
        for block in root.iter("conditionalFormatting"):
            x14 = block.ns == X14
            sqref = block.find("sqref")
            ref = (sqref.text if x14 and sqref is not None else block.get("sqref")) or ""
            for rule in block.findall("cfRule"):
                if x14:
                    sources = [Source(part, n) for n in rule.iter("f")]
                else:
                    sources = [Source(part, n) for n in rule.findall("formula")]
                    sources += [Source(part, n, "val") for n in rule.iter("cfvo") if n.get("val")]
                nodes = [rule] + ([partner[id(rule)]] if id(rule) in partner else [])
                self._rule("conditional-format", sheet, part, nodes, sources, ref, f"{rule.get('type') or 'cf'} rule")

    def _validations(self, sheet: str, part: str, root: Node) -> None:
        for rule in root.iter("dataValidation"):
            if rule.ns == X14:
                sqref = rule.find("sqref")
                sources, ref = [Source(part, n) for n in rule.iter("f")], sqref.text if sqref is not None else ""
            else:
                nodes = [rule.find("formula1"), rule.find("formula2")]
                sources, ref = [Source(part, n) for n in nodes if n is not None], rule.get("sqref") or ""
            self._rule("data-validation", sheet, part, [rule], sources, ref, f"{rule.get('type') or 'any'} rule")

    def _sparklines(self, sheet: str, part: str, root: Node) -> None:
        for group in root.iter("sparklineGroup"):
            axis = group.find("f")
            if axis is not None:
                self._rule("sparkline", sheet, part, [group], [Source(part, axis)], "", "date axis")
            for line in group.iter("sparkline"):
                sqref = line.find("sqref")
                self._rule(
                    "sparkline",
                    sheet,
                    part,
                    [line],
                    [Source(part, n) for n in line.findall("f")],
                    sqref.text if sqref is not None else "",
                    "",
                )

    def _hyperlinks(self, sheet: str, part: str, root: Node, rels: dict[str, Rel]) -> None:
        for link in root.iter("hyperlink"):
            rel = rels.get(link.rid() or "")
            if rel is not None and rel.external:
                finding = Finding(
                    "hyperlink", sheet=sheet, ref=link.get("ref") or "", text=link.get("display") or "", target=rel.target
                )
                finding.edits += [Edit("remove", part, link), Edit("rel", part, attr=rel.id)]
                self.findings.append(finding)

    def _controls(self, sheet: str, part: str, root: Node, rels: dict[str, Rel]) -> dict[str, dict]:
        """Form controls (buttons, drop-downs ...) by shape id."""
        controls: dict[str, dict] = {}
        for control in root.iter("control"):
            shape_id = control.get("shapeId") or ""
            if shape_id in controls:
                continue  # the mc:Fallback copy of the same control
            info: dict = {"name": control.get("name") or f"control {shape_id}", "macro": None, "links": {}}
            controls[shape_id] = info
            props = control.find("controlPr")
            if props is not None and (hit := self._analyze(props.get("macro"), sheet)):
                info["macro"] = self._finding(
                    "shape-macro", [hit], sheet=sheet, object=info["name"], text=props.get("macro"), note="form control macro"
                )
                info["macro"].edits.append(Edit("drop-attr", part, props, "macro"))
                info["macro"].sources.append(Source(part, props, "macro"))
            rel = rels.get(control.rid() or "")
            if rel is None or not self.pkg.has(rel.part):
                continue
            form = self.pkg.xml(rel.part or "")
            for attr, note in _CONTROL_NOTES.items():
                if hit := self._analyze(form.get(attr), sheet):
                    finding = self._finding(
                        "object-link", [hit], sheet=sheet, object=info["name"], text=form.get(attr), note=note
                    )
                    finding.edits.append(Edit("drop-attr", rel.part or "", form, attr))
                    finding.sources.append(Source(rel.part or "", form, attr))
                    info["links"][attr.lower()] = finding
        return controls

    def _vml(self, sheet: str, part: str, controls: dict[str, dict]) -> None:
        """Legacy VML copy of form controls: Excel keeps macros and links here too."""
        data = self.pkg.read(part)
        for shape in _VML_SHAPE.finditer(data):
            spid = shape.group(1).decode("utf-8", "replace")
            info = controls.setdefault(re.sub(r"\D", "", spid.rpartition("_s")[2]), {"name": spid, "macro": None, "links": {}})
            for match in _VML_FMLA.finditer(data, shape.start(2), shape.end(2)):
                value = unescape(match.group(3).decode("utf-8", "replace"))
                hit = self._analyze(value, sheet)
                if hit is None:
                    continue
                kind = match.group(2).decode()
                finding = info["macro"] if kind == "FmlaMacro" else info["links"].get(kind.lower())
                if finding is None:
                    ftype = "shape-macro" if kind == "FmlaMacro" else "object-link"
                    note = (
                        "form control macro" if kind == "FmlaMacro" else _CONTROL_NOTES.get("f" + kind[1:], "form control link")
                    )
                    finding = self._finding(ftype, [hit], sheet=sheet, object=info["name"], text=value, note=note)
                    if kind == "FmlaMacro":
                        info["macro"] = finding
                    else:
                        info["links"][kind.lower()] = finding
                finding.edits.append(Edit("span", part, span=(match.start(), match.end())))
                finding.sources.append(Source(part, span=(match.start(3), match.end(3))))

    def _ole_objects(self, sheet: str, part: str, root: Node) -> None:
        seen: dict[str, Finding] = {}
        for obj in root.iter("oleObject"):
            hit = self._analyze(obj.get("link"), sheet)
            if hit is None:
                continue
            key = obj.get("shapeId") or str(id(obj))
            if key not in seen:
                seen[key] = self._finding(
                    "object-link",
                    [hit],
                    sheet=sheet,
                    object=obj.get("progId") or "OLE object",
                    text=obj.get("link"),
                    note="linked OLE object",
                    cleanable=False,
                )
            seen[key].sources.append(Source(part, obj, "link"))

    def _drawing(self, sheet: str, part: str) -> None:
        rels = self.pkg.rels(part)
        for node in self.pkg.xml(part).iter():
            for attr, ftype, note in (
                ("macro", "shape-macro", "macro assigned to a shape"),
                ("textlink", "object-link", "shape text linked to a cell"),
            ):
                if hit := self._analyze(node.attrib.get(attr), sheet):
                    finding = self._finding(
                        ftype, [hit], sheet=sheet, object=_shape_name(node), text=node.attrib[attr], note=note
                    )
                    finding.edits.append(Edit("blank", part, node, attr))
                    finding.sources.append(Source(part, node, attr))
            rel = rels.get(node.rid() or "")
            if rel is None:
                continue
            if node.local in ("hlinkClick", "hlinkHover") and rel.external:
                finding = Finding("hyperlink", sheet=sheet, object=_shape_name(node), target=rel.target, note="shape hyperlink")
                finding.edits += [Edit("remove", part, node), Edit("rel", part, attr=rel.id)]
                self.findings.append(finding)
            elif node.local == "chart" and self.pkg.has(rel.part):
                self._chart(sheet, rel.part or "", _shape_name(node))

    def _chart(self, sheet: str, part: str, name: str) -> None:
        name = name or posixpath.basename(part)
        for formula in self.pkg.xml(part).iter("f"):
            if hit := self._analyze(formula.text, None):
                role, container, action = _chart_role(formula)
                finding = self._finding(
                    "chart", [hit], sheet=sheet, object=name, text=formula.text, note=role, cleanable=container is not None
                )
                if container is not None:
                    finding.edits.append(Edit(action, part, container))
                finding.sources.append(Source(part, formula))
        for rel in self.pkg.rels(part).values():
            if rel.kind == "chartUserShapes" and self.pkg.has(rel.part):
                self._drawing(sheet, rel.part or "")
            elif rel.external and rel.kind in ("package", "oleObject"):
                self.findings.append(
                    Finding(
                        "chart",
                        sheet=sheet,
                        object=name,
                        text="external chart data",
                        target=_display_target(rel.target),
                        note="chart data kept in another file",
                        cleanable=False,
                    )
                )

    def _table(self, sheet: str, part: str) -> None:
        root = self.pkg.xml(part)
        table = root.get("displayName") or root.get("name") or posixpath.basename(part)
        for column in root.iter("tableColumn"):
            for local, note in (("calculatedColumnFormula", "calculated column"), ("totalsRowFormula", "totals row")):
                node = column.find(local)
                if node is not None and (hit := self._analyze(node.text, sheet)):
                    finding = self._finding(
                        "table-formula",
                        [hit],
                        sheet=sheet,
                        object=f"{table}[{column.get('name')}]",
                        ref=root.get("ref") or "",
                        text=node.text,
                        note=note,
                    )
                    finding.edits.append(Edit("remove", part, node))
                    if local == "totalsRowFormula" and column.get("totalsRowFunction") == "custom":
                        finding.edits.append(Edit("drop-attr", part, column, "totalsRowFunction"))
                    finding.sources.append(Source(part, node))
