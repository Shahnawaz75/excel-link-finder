"""Low-level OOXML helpers: zip parts, relationships and a tiny XML tree with byte offsets.

Excel is strict about the XML it reads. Re-serialising a part with ElementTree renames
namespace prefixes and drops declarations that are only used by ``mc:Ignorable``, which
is enough to trigger Excel's "We found a problem with some content" repair prompt.
So this module never re-serialises: :func:`parse_xml` records where every element
starts and ends in the original bytes, and :class:`Editor` splices changes into them.
"""

from __future__ import annotations

import posixpath
import re
import zipfile
from collections.abc import Callable, Iterator
from dataclasses import dataclass
from pathlib import Path
from urllib.parse import unquote
from xml.parsers import expat

_HEAD = re.compile(rb"""<[^\s/>]+(?:\s+[^\s=/>]+\s*=\s*(?:"[^"]*"|'[^']*'))*\s*/?>""")


class Node:
    """An XML element plus its byte offsets in the part it was parsed from."""

    __slots__ = ("ns", "local", "attrib", "children", "parent", "text", "start", "end", "tail_start")

    def __init__(self, ns: str, local: str, attrib: dict[str, str], parent: Node | None, start: int) -> None:
        self.ns = ns
        self.local = local
        self.attrib = attrib
        self.parent = parent
        self.children: list[Node] = []
        self.text = ""
        self.start = start
        self.end = start
        self.tail_start = -1  # offset of "</tag>"; -1 for an empty-element tag like <a/>

    def get(self, name: str, default: str | None = None) -> str | None:
        """Return an attribute by local name, whatever its namespace."""
        if name in self.attrib:
            return self.attrib[name]
        for key, value in self.attrib.items():
            if key.endswith("}" + name):
                return value
        return default

    def rid(self) -> str | None:
        """Return the ``r:id`` attribute (relationship id), if any."""
        for key, value in self.attrib.items():
            if key.endswith("/relationships}id"):
                return value
        return None

    def find(self, local: str) -> Node | None:
        """First direct child with this local name."""
        return next((c for c in self.children if c.local == local), None)

    def findall(self, local: str) -> list[Node]:
        """All direct children with this local name."""
        return [c for c in self.children if c.local == local]

    def iter(self, *locals_: str) -> Iterator[Node]:
        """Depth-first walk over this node and its descendants, optionally filtered by local name."""
        stack = [self]
        while stack:
            node = stack.pop()
            if not locals_ or node.local in locals_:
                yield node
            stack.extend(reversed(node.children))

    def ancestor(self, *locals_: str) -> Node | None:
        """Nearest ancestor with one of these local names."""
        node = self.parent
        while node is not None and node.local not in locals_:
            node = node.parent
        return node


def parse_xml(data: bytes, keep: Callable[[Node], bool] | None = None) -> Node:
    """Parse ``data`` into a :class:`Node` tree.

    ``keep`` is called when each element closes; returning False drops the element from
    its parent. The scanner uses this to keep only interesting cells of huge sheets.
    """
    parser = expat.ParserCreate(namespace_separator=" ")
    parser.buffer_text = True
    stack: list[Node] = []
    roots: list[Node] = []

    def start(name: str, attrs: dict[str, str]) -> None:
        ns, _, local = name.rpartition(" ")
        attrib = {(f"{{{k.rpartition(' ')[0]}}}{k.rpartition(' ')[2]}" if " " in k else k): v for k, v in attrs.items()}
        parent = stack[-1] if stack else None
        node = Node(ns, local, attrib, parent, parser.CurrentByteIndex)
        (parent.children if parent is not None else roots).append(node)
        stack.append(node)

    def end(_name: str) -> None:
        node = stack.pop()
        pos = parser.CurrentByteIndex
        # For an empty-element tag (<a/>) expat reports the position just after "/>";
        # otherwise it reports where the end tag "</a>" starts.
        head = _HEAD.match(data, node.start) if data[pos - 2 : pos] == b"/>" else None
        if head is not None and head.end() == pos:
            node.end = pos
        else:
            node.tail_start = pos
            node.end = data.index(b">", pos) + 1
        if keep is not None and node.parent is not None and not keep(node):
            node.parent.children.pop()

    def chars(text: str) -> None:
        if stack:
            stack[-1].text += text

    parser.StartElementHandler = start
    parser.EndElementHandler = end
    parser.CharacterDataHandler = chars
    parser.Parse(data, True)
    return roots[0]


def head_end(data: bytes, node: Node) -> int:
    """Offset just past the ``>`` that closes the start tag of ``node``."""
    match = _HEAD.match(data, node.start)
    if match is None:  # pragma: no cover - expat already accepted the tag
        raise ValueError(f"cannot find the start tag of <{node.local}>")
    return match.end()


def inner_xml(data: bytes, node: Node | None) -> str:
    """Raw XML between the start and end tag of ``node`` ("" for empty elements)."""
    if node is None or node.tail_start < 0:
        return ""
    return data[head_end(data, node) : node.tail_start].decode("utf-8")


def prefix_of(data: bytes, node: Node) -> str:
    """Namespace prefix of ``node`` as written in the file, e.g. ``"c:"`` or ``""``."""
    qname = re.match(rb"<([^\s/>]+)", data[node.start : node.start + 200])
    name = qname.group(1).decode() if qname else node.local
    return name.rpartition(":")[0] + ":" if ":" in name else ""


def escape_text(text: str) -> str:
    """Escape character data for XML."""
    return text.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")


def escape_attr(text: str) -> str:
    """Escape an attribute value for XML (double-quoted)."""
    return escape_text(text).replace('"', "&quot;")


class Editor:
    """Collects byte-range changes for one XML part and applies them in a single pass."""

    def __init__(self, data: bytes) -> None:
        self.data = data
        self._edits: list[tuple[int, int, bytes]] = []
        self._heads: dict[int, tuple[Node, str]] = {}

    @property
    def changed(self) -> bool:
        return bool(self._edits or self._heads)

    def replace(self, start: int, end: int, text: str) -> None:
        self._edits.append((start, end, text.encode("utf-8")))

    def remove(self, node: Node) -> None:
        self.replace(node.start, node.end, "")

    def _head(self, node: Node) -> str:
        if node.start in self._heads:
            return self._heads[node.start][1]
        return self.data[node.start : head_end(self.data, node)].decode("utf-8")

    def set_attr(self, node: Node, name: str, value: str | None) -> None:
        """Set (or with ``None`` remove) an unprefixed attribute on ``node``'s start tag."""
        head = self._head(node)
        match = re.search(r"(\s+)" + re.escape(name) + r"""\s*=\s*("[^"]*"|'[^']*')""", head)
        if value is None:
            new = head[: match.start()] + head[match.end() :] if match else head
        elif match:
            new = f'{head[: match.start()]}{match.group(1)}{name}="{escape_attr(value)}"{head[match.end() :]}'
        else:
            cut = len(head) - (2 if head.endswith("/>") else 1)
            new = f'{head[:cut].rstrip()} {name}="{escape_attr(value)}"{head[cut:]}'
        self._heads[node.start] = (node, new)

    def set_text(self, node: Node, text: str) -> None:
        """Replace the character content of a leaf element."""
        if node.tail_start < 0:
            head = self._head(node)
            self._heads.pop(node.start, None)
            qname = re.match(r"<([^\s/>]+)", head).group(1)  # type: ignore[union-attr]
            self.replace(node.start, node.end, f"{head[:-2].rstrip()}>{escape_text(text)}</{qname}>")
        else:
            self.replace(head_end(self.data, node), node.tail_start, escape_text(text))

    def apply(self) -> bytes:
        """Return the edited bytes. Edits nested inside a larger edit are dropped."""
        edits = self._edits + [(n.start, head_end(self.data, n), h.encode("utf-8")) for n, h in self._heads.values()]
        edits.sort(key=lambda e: (e[0], -e[1]))
        kept: list[tuple[int, int, bytes]] = []
        for start, end, text in edits:
            if kept and start < kept[-1][1]:
                if end <= kept[-1][1]:
                    continue
                raise ValueError("overlapping XML edits")
            kept.append((start, end, text))
        out, pos = [], 0
        for start, end, text in kept:
            out += [self.data[pos:start], text]
            pos = end
        out.append(self.data[pos:])
        return b"".join(out)


@dataclass(frozen=True)
class Rel:
    """One relationship from a ``.rels`` part."""

    id: str
    type: str
    target: str
    external: bool
    part: str | None  # resolved part name for internal targets

    @property
    def kind(self) -> str:
        """Last segment of the relationship type, e.g. ``worksheet`` or ``externalLinkPath``."""
        return self.type.rsplit("/", 1)[-1]


def rels_name(part: str) -> str:
    """Name of the ``.rels`` part that belongs to ``part`` ("" is the package root)."""
    folder, _, name = part.rpartition("/")
    return f"{folder}/_rels/{name}.rels" if folder else f"_rels/{name}.rels"


def resolve(base_part: str, target: str) -> str:
    """Resolve an internal relationship target relative to the part that owns it."""
    target = unquote(target)
    if target.startswith("/"):
        return target.lstrip("/")
    return posixpath.normpath(posixpath.join(posixpath.dirname(base_part), target)).lstrip("/")


class Package:
    """Read-only access to the parts of an OOXML zip package (names are case-insensitive)."""

    def __init__(self, path: str | Path) -> None:
        self.zip = zipfile.ZipFile(path)
        self._names = {info.filename.lower(): info.filename for info in self.zip.infolist()}
        self._rels: dict[str, dict[str, Rel]] = {}

    def __enter__(self) -> Package:
        return self

    def __exit__(self, *exc: object) -> None:
        self.close()

    def close(self) -> None:
        self.zip.close()

    def names(self) -> list[str]:
        return [info.filename for info in self.zip.infolist()]

    def has(self, part: str | None) -> bool:
        return part is not None and part.lower() in self._names

    def read(self, part: str) -> bytes:
        return self.zip.read(self._names[part.lower()])

    def xml(self, part: str, keep: Callable[[Node], bool] | None = None) -> Node:
        return parse_xml(self.read(part), keep)

    def rels(self, part: str) -> dict[str, Rel]:
        """Relationships of ``part`` keyed by id (empty if it has none)."""
        if part not in self._rels:
            found: dict[str, Rel] = {}
            name = rels_name(part)
            if self.has(name):
                for node in self.xml(name).findall("Relationship"):
                    target = node.get("Target") or ""
                    external = (node.get("TargetMode") or "").lower() == "external"
                    rid = node.get("Id") or ""
                    found[rid] = Rel(rid, node.get("Type") or "", target, external, None if external else resolve(part, target))
            self._rels[part] = found
        return self._rels[part]

    def content_type(self, part: str) -> str:
        """Content type of ``part`` from ``[Content_Types].xml``."""
        types = self.xml("[Content_Types].xml")
        for node in types.findall("Override"):
            if (node.get("PartName") or "").lstrip("/").lower() == part.lower():
                return node.get("ContentType") or ""
        ext = part.rpartition(".")[2].lower()
        for node in types.findall("Default"):
            if (node.get("Extension") or "").lower() == ext:
                return node.get("ContentType") or ""
        return ""


def _rel_attr_values(data: bytes) -> tuple[set[str], str | None]:
    """Values of all ``r:*`` attributes in a part, plus a parse error message if it is not XML."""
    values: set[str] = set()
    parser = expat.ParserCreate(namespace_separator=" ")

    def start(_name: str, attrs: dict[str, str]) -> None:
        values.update(v for k, v in attrs.items() if k.split(" ")[0].endswith("/relationships") and " " in k)

    parser.StartElementHandler = start
    try:
        parser.Parse(data, True)
    except expat.ExpatError as exc:
        return values, str(exc)
    return values, None


def check_package(path: str | Path) -> list[str]:
    """Check the structure Excel relies on; returns a list of problems (empty = consistent).

    * every part has a content type and every content-type override names an existing part,
    * every internal relationship target exists,
    * every XML part is well-formed and every ``r:id``-style attribute names a relationship
      that exists in the part's ``.rels``.
    """
    problems: list[str] = []
    with zipfile.ZipFile(path) as zf:
        names = [i.filename for i in zf.infolist() if not i.filename.endswith("/")]
        lower = {n.lower() for n in names}
        if "[content_types].xml" not in lower:
            return ["missing [Content_Types].xml"]
        types = parse_xml(zf.read(next(n for n in names if n.lower() == "[content_types].xml")))
        defaults = {(d.get("Extension") or "").lower() for d in types.findall("Default")}
        overrides = {(o.get("PartName") or "").lstrip("/").lower() for o in types.findall("Override")}
        problems += [f"content type override for missing part /{o}" for o in sorted(overrides - lower)]
        rel_ids: dict[str, set[str]] = {}
        for name in names:
            low = name.lower()
            if low == "[content_types].xml":
                continue
            if low not in overrides and low.rpartition(".")[2] not in defaults:
                problems.append(f"no content type for {name}")
            if not low.endswith(".rels"):
                continue
            folder, _, file = name.rpartition("/")
            source = posixpath.join(posixpath.dirname(folder), file[: -len(".rels")]) if folder != "_rels" else ""
            try:
                rels = parse_xml(zf.read(name))
            except expat.ExpatError as exc:
                problems.append(f"{name}: not well-formed XML ({exc})")
                continue
            rel_ids[source.lower()] = set()
            for rel in rels.findall("Relationship"):
                rel_ids[source.lower()].add(rel.get("Id") or "")
                if (rel.get("TargetMode") or "").lower() != "external":
                    target = resolve(source, rel.get("Target") or "")
                    if target.lower() not in lower:
                        problems.append(f"{name}: relationship {rel.get('Id')} points to missing part {target}")
        for name in names:
            low = name.lower()
            if not low.endswith(".xml") or low == "[content_types].xml":
                continue
            used, error = _rel_attr_values(zf.read(name))
            if error:
                problems.append(f"{name}: not well-formed XML ({error})")
            missing = sorted(used - rel_ids.get(low, set()) - {""})
            problems += [f"{name}: r:id {rid} has no relationship" for rid in missing]
    return problems
