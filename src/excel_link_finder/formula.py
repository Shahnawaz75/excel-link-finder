"""Find references to other workbooks inside Excel formulas, as they are stored in .xlsx files.

In the file, Excel writes an external reference with the number of the external link
instead of the file name: ``[1]Prices!$B$2`` means "sheet Prices of external link 1".
``[0]`` means *this* workbook. Hand-made files may also contain the full form
``'C:\\Data\\[Budget.xlsx]Plan'!A1`` or ``Budget.xlsx!SomeName``. Brackets are also used
by table (structured) references such as ``Orders[[#This Row],[Qty]]``, which are not
links, so a small tokenizer is used instead of a plain regular expression.
"""

from __future__ import annotations

import re
from dataclasses import dataclass

_IDENT = frozenset("ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789_.\\?$")
_ERROR = re.compile(
    r"#(?:NULL!|DIV/0!|VALUE!|REF!|NAME\?|NUM!|N/A|GETTING_DATA|SPILL!|CALC!|FIELD!|BLOCKED!|CONNECT!|BUSY!|UNKNOWN!)", re.I
)
_BOOK = re.compile(r"\.(?:xl[a-z]{1,2}|csv)$", re.I)  # .xls .xlsx .xlsm .xlsb .xla .xlam .xltx .xltm .csv
_CELL = re.compile(r"^\$?[A-Za-z]{1,3}\$?\d+$")
_QUOTED_BOOK = re.compile(r"^(?P<path>[^\[]*)\[(?P<book>[^\]]+)\](?P<sheet>.*)$", re.S)

Token = tuple[str, int, int]  # (kind, start, end)


@dataclass(frozen=True)
class ExternalRef:
    """One reference to another workbook found in a formula."""

    index: int | None  # external link number [N]; None when the workbook is written by name or path
    book: str  # workbook file name for name/path references, "" for [N] references
    path: str  # folder written in front of the book name, if any
    start: int  # span of the reference text in the formula ("[N]" for numbered references)
    end: int

    @property
    def target(self) -> str:
        """File the reference points to, for name/path references."""
        return self.path + self.book


def _is_ident(ch: str) -> bool:
    return ch in _IDENT or ord(ch) > 127


def tokenize(formula: str) -> list[Token]:
    """Split a formula into strings, quoted names, bracket groups, identifiers and operators."""
    tokens: list[Token] = []
    i, n = 0, len(formula)
    while i < n:
        ch = formula[i]
        if ch in "\"'":
            j = i + 1
            while j < n:
                if formula[j] == ch:
                    if j + 1 < n and formula[j + 1] == ch:
                        j += 2
                        continue
                    break
                j += 1
            tokens.append(("str" if ch == '"' else "quoted", i, min(j + 1, n)))
            i = j + 1
        elif ch == "[":
            depth, j = 0, i
            while j < n:
                c = formula[j]
                if c == "'":  # escape character inside structured references
                    j += 2
                    continue
                if c == "[":
                    depth += 1
                elif c == "]":
                    depth -= 1
                    if depth == 0:
                        break
                j += 1
            tokens.append(("bracket", i, min(j + 1, n)))
            i = j + 1
        elif ch == "#" and (m := _ERROR.match(formula, i)):
            tokens.append(("error", i, m.end()))
            i = m.end()
        elif _is_ident(ch):
            j = i
            while j < n and _is_ident(formula[j]):
                j += 1
            tokens.append(("ident", i, j))
            i = j
        else:
            if not ch.isspace():
                tokens.append(("op", i, i + 1))
            i += 1
    return tokens


def _text(formula: str, token: Token | None) -> str:
    return formula[token[1] : token[2]] if token else ""


def external_refs(formula: str) -> list[ExternalRef]:
    """Return every reference to another workbook in ``formula``."""
    tokens = tokenize(formula)
    refs: list[ExternalRef] = []
    for i, tok in enumerate(tokens):
        kind, start, end = tok
        prev = tokens[i - 1] if i else None
        nxt = tokens[i + 1] if i + 1 < len(tokens) else None
        if kind == "bracket":
            if prev and prev[0] in ("ident", "bracket") and prev[2] == start:
                continue  # Table1[Column]: a structured reference
            inner = formula[start + 1 : end - 1]
            if inner.isdigit():
                if int(inner) > 0:  # [0] is the workbook itself
                    refs.append(ExternalRef(int(inner), "", "", start, end))
            elif _BOOK.search(inner) and nxt and nxt[0] in ("ident", "op"):
                refs.append(ExternalRef(None, inner, "", start, end))  # [Book.xlsx]Sheet1!A1
        elif kind == "quoted" and _text(formula, nxt) == "!":
            raw = formula[start + 1 : end - 1]
            text = raw.replace("''", "'")
            match = _QUOTED_BOOK.match(text)
            if match:
                book = match.group("book")
                if not book.isdigit():
                    refs.append(ExternalRef(None, book, match.group("path"), start, end))
                elif int(book) > 0:
                    pos = start + 1 + raw.index("[")
                    refs.append(ExternalRef(int(book), "", match.group("path"), pos, pos + len(book) + 2))
            elif _BOOK.search(text) or "\\" in text or "/" in text:  # 'C:\Data\Book.xlsx'!Name
                folder, sep, book = text.replace("/", "\\").rpartition("\\")
                refs.append(ExternalRef(None, book, folder + sep, start, end))
        elif kind == "ident" and _text(formula, nxt) == "!" and _BOOK.search(formula[start:end]):
            refs.append(ExternalRef(None, formula[start:end], "", start, end))  # Book.xlsx!Name
    return refs


def name_refs(formula: str) -> list[tuple[str | None, str]]:
    """Identifiers that may be defined names, as ``(sheet or None, name)`` in lower case.

    A sheet is given for sheet-qualified names such as ``Report!LocalName``. Names that
    belong to another workbook (``[1]!Name``) are left out: they are external references.
    """
    tokens = tokenize(formula)
    found: list[tuple[str | None, str]] = []
    for i, (kind, start, end) in enumerate(tokens):
        if kind != "ident":
            continue
        word = formula[start:end]
        if not (word[0].isalpha() or word[0] in "_\\") or _CELL.match(word) or word.upper() in ("TRUE", "FALSE"):
            continue
        nxt = tokens[i + 1] if i + 1 < len(tokens) else None
        if _text(formula, nxt) in ("(", "!") or (nxt and nxt[0] == "bracket" and nxt[1] == end):
            continue  # function call, sheet name or table name
        sheet: str | None = None
        if i >= 2 and _text(formula, tokens[i - 1]) == "!":
            qualifier = tokens[i - 2]
            sheet = formula[qualifier[1] : qualifier[2]]
            if qualifier[0] == "quoted":
                sheet = sheet[1:-1].replace("''", "'")
            before = tokens[i - 3] if i >= 3 else None
            if qualifier[0] == "bracket" or "[" in sheet or (before and before[0] == "bracket" and before[2] == qualifier[1]):
                continue
        found.append((sheet.lower() if sheet else None, word.lower()))
    return found


def renumber(formula: str, mapping: dict[int, int]) -> str:
    """Rewrite ``[N]`` external link numbers using ``mapping`` (old number -> new number)."""
    out, pos = [], 0
    for ref in external_refs(formula):
        if ref.index is not None and mapping.get(ref.index, ref.index) != ref.index:
            out += [formula[pos : ref.start], f"[{mapping[ref.index]}]"]
            pos = ref.end
    out.append(formula[pos:])
    return "".join(out)
