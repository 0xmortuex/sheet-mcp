"""Formula analysis built on openpyxl's tokenizer: sheet references, renaming,
R1C1 patterns, and the check_formulas engine. Nothing here evaluates formulas."""

from __future__ import annotations

import bisect
import re
from collections import Counter, defaultdict
from collections.abc import Iterable
from dataclasses import dataclass, field

from openpyxl.formula import Tokenizer
from openpyxl.utils import get_column_letter
from openpyxl.utils.exceptions import InvalidFileException

from .errors import SheetError

_PREFIX = re.compile(r"^(?P<sheet>'(?:[^']|'')+'|[^'!\[\]:]+(?::[^'!\[\]:]+)?)!(?P<rest>.+)$")
_CELL = re.compile(r"^(\$?)([A-Za-z]{1,3})(\$?)(\d+)$")
_COL = re.compile(r"^(\$?)([A-Za-z]{1,3})$")
_ROW = re.compile(r"^(\$?)(\d+)$")
_PLAIN_NAME = re.compile(r"^[A-Za-z_][A-Za-z0-9_.]*$")
_CELLISH = re.compile(r"^(?:[A-Za-z]{1,3}\d+|[Rr]\d*[Cc]\d*|[Rr]|[Cc])$")

VOLATILE = frozenset("NOW TODAY RAND RANDBETWEEN RANDARRAY OFFSET INDIRECT INFO CELL".split())
# Literals nobody calls "magic": counting, units, percent, time, months.
COMMON_CONSTANTS = frozenset(float(n) for n in (*range(11), 12, 24, 60, 100, 1000))
AGGREGATES = frozenset("SUM SUBTOTAL AVERAGE COUNT COUNTA MIN MAX AGGREGATE".split())


def _tokens(formula: str) -> list[tuple[str, str, str]]:
    f = formula if formula.startswith("=") else "=" + formula
    try:
        return [(t.value, t.type, t.subtype) for t in Tokenizer(f).items]
    except (IndexError, InvalidFileException, ValueError) as e:
        raise SheetError(f"cannot parse formula {formula!r}: {e}") from e


def _split_prefix(operand: str) -> tuple[list[str] | None, bool, str]:
    """(sheet names in the prefix, was_quoted, rest) or (None, False, operand)."""
    m = _PREFIX.match(operand)
    if not m:
        return None, False, operand
    sheet = m.group("sheet")
    quoted = sheet.startswith("'")
    if quoted:
        sheet = sheet[1:-1].replace("''", "'")
    return sheet.split(":"), quoted, m.group("rest")


def sheet_refs(formula: str) -> set[str]:
    out: set[str] = set()
    for value, typ, sub in _tokens(formula):
        if typ == "OPERAND" and sub == "RANGE":
            names, _q, _rest = _split_prefix(value)
            if names:
                out.update(names)
    return out


def functions_used(formula: str) -> list[str]:
    return [v[:-1] for v, typ, sub in _tokens(formula) if typ == "FUNC" and sub == "OPEN" and v != "("]


def quote_sheet(name: str) -> str:
    if _PLAIN_NAME.match(name) and not _CELLISH.match(name) and "." not in name[:1]:
        return name
    return "'" + name.replace("'", "''") + "'"


def _tok_render(toks: list[tuple[str, str, str]], original: str) -> str:
    return ("=" if original.startswith("=") else "") + "".join(v for v, _t, _s in toks)


def rename_in_formula(formula: str, old: str, new: str) -> str:
    """Rename sheet references old -> new in reference tokens only (never string literals)."""
    toks = _tokens(formula)
    changed = False
    out: list[tuple[str, str, str]] = []
    old_f = old.casefold()
    for value, typ, sub in toks:
        if typ == "OPERAND" and sub == "RANGE":
            names, _q, rest = _split_prefix(value)
            if names and any(n.casefold() == old_f for n in names):
                mapped = [new if n.casefold() == old_f else n for n in names]
                joined = ":".join(mapped)
                need = any(quote_sheet(n) != n for n in mapped)
                prefix = "'" + joined.replace("'", "''") + "'" if need else joined
                value = f"{prefix}!{rest}"
                changed = True
        out.append((value, typ, sub))
    return _tok_render(out, formula) if changed else formula


# ------------------------------------------------------------------ R1C1


def col_num(letters: str) -> int:
    n = 0
    for ch in letters.upper():
        n = n * 26 + (ord(ch) - 64)
    return n


def _rc(abs_: str, target: int, cur: int, letter: str) -> str:
    if abs_:
        return f"{letter}{target}"
    d = target - cur
    return letter if d == 0 else f"{letter}[{d}]"


def _piece_r1c1(piece: str, row: int, col: int) -> str | None:
    m = _CELL.match(piece)
    if m:
        return _rc(m.group(3), int(m.group(4)), row, "R") + _rc(m.group(1), col_num(m.group(2)), col, "C")
    m = _COL.match(piece)
    if m:
        return _rc(m.group(1), col_num(m.group(2)), col, "C")
    m = _ROW.match(piece)
    if m:
        return _rc(m.group(1), int(m.group(2)), row, "R")
    return None


def r1c1(formula: str, row: int, col: int) -> str:
    """Position-independent form of a formula, e.g. B2 '=A2*2' -> '=RC[-1]*2'."""
    out = ["="]
    for value, typ, sub in _tokens(formula):
        if typ == "WHITE-SPACE":
            continue
        if typ == "OPERAND" and sub == "RANGE":
            names, _q, rest = _split_prefix(value)
            prefix = value[: len(value) - len(rest)] if names else ""
            pieces = rest.split(":")
            conv = [_piece_r1c1(p, row, col) for p in pieces]
            if all(c is not None for c in conv):
                value = prefix + ":".join(c for c in conv if c is not None)
        out.append(value)
    return "".join(out)


def parse_ref(rest: str) -> tuple[int, int, int, int] | None:
    """A1, A1:B2, A:A, 1:3 -> (min_row, min_col, max_row, max_col) with open ends clamped."""
    pieces = rest.replace("$", "").split(":")
    if len(pieces) > 2:
        return None
    rows: list[int] = []
    cols: list[int] = []
    kinds = set()
    for p in pieces:
        m = _CELL.match(p)
        if m:
            cols.append(col_num(m.group(2)))
            rows.append(int(m.group(4)))
            kinds.add("cell")
            continue
        m = _COL.match(p)
        if m:
            cols.append(col_num(m.group(2)))
            kinds.add("col")
            continue
        m = _ROW.match(p)
        if m:
            rows.append(int(m.group(2)))
            kinds.add("row")
            continue
        return None
    if len(kinds) != 1:
        return None
    kind = kinds.pop()
    if kind == "cell":
        return min(rows), min(cols), max(rows), max(cols)
    if kind == "col":
        return 1, min(cols), 1048576, max(cols)
    return min(rows), 1, max(rows), 16384


def is_ref_like(piece: str) -> bool:
    return _CELL.match(piece) is not None or _COL.match(piece) is not None or _ROW.match(piece) is not None


# ------------------------------------------------------------------ analysis


@dataclass
class FormulaCell:
    sheet: str
    row: int
    col: int
    text: str

    @property
    def addr(self) -> str:
        return f"{get_column_letter(self.col)}{self.row}"


@dataclass
class Issue:
    kind: str
    where: str
    message: str
    group: str = ""


@dataclass
class Findings:
    issues: dict[str, list[Issue]] = field(default_factory=dict)
    notes: list[str] = field(default_factory=list)

    def add(self, kind: str, where: str, message: str, group: str = "") -> None:
        self.issues.setdefault(kind, []).append(Issue(kind, where, message, group))


def _is_infix_arith(tok: tuple[str, str, str] | None) -> bool:
    return tok is not None and tok[1] == "OPERATOR-INFIX" and tok[0] in ("+", "-", "*", "/", "^")


def analyze_one(fc: FormulaCell, names: set[str], tables: set[str], sheets: set[str], out: Findings) -> tuple[str, list[tuple[str, tuple[int, int, int, int]]]]:
    """Per-formula checks. Returns (r1c1 pattern, [(sheet, bounds)] dependencies)."""
    toks = _tokens(fc.text)
    sig = [t for t in toks if t[1] != "WHITE-SPACE"]
    deps: list[tuple[str, tuple[int, int, int, int]]] = []
    where = f"{fc.sheet}!{fc.addr}"
    pattern = r1c1(fc.text, fc.row, fc.col)
    group = f"{fc.sheet}|{pattern}"
    sheets_f = {s.casefold(): s for s in sheets}
    for i, tok in enumerate(sig):
        value, typ, sub = tok
        if typ == "OPERAND" and sub == "ERROR" and value.upper() == "#REF!":
            out.add("broken_reference", where, f"{fc.text} contains #REF!")
        elif typ == "OPERAND" and sub == "RANGE":
            pnames, _q, rest = _split_prefix(value)
            target = fc.sheet
            if pnames:
                missing = [n for n in pnames if n.casefold() not in sheets_f]
                if missing:
                    out.add(
                        "broken_reference",
                        where,
                        f"{fc.text} refers to missing sheet {missing[0]!r}",
                    )
                    continue
                if len(pnames) == 1:
                    target = sheets_f[pnames[0].casefold()]
                else:
                    continue
            if "[" in rest:
                tname = rest.split("[", 1)[0]
                if tname and tname.casefold() not in {t.casefold() for t in tables}:
                    out.add("unknown_name", where, f"{fc.text} uses structured reference to unknown table {tname!r}")
                continue
            bounds = parse_ref(rest)
            if bounds is not None:
                deps.append((target, bounds))
            elif "#REF" in rest.upper():
                out.add("broken_reference", where, f"{fc.text} contains #REF!")
            elif all(is_ref_like(p) for p in rest.split(":")):
                pass
            elif not pnames and rest.casefold() not in {n.casefold() for n in names}:
                if _PLAIN_NAME.match(rest) or rest.startswith("\\"):
                    out.add("unknown_name", where, f"{fc.text} uses name {rest!r}, which is not defined (Excel shows #NAME?)")
        elif typ == "OPERAND" and sub == "NUMBER":
            try:
                num = float(value)
            except ValueError:
                continue
            prev = sig[i - 1] if i > 0 else None
            nxt = sig[i + 1] if i + 1 < len(sig) else None
            if abs(num) not in COMMON_CONSTANTS and (_is_infix_arith(prev) or _is_infix_arith(nxt)):
                out.add("hardcoded_number", where, f"{fc.text} has literal {value} inside the arithmetic", group)
        elif typ == "FUNC" and sub == "OPEN" and value != "(":
            fn = value[:-1].upper()
            if fn.startswith("_XLFN."):
                fn = fn[6:]
            if fn in VOLATILE:
                out.add("volatile", where, f"{fn}() recalculates on every change: {fc.text}", f"{group}|{fn}")
    return pattern, deps


class ColumnIndex:
    """Formula cells of one sheet indexed by column for fast range lookups."""

    def __init__(self) -> None:
        self.cols: dict[int, list[int]] = defaultdict(list)

    def add(self, row: int, col: int) -> None:
        self.cols[col].append(row)

    def freeze(self) -> None:
        for lst in self.cols.values():
            lst.sort()

    def in_bounds(self, b: tuple[int, int, int, int]) -> Iterable[tuple[int, int]]:
        r0, c0, r1, c1 = b
        for c in range(c0, min(c1, 16384) + 1):
            lst = self.cols.get(c)
            if not lst:
                continue
            lo = bisect.bisect_left(lst, r0)
            hi = bisect.bisect_right(lst, r1)
            for r in lst[lo:hi]:
                yield r, c


EDGE_BUDGET = 3_000_000


def find_cycles(
    deps: dict[tuple[str, int, int], list[tuple[str, tuple[int, int, int, int]]]],
    index: dict[str, ColumnIndex],
) -> tuple[list[list[tuple[str, int, int]]], bool]:
    """Strongly connected components with a cycle. Returns (cycles, complete)."""
    graph: dict[tuple[str, int, int], list[tuple[str, int, int]]] = {}
    edges = 0
    for node, dl in deps.items():
        out: list[tuple[str, int, int]] = []
        for sheet, b in dl:
            ci = index.get(sheet)
            if ci is None:
                continue
            for r, c in ci.in_bounds(b):
                out.append((sheet, r, c))
                edges += 1
                if edges > EDGE_BUDGET:
                    return [], False
        graph[node] = out
    idx: dict[tuple[str, int, int], int] = {}
    low: dict[tuple[str, int, int], int] = {}
    on: set[tuple[str, int, int]] = set()
    stack: list[tuple[str, int, int]] = []
    cycles: list[list[tuple[str, int, int]]] = []
    counter = 0
    for root in graph:
        if root in idx:
            continue
        work: list[tuple[tuple[str, int, int], int]] = [(root, 0)]
        while work:
            node, i = work.pop()
            if i == 0:
                idx[node] = low[node] = counter
                counter += 1
                stack.append(node)
                on.add(node)
            nbrs = graph.get(node, [])
            descended = False
            while i < len(nbrs):
                w = nbrs[i]
                i += 1
                if w not in idx:
                    work.append((node, i))
                    work.append((w, 0))
                    descended = True
                    break
                if w in on:
                    low[node] = min(low[node], idx[w])
            if descended:
                continue
            if low[node] == idx[node]:
                comp = []
                while True:
                    w = stack.pop()
                    on.discard(w)
                    comp.append(w)
                    if w == node:
                        break
                if len(comp) > 1 or node in graph.get(node, []):
                    cycles.append(sorted(comp))
            if work:
                parent = work[-1][0]
                low[parent] = min(low[parent], low[node])
    return cycles, True


def pattern_outliers(
    by_col: dict[int, dict[int, str]],
    constants: dict[int, set[int]],
    sheet: str,
    texts: dict[tuple[int, int], str],
    out: Findings,
) -> None:
    """Formulas (or hardcoded numbers) that break the pattern both neighbours share."""
    for col, rows in by_col.items():
        if len(rows) < 3:
            continue
        counts = Counter(rows.values())
        for r, pat in rows.items():
            up, down = rows.get(r - 1), rows.get(r + 1)
            if up is not None and up == down and up != pat and counts[up] >= 2:
                addr = f"{get_column_letter(col)}{r}"
                out.add(
                    "inconsistent_formula",
                    f"{sheet}!{addr}",
                    f"{texts[(r, col)]} differs from the formulas directly above and below "
                    f"(they share the R1C1 pattern {up})",
                )
        for r in constants.get(col, ()):
            up, down = rows.get(r - 1), rows.get(r + 1)
            if up is not None and up == down:
                addr = f"{get_column_letter(col)}{r}"
                out.add(
                    "hardcoded_value_in_formula_column",
                    f"{sheet}!{addr}",
                    "a typed number sits between formulas that share one pattern "
                    f"({up}); it will not follow changes to its inputs",
                )
