"""Cell values: how they are shown to the agent, and how agent input is coerced."""

from __future__ import annotations

import datetime as dt
import math
import re
from dataclasses import dataclass, field
from typing import Any

from openpyxl.cell.cell import ILLEGAL_CHARACTERS_RE
from openpyxl.styles.numbers import is_date_format

from .errors import SheetError

MAX_TEXT = 32767
MAX_FORMULA = 8192
_ISO_DATE = re.compile(r"^\d{4}-\d{2}-\d{2}$")
_ISO_DATETIME = re.compile(r"^\d{4}-\d{2}-\d{2}[T ]\d{2}:\d{2}(:\d{2}(\.\d{1,6})?)?$")
_QUOTED = re.compile(r'"[^"]*"|\\.')

# Functions Excel stores with an _xlfn. prefix; written bare they show #NAME?.
XLFN_FUNCS = frozenset(
    "IFS SWITCH MAXIFS MINIFS CONCAT TEXTJOIN XLOOKUP XMATCH IFNA XOR DAYS ISOWEEKNUM "
    "STDEV.S STDEV.P VAR.S VAR.P".split()
)
# Dynamic-array functions: openpyxl cannot write the array metadata they need.
DYNAMIC_FUNCS = frozenset("FILTER SORT SORTBY UNIQUE SEQUENCE RANDARRAY LET LAMBDA".split())


@dataclass
class CellView:
    coord: str
    value: Any = None
    formula: str | None = None
    fmt: str = "General"
    dtype: str = "n"  # openpyxl data_type: n s b d e f
    no_cache: bool = False
    cached_error: bool = False  # a formula whose cached result is an error value

    @property
    def is_error(self) -> bool:
        return self.dtype == "e" or self.cached_error


def formula_text(v: Any) -> str:
    """Formula text of an openpyxl cell value that has data_type 'f'."""
    if isinstance(v, str):
        return str(v)
    text = getattr(v, "text", None)
    ref = getattr(v, "ref", None)
    if text is not None:  # ArrayFormula
        body = text if text.startswith("=") else "=" + text
        return str("{" + body + "}" + (f" [array {ref}]" if ref else ""))
    if hasattr(v, "r1"):  # DataTableFormula
        return f"{{=TABLE(row={getattr(v, 'r1', None)}, col={getattr(v, 'r2', None)})}} [data table {ref}]"
    return str(v)


def has_time_part(fmt: str) -> bool:
    f = _QUOTED.sub("", fmt).lower()
    return "h" in f or "s" in f


def iso(v: Any, fmt: str = "General") -> str:
    if isinstance(v, dt.datetime):
        if v.time() == dt.time(0) and not has_time_part(fmt):
            return v.date().isoformat()
        return v.isoformat()
    if isinstance(v, (dt.date, dt.time)):
        return v.isoformat()
    if isinstance(v, dt.timedelta):
        total = int(v.total_seconds())
        sign = "-" if total < 0 else ""
        total = abs(total)
        return f"{sign}{total // 3600}:{total % 3600 // 60:02d}:{total % 60:02d}"
    raise TypeError(type(v))


def display(v: Any, fmt: str = "General") -> str:
    if v is None:
        return ""
    if isinstance(v, bool):
        return "TRUE" if v else "FALSE"
    if isinstance(v, (dt.datetime, dt.date, dt.time, dt.timedelta)):
        return iso(v, fmt)
    return str(v)


def jsonable(v: Any, fmt: str = "General") -> Any:
    if isinstance(v, (dt.datetime, dt.date, dt.time, dt.timedelta)):
        return iso(v, fmt)
    if isinstance(v, float) and not math.isfinite(v):
        return str(v)
    return v


_TURKISH = str.maketrans(
    "İıüÜöÖşŞçÇğĞ", "iiuuoosscc" + "gg"
)


def fold(s: str) -> str:
    """Case- and Turkish-diacritic-insensitive key: I, i, dotless i and dotted I all fold to
    i, and u/o/s/c/g with their Turkish marks fold to the plain letters."""
    return s.translate(_TURKISH).casefold().replace("i̇", "i")


# ---------------------------------------------------------------- coercion


@dataclass
class Coerced:
    kind: str  # formula | text | number | bool | date | datetime | empty
    value: Any
    warnings: list[str] = field(default_factory=list)

    def describe(self) -> str:
        if self.kind == "empty":
            return "(cleared)"
        if self.kind == "formula":
            return str(self.value)
        if self.kind == "text":
            return f"text {self.value!r}"
        if self.kind == "bool":
            return "TRUE" if self.value else "FALSE"
        if self.kind in ("date", "datetime"):
            return f"{self.kind} {self.value.isoformat()}"
        return f"number {self.value!r}"


def check_formula(text: str, sheet_names: set[str]) -> list[str]:
    """Reject formulas that would corrupt the sheet; return soft warnings."""
    from .formulas import functions_used, sheet_refs

    if len(text) > MAX_FORMULA:
        raise SheetError(f"formula is {len(text)} characters long; Excel's limit is {MAX_FORMULA}")
    if any(ord(c) < 32 and c not in "\t\n" for c in text):
        raise SheetError("formula contains control characters")
    depth = 0
    in_str = False
    for ch in text[1:]:
        if ch == '"':
            in_str = not in_str
        elif not in_str:
            if ch == "(":
                depth += 1
            elif ch == ")":
                depth -= 1
                if depth < 0:
                    break
    if in_str or depth != 0:
        raise SheetError(
            f"formula {text!r} has unbalanced parentheses or quotes. "
            "To store text that starts with '=', write it as \"'=...\" (leading apostrophe)."
        )
    bare = re.sub(r'"[^"]*"', '""', text)
    if re.search(r"\|\s*'", bare) or re.search(
        r"(?<![A-Za-z0-9_.])(?:CALL|REGISTER|REGISTER\.ID|EXEC)\s*\(", bare, re.IGNORECASE
    ):
        raise SheetError(
            f"formula {text!r} looks like a DDE / external-code call (app|'command'!topic, CALL, REGISTER, EXEC), "
            "which can run programs when the workbook is opened. Refused."
        )
    warnings = []
    folded = {n.casefold() for n in sheet_names}
    for ref in sorted(sheet_refs(text)):
        if ref.casefold() not in folded:
            raise SheetError(
                f"formula {text!r} refers to sheet {ref!r}, which does not exist "
                f"(sheets: {', '.join(sorted(sheet_names))}). Excel would show #REF!."
            )
    for fn in functions_used(text):
        up = fn.upper()
        if up in ("WEBSERVICE", "RTD", "FILTERXML"):
            warnings.append(
                f"{up} fetches data from outside the workbook when it is opened; make sure that is intended"
            )
        elif up in XLFN_FUNCS:
            warnings.append(
                f"{up} is a newer function; write it as _xlfn.{up}(...) or Excel shows #NAME?"
            )
        elif up in DYNAMIC_FUNCS:
            warnings.append(
                f"{up} is a dynamic-array function; openpyxl cannot write the array metadata "
                "it needs, so Excel will treat it as a legacy formula (implicit intersection)"
            )
    return warnings


def coerce(raw: Any, sheet_names: set[str]) -> Coerced:
    """Turn JSON input into a cell value using the documented rules:

    null -> clear; bool -> boolean; int/float -> number; str starting with '=' -> formula;
    str starting with a single quote -> literal text (quote removed); ISO date or datetime
    strings (YYYY-MM-DD, YYYY-MM-DDTHH:MM[:SS]) -> real dates; any other str -> text.
    """
    if raw is None:
        return Coerced("empty", None)
    if isinstance(raw, bool):
        return Coerced("bool", raw)
    if isinstance(raw, (int, float)):
        if isinstance(raw, float) and not math.isfinite(raw):
            raise SheetError("NaN and infinity cannot be stored in a cell")
        warnings = []
        if isinstance(raw, int) and abs(raw) >= 10**15:
            warnings.append(
                "Excel keeps 15 significant digits, so this integer will lose precision; "
                "pass it as text (leading apostrophe) if it is an identifier"
            )
        return Coerced("number", raw, warnings)
    if not isinstance(raw, str):
        raise SheetError(
            f"unsupported value {raw!r} ({type(raw).__name__}); use string, number, boolean or null"
        )
    if len(raw) > MAX_TEXT:
        raise SheetError(f"text is {len(raw)} characters; Excel's cell limit is {MAX_TEXT}")
    if ILLEGAL_CHARACTERS_RE.search(raw):
        raise SheetError(
            "text contains control characters that Excel cannot store (e.g. \x00-\x08, \x0b, \x0c, \x0e-\x1f)"
        )
    if raw.startswith("="):
        if len(raw) == 1:
            raise SheetError("a formula needs something after '='")
        return Coerced("formula", raw, check_formula(raw, sheet_names))
    if raw.startswith("'"):
        return Coerced("text", raw[1:])
    if _ISO_DATETIME.match(raw):
        try:
            return Coerced("datetime", dt.datetime.fromisoformat(raw))
        except ValueError:
            pass
    elif _ISO_DATE.match(raw):
        try:
            return Coerced("date", dt.date.fromisoformat(raw))
        except ValueError:
            pass
    return Coerced("text", raw)


def apply(cell: Any, c: Coerced) -> None:
    """Write a coerced value, keeping the cell's existing style where possible."""
    if c.kind == "empty":
        cell.value = None
    elif c.kind == "formula":
        cell.value = c.value
    elif c.kind == "text":
        cell.value = c.value
        cell.data_type = "s"  # force: '=...' and '#REF!' must stay literal text
    elif c.kind in ("date", "datetime"):
        cell.value = c.value
        if not is_date_format(cell.number_format):
            cell.number_format = "yyyy-mm-dd" if c.kind == "date" else "yyyy-mm-dd hh:mm:ss"
    else:
        cell.value = c.value


def before_text(v: Any, dtype: str) -> str:
    if v is None:
        return "(empty)"
    if dtype == "f":
        return formula_text(v)
    return display(v)
