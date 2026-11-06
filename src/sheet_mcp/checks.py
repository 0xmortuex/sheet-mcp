"""check_formulas: find formulas that are probably wrong without calculating anything."""

from __future__ import annotations

from collections import defaultdict
from contextlib import closing

from openpyxl.utils import get_column_letter

from . import reader
from .errors import SheetError
from .formulas import (
    ColumnIndex,
    Findings,
    FormulaCell,
    Issue,
    analyze_one,
    find_cycles,
    pattern_outliers,
)
from .package import PackageInfo

TITLES = {
    "broken_reference": "Broken references (#REF! or a missing sheet)",
    "unknown_name": "Unknown names (Excel shows #NAME?)",
    "cached_error": "Cells currently showing an error value",
    "circular": "Circular references",
    "inconsistent_formula": "Formulas that break the column pattern",
    "hardcoded_value_in_formula_column": "Typed numbers inside a column of formulas",
    "hardcoded_number": "Hardcoded numbers inside formulas",
    "volatile": "Volatile functions",
}


def check(info: PackageInfo, sheet: str | None, max_issues: int, max_formulas: int = 100_000) -> str:
    if not 1 <= max_issues <= 500:
        raise SheetError("max_issues must be between 1 and 500")
    targets = [info.worksheet(sheet)] if sheet else [s for s in info.sheets if s.kind == "worksheet"]
    all_sheets = {s.name for s in info.sheets}
    names = {n.name for n in info.names}
    tables = {t.name for s in info.sheets for t in s.tables}
    out = Findings()
    index: dict[str, ColumnIndex] = {}
    deps_map: dict[tuple[str, int, int], list[tuple[str, tuple[int, int, int, int]]]] = {}
    analyzed = 0
    skipped = 0
    unparsable = 0
    for sc in targets:
        if sc.n_formulas == 0 or sc.max_row == 0:
            continue
        ci = index.setdefault(sc.name, ColumnIndex())
        by_col: dict[int, dict[int, str]] = defaultdict(dict)
        consts: dict[int, set[int]] = defaultdict(set)
        texts: dict[tuple[int, int], str] = {}
        bounds: reader.Bounds = (1, 1, sc.max_row, sc.max_col)
        with closing(reader.rows(info, sc.name, bounds)) as rows:
            for r_idx, row in enumerate(rows):
                r = r_idx + 1
                for k, c in enumerate(row):
                    col = k + 1
                    if c.formula is None:
                        v = c.value
                        if isinstance(v, (int, float)) and not isinstance(v, bool) and c.dtype == "n":
                            consts[col].add(r)
                        continue
                    if analyzed >= max_formulas:
                        skipped += 1
                        continue
                    text = c.formula
                    if text.startswith("{"):
                        text = text[1 : text.index("}")] if "}" in text else text[1:]
                    if "TABLE(row=" in text:
                        continue
                    fc = FormulaCell(sc.name, r, col, text)
                    try:
                        pattern, deps = analyze_one(fc, names, tables, all_sheets, out)
                    except SheetError:
                        unparsable += 1
                        continue
                    analyzed += 1
                    by_col[col][r] = pattern
                    texts[(r, col)] = text
                    ci.add(r, col)
                    if deps:
                        deps_map[(sc.name, r, col)] = deps
        ci.freeze()
        pattern_outliers(by_col, consts, sc.name, texts, out)
        for code, addrs in sc.errors.items():
            for a in addrs:
                out.add("cached_error", f"{sc.name}!{a}", f"shows {code}")
            if sc.error_total > sum(len(v) for v in sc.errors.values()):
                out.notes.append(f"{sc.name}: more {code} cells exist than the {len(addrs)} listed")
    cycles, complete = find_cycles(deps_map, index)
    for comp in cycles:
        cells = [f"{s}!{get_column_letter(c)}{r}" for s, r, c in comp]
        shown = ", ".join(cells[:8]) + (f" ... ({len(cells)} cells)" if len(cells) > 8 else "")
        out.add("circular", cells[0], f"cells that depend on each other: {shown}")
    if not complete:
        out.notes.append("circular-reference check stopped early: too many cell dependencies to expand")

    total_formulas = sum(s.n_formulas for s in targets)
    lines = [
        f"Checked {analyzed} of {total_formulas} formulas on {len([s for s in targets if s.n_formulas])} sheet(s). "
        "Nothing was calculated; checks are structural."
    ]
    if skipped:
        lines.append(f"Stopped analysing after {max_formulas} formulas ({skipped} skipped); pass sheet= to focus.")
    if unparsable:
        lines.append(f"{unparsable} formula(s) could not be parsed and were skipped.")
    if not out.issues:
        lines.append("No problems found.")
    else:
        lines.append(
            "Findings: " + ", ".join(f"{len(v)} {k}" for k, v in out.issues.items() if v)
        )
    for kind, title in TITLES.items():
        issues = out.issues.get(kind)
        if not issues:
            continue
        lines.append("")
        if kind in ("hardcoded_number", "volatile"):
            groups: dict[str, list[Issue]] = defaultdict(list)
            for i in issues:
                groups[i.group].append(i)
            lines.append(f"{title}: {len(issues)} cell(s) in {len(groups)} formula pattern(s)")
            for g in list(groups.values())[:max_issues]:
                first, n = g[0], len(g)
                more = f" (and {n - 1} more cells with the same pattern, last {g[-1].where})" if n > 1 else ""
                lines.append(f"  {first.where}: {first.message}{more}")
            if len(groups) > max_issues:
                lines.append(f"  ... {len(groups) - max_issues} more patterns")
        else:
            lines.append(f"{title}: {len(issues)}")
            for i in issues[:max_issues]:
                lines.append(f"  {i.where}: {i.message}")
            if len(issues) > max_issues:
                lines.append(f"  ... {len(issues) - max_issues} more (raise max_issues)")
    lines.append("")
    lines.append(
        "Limits: circular references are found only through direct cell/range references "
        "(not INDIRECT, OFFSET, names or table references); column-pattern checks only flag a "
        "cell that differs from the formulas directly above and below it."
    )
    lines.extend(out.notes)
    return "\n".join(lines)
