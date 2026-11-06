"""Path sandbox: every file the tools touch must live under one directory."""

from __future__ import annotations

from pathlib import Path

from .errors import SheetError

_root: Path | None = None

WORKBOOK_SUFFIXES = (".xlsx", ".xlsm")


def configure(root: str | Path | None) -> None:
    global _root
    _root = Path(root).resolve() if root is not None else None


def root() -> Path:
    return _root if _root is not None else Path.cwd().resolve()


def resolve(user_path: str, *, must_exist: bool, what: str = "file") -> Path:
    """Resolve a path inside the sandbox. Relative paths are relative to the sandbox root."""
    if not user_path or "\x00" in user_path:
        raise SheetError(f"{what} path is empty or invalid")
    base = root()
    p = Path(user_path)
    if not p.is_absolute():
        p = base / p
    try:
        resolved = p.resolve()
    except (OSError, ValueError) as e:
        raise SheetError(f"cannot resolve {what} path {user_path!r}: {e}") from e
    if not resolved.is_relative_to(base):
        raise SheetError(
            f"{what} path {user_path!r} is outside the allowed directory {base}. "
            "Use a path inside it (the server was started with --dir)."
        )
    if must_exist and not resolved.is_file():
        raise SheetError(f"{what} not found: {resolved}. Paths are relative to {base}.")
    return resolved


def resolve_workbook(user_path: str) -> Path:
    p = resolve(user_path, must_exist=True, what="workbook")
    if p.suffix.lower() not in WORKBOOK_SUFFIXES:
        raise SheetError(
            f"{p.name}: only .xlsx and .xlsm workbooks are supported "
            "(for .csv/.tsv use csv_to_xlsx first; for old .xls re-save as .xlsx in Excel)."
        )
    return p
