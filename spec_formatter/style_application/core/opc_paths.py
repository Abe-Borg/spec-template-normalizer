"""Fail-closed helpers for OPC package part and relationship paths."""

from __future__ import annotations

import posixpath
from pathlib import PurePosixPath, PureWindowsPath
from urllib.parse import unquote, urlsplit


_RESERVED_WORD_PARTS = {
    "word/document.xml",
    "word/styles.xml",
    "word/numbering.xml",
    "word/settings.xml",
    "word/fonttable.xml",
    "word/websettings.xml",
}
_RESERVED_WINDOWS_NAMES = {"CON", "PRN", "AUX", "NUL"} | {
    f"{prefix}{number}"
    for prefix in ("COM", "LPT")
    for number in range(1, 10)
}


def is_safe_package_part_name(name: str) -> bool:
    """Return whether *name* is a normalized, relative OPC ZIP member name."""
    if not isinstance(name, str) or not name or name.endswith("/"):
        return False
    if "\x00" in name or "\\" in name or any(ord(ch) < 32 for ch in name):
        return False
    parsed = urlsplit(name)
    decoded = unquote(name)
    windows_path = PureWindowsPath(decoded)
    path = PurePosixPath(decoded)
    if (
        parsed.scheme
        or parsed.netloc
        or parsed.query
        or parsed.fragment
        or decoded.startswith(("/", "//"))
        or windows_path.drive
        or windows_path.is_absolute()
        or path.is_absolute()
        or any(part in {"", ".", ".."} or ":" in part for part in path.parts)
    ):
        return False
    for part in path.parts:
        stem = part.split(".", 1)[0].upper()
        if part.endswith((" ", ".")) or stem in _RESERVED_WINDOWS_NAMES:
            return False
    # Reject encoded aliases and non-canonical spellings so security checks and
    # ZIP member comparisons operate on one exact representation.
    return decoded == name and posixpath.normpath(name) == name


def is_safe_header_footer_part_name(
    name: str,
    expected_kind: str | None = None,
) -> bool:
    """Validate a header/footer owner part without relying on its basename.

    OOXML relationship types, not filenames, identify headers and footers.
    Phase 1 may therefore carry safe custom paths such as
    ``word/headers/default.xml``.
    """
    if expected_kind not in {None, "header", "footer"}:
        return False
    folded = name.casefold() if isinstance(name, str) else ""
    if (
        not is_safe_package_part_name(name)
        or not folded.startswith("word/")
        or not folded.endswith(".xml")
        or "/_rels/" in folded
        or folded.startswith("word/media/")
        or folded in _RESERVED_WORD_PARTS
    ):
        return False
    return True


def relationship_part_name_for_owner(owner_part: str) -> str:
    """Return the OPC relationships part owned by *owner_part*."""
    if not is_safe_package_part_name(owner_part):
        raise ValueError(f"Unsafe OPC owner part name: {owner_part!r}")
    directory, basename = posixpath.split(owner_part)
    if not basename:
        raise ValueError(f"OPC owner part has no basename: {owner_part!r}")
    return posixpath.join(directory, "_rels", f"{basename}.rels")


def resolve_internal_relationship_target(owner_part: str, target: str) -> str:
    """Resolve a non-external relationship target to a safe package part."""
    if not isinstance(target, str) or not target or "\x00" in target or "\\" in target:
        raise ValueError(f"Unsafe relationship target: {target!r}")
    if "%2e" in target.casefold():
        raise ValueError(f"Unsafe encoded-dot relationship target: {target!r}")
    decoded = unquote(target)
    parsed = urlsplit(decoded)
    if parsed.scheme or parsed.netloc or parsed.query or parsed.fragment:
        raise ValueError(f"Unsafe relationship target: {target!r}")
    if decoded.startswith("/"):
        resolved = posixpath.normpath(decoded.lstrip("/"))
    else:
        resolved = posixpath.normpath(
            posixpath.join(posixpath.dirname(owner_part), decoded)
        )
    if not is_safe_package_part_name(resolved):
        raise ValueError(f"Unsafe relationship target: {target!r}")
    return resolved


def relationship_target_for_part(owner_part: str, target_part: str) -> str:
    """Return a normalized relative target from *owner_part* to *target_part*."""
    if not is_safe_package_part_name(owner_part):
        raise ValueError(f"Unsafe OPC owner part name: {owner_part!r}")
    if not is_safe_package_part_name(target_part):
        raise ValueError(f"Unsafe OPC target part name: {target_part!r}")
    target = posixpath.relpath(target_part, posixpath.dirname(owner_part))
    try:
        resolved = resolve_internal_relationship_target(owner_part, target)
    except ValueError as exc:
        raise ValueError(
            f"Cannot express OPC target safely: {owner_part!r} -> {target_part!r}"
        ) from exc
    if resolved != target_part:
        raise ValueError(
            f"OPC target resolution mismatch: {owner_part!r} -> {target_part!r}"
        )
    return target
