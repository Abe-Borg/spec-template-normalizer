# docx_patch.py
from __future__ import annotations

from pathlib import Path
import re
import zipfile
import xml.etree.ElementTree as ET
from typing import Dict, List, Set, Union

from .core.ooxml_text import prepare_xml_text_for_utf8
from .core.opc_paths import (
    is_safe_header_footer_part_name,
    is_safe_package_part_name,
    relationship_part_name_for_owner,
)

BytesOrStr = Union[bytes, str]


_HEADER_FOOTER_PART_RE = re.compile(
    r"word/(?P<kind>header|footer)[A-Za-z0-9_.-]*\.xml\Z"
)
_HEADER_FOOTER_RELS_RE = re.compile(
    r"word/_rels/(?:header|footer)[A-Za-z0-9_.-]*\.xml\.rels\Z"
)


def is_allowed_header_footer_part_name(
    name: str,
    expected_kind: str | None = None,
) -> bool:
    """Return whether *name* is a safe header/footer XML package part.

    OOXML part names are not required to use numeric suffixes (for example,
    ``word/headerFirst.xml`` is valid).  This predicate is shared with the
    importer so a part accepted there cannot be rejected later while the
    output package is assembled.
    """
    return is_safe_header_footer_part_name(name, expected_kind)


def is_allowed_header_footer_rels_name(
    name: str,
    owner_part: str | None = None,
) -> bool:
    """Return whether *name* is a safe relationships part for a header/footer."""
    if owner_part is not None:
        try:
            return (
                is_safe_header_footer_part_name(owner_part)
                and name == relationship_part_name_for_owner(owner_part)
            )
        except ValueError:
            return False
    return bool(_HEADER_FOOTER_RELS_RE.fullmatch(name))


def validate_xml_wellformedness(replacements: Dict[str, bytes]) -> List[str]:
    """Parse each XML replacement part with ElementTree to verify well-formedness."""
    errors: List[str] = []
    for name, content in replacements.items():
        if not (name.endswith(".xml") or name.endswith(".rels") or name == "[Content_Types].xml"):
            continue
        try:
            ET.fromstring(content)
        except ET.ParseError as exc:
            errors.append(f"{name}: XML parse error: {exc}")
    return errors


def _is_allowed_patch(
    name: str,
    allowed_patches: Set[str],
    allowed_dynamic_parts: Set[str],
) -> bool:
    if name in allowed_patches:
        return True
    if (
        _HEADER_FOOTER_PART_RE.fullmatch(name)
        or _HEADER_FOOTER_RELS_RE.fullmatch(name)
    ):
        return True
    if name.startswith("word/media/") and is_safe_package_part_name(name):
        return True
    if name in allowed_dynamic_parts and is_safe_package_part_name(name):
        return True
    return False


def patch_docx(
    src_docx: Path,
    out_docx: Path,
    replacements: Dict[str, BytesOrStr],
    exclude_parts: Set[str] | None = None,
    allowed_dynamic_parts: Set[str] | None = None,
) -> None:
    """
    Create out_docx by copying every ZIP entry from src_docx unchanged,
    except for entries whose internal paths match keys in `replacements`.

    This is NOT a "rebuild from extracted folder".
    It's a surgical patch: swap specific parts, preserve everything else.
    """
    src_docx = Path(src_docx)
    out_docx = Path(out_docx)
    exclude_parts = set(exclude_parts or set())
    allowed_dynamic_parts = set(allowed_dynamic_parts or set())

    dynamic_names: Dict[str, str] = {}
    for name in allowed_dynamic_parts:
        normalized_name = name.casefold()
        previous_name = dynamic_names.get(normalized_name)
        if previous_name is not None:
            raise RuntimeError(
                "Dynamic patch allow-list contains duplicate names "
                f"(case-insensitive): {previous_name!r} and {name!r}"
            )
        if not is_safe_package_part_name(name):
            raise RuntimeError(f"Unsafe dynamic patch target: {name!r}")
        dynamic_names[normalized_name] = name

    rep_bytes: Dict[str, bytes] = {}
    replacement_names: Dict[str, str] = {}
    for k, v in replacements.items():
        normalized_name = k.casefold()
        previous_name = replacement_names.get(normalized_name)
        if previous_name is not None:
            raise RuntimeError(
                "Replacement parts contain duplicate names (case-insensitive): "
                f"{previous_name!r} and {k!r}"
            )
        replacement_names[normalized_name] = k
        if isinstance(v, str):
            if k.endswith((".xml", ".rels")) or k == "[Content_Types].xml":
                v = prepare_xml_text_for_utf8(v)
            rep_bytes[k] = v.encode("utf-8")
        else:
            rep_bytes[k] = v

    FORBIDDEN_EXACT = set()

    ALLOWED_PATCHES = {
        "word/document.xml",
        "word/styles.xml",
        "word/theme/theme1.xml",
        "word/numbering.xml",
        "word/settings.xml",
        "word/fontTable.xml",
        "[Content_Types].xml",
        "word/_rels/document.xml.rels",
    }

    for name in rep_bytes:
        if name in FORBIDDEN_EXACT:
            raise RuntimeError(f"Forbidden patch target: {name}")

        if not _is_allowed_patch(name, ALLOWED_PATCHES, allowed_dynamic_parts):
            raise RuntimeError(
                f"Illegal patch target: {name}\n"
                f"Allowed base set: {sorted(ALLOWED_PATCHES)} plus header/footer/media patterns"
            )

    for name in exclude_parts:
        if not _is_allowed_patch(name, ALLOWED_PATCHES, allowed_dynamic_parts):
            raise RuntimeError(f"Illegal excluded part: {name}")

    # Validate XML well-formedness before writing — refuse to build a broken DOCX
    xml_errors = validate_xml_wellformedness(rep_bytes)
    if xml_errors:
        raise RuntimeError(
            "XML well-formedness check failed — refusing to build DOCX:\n"
            + "\n".join(f"  - {e}" for e in xml_errors)
        )

    with zipfile.ZipFile(src_docx, "r") as zin:
        infos = zin.infolist()
        seen_names: Dict[str, str] = {}
        for info in infos:
            normalized = info.filename.casefold()
            previous = seen_names.get(normalized)
            if previous is not None:
                raise RuntimeError(
                    "Source DOCX contains duplicate ZIP member names "
                    f"(case-insensitive): {previous!r} and {info.filename!r}"
                )
            seen_names[normalized] = info.filename

        # ZIP member names are case-sensitive at the container level, but OPC
        # consumers commonly resolve them case-insensitively.  A replacement
        # whose spelling differs only by case would otherwise leave the source
        # member in place and append a second, ambiguous output part.
        for normalized, replacement_name in replacement_names.items():
            source_name = seen_names.get(normalized)
            if source_name is not None and source_name != replacement_name:
                raise RuntimeError(
                    "Replacement part collides with a source ZIP member "
                    "(case-insensitive): "
                    f"{replacement_name!r} and {source_name!r}"
                )

        out_docx.parent.mkdir(parents=True, exist_ok=True)
        if out_docx.exists():
            out_docx.unlink()

        with zipfile.ZipFile(out_docx, "w") as zout:
            # preserve archive comment if any
            zout.comment = zin.comment

            src_names = set(seen_names.values())

            # For new parts (like theme1.xml if it didn't exist), we'll add them
            new_parts = [name for name in rep_bytes.keys() if name not in src_names]

            for info in infos:
                name = info.filename
                if name in exclude_parts:
                    continue
                data = rep_bytes.get(name, zin.read(info))

                # Preserve per-entry compression type where possible
                zout.writestr(info, data, compress_type=info.compress_type)

            # Add any new parts that didn't exist in source
            for new_name in new_parts:
                zout.writestr(new_name, rep_bytes[new_name])
