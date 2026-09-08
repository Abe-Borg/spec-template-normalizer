"""
arch_env_applier.py — Phase 2 Environment Application

Applies the formatting environment captured in arch_template_registry.json
to a target document. This ensures that imported styles render correctly
by providing the same context (theme fonts, docDefaults, etc.) they expect.

Application order (deterministic):
1. Theme (fonts/colors foundation)
2. Settings + compat flags (rendering behavior)  
3. Font table (declared fonts)
4. docDefaults (baseline rPr/pPr)
5. Styles (with materialized typography)

NOTE: This module does NOT touch:
- numbering.xml (handled separately with explicit numPr materialization)

It now imports architect headers/footers after page layout sync.

It DOES sync page layout in document.xml sectPr blocks for managed tags:
- w:pgSz
- w:pgMar
- w:cols
- w:docGrid

Usage:
    from arch_env_applier import apply_environment_to_target
    
    apply_environment_to_target(
        target_extract_dir=Path("mech_spec_extracted"),
        registry=loaded_registry_dict,
        log=[]
    )
"""

from __future__ import annotations

import re
from pathlib import Path
from typing import Any, Dict, List, Optional

import xml.etree.ElementTree as ET

from .core.registry import _check_xml_fragment
from .core.ooxml_namespaces import (
    CT_NS,
    PKG_REL_NS,
    R_NS,
    serialize_content_types,
    serialize_package_relationships,
)
from .core.ooxml_text import prepare_xml_text_for_utf8, read_xml_text, write_xml_text
from .core.untrusted_xml import parse_untrusted_xml
from .core.section_mapping import choose_section_sources
from .core.sectpr_tools import (
    CANONICAL_SECTPR_ORDER,
    child_tag_name,
    extract_all_sectpr_blocks,
    extract_sectpr_children,
    extract_tag_block,
    has_body_level_sectpr,
    replace_nth_sectpr_block,
    strip_tag_block,
)

MANAGED_LAYOUT_TAGS = (
    "pgSz", "pgMar", "cols", "docGrid",
)
_CANONICAL_SECTPR_ORDER = CANONICAL_SECTPR_ORDER

# ─────────────────────────────────────────────────────────────────────────────
# docDefaults application
# ─────────────────────────────────────────────────────────────────────────────

def _extract_doc_defaults_block(styles_xml: str) -> Optional[str]:
    """Extract existing <w:docDefaults>...</w:docDefaults> block."""
    m = re.search(r'(<w:docDefaults\b[\s\S]*?</w:docDefaults>)', styles_xml)
    return m.group(1) if m else None

def _build_doc_defaults_block(
    default_rpr: Optional[str],
    default_ppr: Optional[str]
) -> str:
    """
    Build a complete <w:docDefaults> block from rPr and pPr.
    """
    parts = ["<w:docDefaults>"]
    
    if default_rpr:
        parts.append(f"  <w:rPrDefault>{default_rpr}</w:rPrDefault>")
    else:
        parts.append("  <w:rPrDefault><w:rPr/></w:rPrDefault>")
    
    if default_ppr:
        parts.append(f"  <w:pPrDefault>{default_ppr}</w:pPrDefault>")
    else:
        parts.append("  <w:pPrDefault><w:pPr/></w:pPrDefault>")
    
    parts.append("</w:docDefaults>")
    return "\n".join(parts)

def apply_doc_defaults(
    styles_xml: str,
    registry: Dict[str, Any],
    log: List[str]
) -> str:
    """
    Replace or insert docDefaults in styles.xml with values from registry.
    
    This is critical because styles inherit from docDefaults, and if the
    target document has different defaults, fonts/spacing will be wrong.
    """
    doc_defaults = registry.get("doc_defaults", {})
    
    arch_rpr = doc_defaults.get("default_run_props", {}).get("rPr")
    arch_ppr = doc_defaults.get("default_paragraph_props", {}).get("pPr")
    
    if not arch_rpr and not arch_ppr:
        log.append("No docDefaults in registry; skipping docDefaults application")
        return styles_xml
    
    new_defaults = _build_doc_defaults_block(arch_rpr, arch_ppr)
    
    existing = _extract_doc_defaults_block(styles_xml)
    if existing:
        # Replace existing docDefaults
        styles_xml = styles_xml.replace(existing, new_defaults, 1)
        log.append("Replaced existing docDefaults with architect values")
    else:
        # Insert after <w:styles ...> opening tag
        m = re.search(r'(<w:styles\b[^>]*>)', styles_xml)
        if m:
            insert_point = m.end()
            styles_xml = (
                styles_xml[:insert_point] + 
                "\n" + new_defaults + "\n" + 
                styles_xml[insert_point:]
            )
            log.append("Inserted docDefaults from architect (none existed)")
        else:
            log.append("WARNING: Could not find <w:styles> tag to insert docDefaults")
    
    return styles_xml

# ─────────────────────────────────────────────────────────────────────────────
# Theme application
# ─────────────────────────────────────────────────────────────────────────────

def apply_theme(
    target_extract_dir: Path,
    registry: Dict[str, Any],
    log: List[str]
) -> None:
    """
    Copy theme1.xml from registry to target.
    
    Theme defines majorFont/minorFont which styles reference via
    w:asciiTheme="majorHAnsi" etc. Without the correct theme,
    font resolution fails.
    """
    theme_data = registry.get("theme", {})
    theme_xml = theme_data.get("theme1_xml")
    
    if not theme_xml:
        log.append("No theme in registry; skipping theme application")
        return
    
    theme_dir = target_extract_dir / "word" / "theme"
    theme_dir.mkdir(parents=True, exist_ok=True)
    
    theme_path = theme_dir / "theme1.xml"
    
    # Check if target already has a theme
    if theme_path.exists():
        log.append("Replacing target theme1.xml with architect theme")
    else:
        log.append("Adding theme1.xml from architect (none existed)")
        # May need to update [Content_Types].xml and relationships
        _ensure_theme_in_content_types(target_extract_dir, log)
        _ensure_theme_in_rels(target_extract_dir, log)
    
    write_xml_text(theme_path, theme_xml)

_CT_OVERRIDE_TYPES = {
    "/word/theme/theme1.xml": "application/vnd.openxmlformats-officedocument.theme+xml",
    "/word/settings.xml": (
        "application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"
    ),
    "/word/fontTable.xml": (
        "application/vnd.openxmlformats-officedocument.wordprocessingml.fontTable+xml"
    ),
}
_DOCUMENT_REL_TYPES = {
    "theme/theme1.xml": f"{R_NS}/theme",
    "settings.xml": f"{R_NS}/settings",
    "fontTable.xml": f"{R_NS}/fontTable",
}


def _ensure_override_in_content_types(
    extract_dir: Path,
    part_name: str,
    log: List[str],
    label: str,
) -> None:
    """Ensure [Content_Types].xml carries an Override for ``part_name``.

    The part is parsed (DOCTYPE-safe), matched case-insensitively on
    ``PartName`` as OPC consumers do, appended through ElementTree, and
    written back with the prefix-stable content-types serializer, so this
    no longer depends on the file's exact spelling of ``</Types>`` or on
    which importer last rewrote it.
    """

    ct_path = extract_dir / "[Content_Types].xml"
    if not ct_path.exists():
        return
    ct_xml = read_xml_text(ct_path)
    if not ct_xml.strip():
        return
    root = parse_untrusted_xml(prepare_xml_text_for_utf8(ct_xml), "[Content_Types].xml")
    if root.tag != f"{{{CT_NS}}}Types":
        raise ValueError("Invalid [Content_Types].xml root element")
    wanted = part_name.casefold()
    for node in root.findall(f"{{{CT_NS}}}Override"):
        if node.attrib.get("PartName", "").casefold() == wanted:
            return
    ET.SubElement(
        root,
        f"{{{CT_NS}}}Override",
        {"PartName": part_name, "ContentType": _CT_OVERRIDE_TYPES[part_name]},
    )
    ct_path.write_bytes(serialize_content_types(root))
    log.append(f"Added {label} to [Content_Types].xml")


def _ensure_relationship_in_document_rels(
    extract_dir: Path,
    target: str,
    log: List[str],
    label: str,
) -> None:
    """Ensure word/_rels/document.xml.rels relates the document to ``target``.

    Matching is by relationship Type URI (the canonical identity of the
    part), not by a quoted substring of the Target attribute.
    """

    rels_path = extract_dir / "word" / "_rels" / "document.xml.rels"
    if not rels_path.exists():
        return
    rels_xml = read_xml_text(rels_path)
    if not rels_xml.strip():
        return
    root = parse_untrusted_xml(
        prepare_xml_text_for_utf8(rels_xml),
        "word/_rels/document.xml.rels",
    )
    if root.tag != f"{{{PKG_REL_NS}}}Relationships":
        raise ValueError("Invalid document.xml.rels root element")
    rel_type = _DOCUMENT_REL_TYPES[target]
    relationships = root.findall(f"{{{PKG_REL_NS}}}Relationship")
    if any(node.attrib.get("Type", "") == rel_type for node in relationships):
        return
    existing_ids = {node.attrib.get("Id", "") for node in relationships}
    numeric_rids = [
        int(match.group(1))
        for rid in existing_ids
        if (match := re.fullmatch(r"rId(\d+)", rid))
    ]
    next_rid = (max(numeric_rids) if numeric_rids else 0) + 1
    while f"rId{next_rid}" in existing_ids:
        next_rid += 1
    new_rid = f"rId{next_rid}"
    ET.SubElement(
        root,
        f"{{{PKG_REL_NS}}}Relationship",
        {"Id": new_rid, "Type": rel_type, "Target": target},
    )
    rels_path.write_bytes(serialize_package_relationships(root))
    log.append(f"Added {label} relationship ({new_rid}) to document.xml.rels")


def _ensure_theme_in_content_types(extract_dir: Path, log: List[str]) -> None:
    """Ensure [Content_Types].xml has an entry for theme1.xml."""
    _ensure_override_in_content_types(extract_dir, "/word/theme/theme1.xml", log, "theme1.xml")


def _ensure_theme_in_rels(extract_dir: Path, log: List[str]) -> None:
    """Ensure word/_rels/document.xml.rels has a relationship for theme."""
    _ensure_relationship_in_document_rels(extract_dir, "theme/theme1.xml", log, "theme")


# ─────────────────────────────────────────────────────────────────────────────
# Settings plumbing helpers
# ─────────────────────────────────────────────────────────────────────────────

_MINIMAL_SETTINGS_XML = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
    '<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
    ' xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"'
    ' xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
    '\n</w:settings>'
)

def _ensure_settings_in_content_types(extract_dir: Path, log: List[str]) -> None:
    """Ensure [Content_Types].xml has an entry for settings.xml."""
    _ensure_override_in_content_types(extract_dir, "/word/settings.xml", log, "settings.xml")


def _ensure_settings_in_rels(extract_dir: Path, log: List[str]) -> None:
    """Ensure word/_rels/document.xml.rels has a relationship for settings."""
    _ensure_relationship_in_document_rels(extract_dir, "settings.xml", log, "settings")


# ─────────────────────────────────────────────────────────────────────────────
# Font table plumbing helpers
# ─────────────────────────────────────────────────────────────────────────────

def _ensure_font_table_in_content_types(extract_dir: Path, log: List[str]) -> None:
    """Ensure [Content_Types].xml has an entry for fontTable.xml."""
    _ensure_override_in_content_types(extract_dir, "/word/fontTable.xml", log, "fontTable.xml")


def _ensure_font_table_in_rels(extract_dir: Path, log: List[str]) -> None:
    """Ensure word/_rels/document.xml.rels has a relationship for fontTable."""
    _ensure_relationship_in_document_rels(extract_dir, "fontTable.xml", log, "fontTable")

# ─────────────────────────────────────────────────────────────────────────────
# Settings/compat application
# ─────────────────────────────────────────────────────────────────────────────

def apply_settings(
    target_extract_dir: Path,
    registry: Dict[str, Any],
    log: List[str],
    registry_dir: Optional[Path] = None,
) -> None:
    """
    Apply only the allowlisted rendering compatibility block from the registry.

    Compat flags affect rendering behavior (list spacing, line breaking, etc.)
    and can cause subtle visual differences if not matched.

    The exact source_settings.xml artifact is provenance, not a replacement
    payload. Word settings can contain document protection, revision tracking,
    mail-merge state, attached-template relationships, and other target-specific
    semantics. Copying it wholesale can both change behavior and create dangling
    relationship IDs because settings.xml.rels is not part of the Phase 1 bundle.

    If the target lacks settings.xml, a minimal valid part is created and
    wired into [Content_Types].xml and document.xml.rels idempotently.
    Malformed compat_xml is rejected before any mutation.
    """
    settings_data = registry.get("settings", {})
    settings_path = target_extract_dir / "word" / "settings.xml"

    compat_xml = settings_data.get("compat", {}).get("compat_xml")
    if not compat_xml:
        log.append("No compat flags in registry; skipping settings application")
        return

    # Validate compat_xml before any mutation
    err = _check_xml_fragment(compat_xml, "w:compat")
    if err:
        log.append(f"WARNING: Skipping compat application — {err}")
        return

    if not settings_path.exists():
        # Create minimal settings.xml and wire package plumbing
        settings_path.parent.mkdir(parents=True, exist_ok=True)
        write_xml_text(settings_path, _MINIMAL_SETTINGS_XML)
        _ensure_settings_in_content_types(target_extract_dir, log)
        _ensure_settings_in_rels(target_extract_dir, log)
        log.append("Created minimal settings.xml (none existed in target)")

    settings_xml = read_xml_text(settings_path)

    # Find and replace existing <w:compat> block
    existing_compat = re.search(r'<w:compat\b[\s\S]*?</w:compat>', settings_xml)

    if existing_compat:
        settings_xml = settings_xml.replace(existing_compat.group(0), compat_xml, 1)
        log.append("Replaced compat flags with architect values")
    else:
        # Insert before </w:settings>
        if "</w:settings>" in settings_xml:
            settings_xml = settings_xml.replace(
                "</w:settings>",
                f"  {compat_xml}\n</w:settings>"
            )
            log.append("Inserted compat flags from architect")

    write_xml_text(settings_path, settings_xml)

    # Post-mutation validation
    final_xml = read_xml_text(settings_path)
    post_err = _check_xml_fragment(final_xml, "w:settings")
    if post_err:
        log.append(f"WARNING: settings.xml may be malformed after mutation — {post_err}")

# ─────────────────────────────────────────────────────────────────────────────
# Font table application
# ─────────────────────────────────────────────────────────────────────────────

def apply_font_table(
    target_extract_dir: Path,
    registry: Dict[str, Any],
    log: List[str]
) -> None:
    """
    Merge font declarations from registry into target fontTable.xml.

    This ensures fonts referenced by architect styles are declared,
    which helps Word resolve them correctly.

    If the target has no fontTable.xml, the architect's font table is
    copied and [Content_Types].xml / document.xml.rels are wired
    idempotently.  The result is validated after mutation.
    """
    fonts_data = registry.get("fonts", {})
    arch_font_xml = fonts_data.get("font_table_xml")

    if not arch_font_xml:
        log.append("No fontTable in registry; skipping font table application")
        return

    font_path = target_extract_dir / "word" / "fontTable.xml"

    if not font_path.exists():
        # Copy the architect's font table and wire package plumbing
        font_path.parent.mkdir(parents=True, exist_ok=True)
        write_xml_text(font_path, arch_font_xml)
        _ensure_font_table_in_content_types(target_extract_dir, log)
        _ensure_font_table_in_rels(target_extract_dir, log)
        log.append("Added fontTable.xml from architect (with content types and rels)")
    else:
        # Merge: add fonts from architect that don't exist in target
        target_font_xml = read_xml_text(font_path)

        # Extract font names from both
        target_fonts = set(re.findall(r'<w:font\s+w:name="([^"]+)"', target_font_xml))
        arch_fonts = re.findall(r'(<w:font\s+w:name="([^"]+)"[\s\S]*?</w:font>)', arch_font_xml)

        fonts_to_add = []
        for font_block, font_name in arch_fonts:
            if font_name not in target_fonts:
                fonts_to_add.append(font_block)

        if not fonts_to_add:
            log.append("All architect fonts already present in target fontTable")
            return

        # Insert before </w:fonts>
        if "</w:fonts>" in target_font_xml:
            insertion = "\n".join(fonts_to_add)
            target_font_xml = target_font_xml.replace(
                "</w:fonts>",
                f"{insertion}\n</w:fonts>"
            )
            write_xml_text(font_path, target_font_xml)
            log.append(f"Added {len(fonts_to_add)} font declarations from architect")

    # Post-mutation validation
    final_xml = read_xml_text(font_path)
    post_err = _check_xml_fragment(final_xml, "w:fonts")
    if post_err:
        log.append(f"WARNING: fontTable.xml may be malformed after mutation — {post_err}")

def _extract_layout_signature(sectpr: str) -> Dict[str, Optional[str]]:
    pgmar = extract_tag_block(sectpr, "pgMar") or ""
    attrs = {}
    for key in ("top", "right", "bottom", "left", "header", "footer"):
        m = re.search(rf'w:{key}="([^"]+)"', pgmar)
        attrs[key] = m.group(1) if m else None
    return attrs

def _merge_managed_layout_tags(target_sectpr: str, source_sectpr: str) -> str:
    self_closing = re.fullmatch(r'(<w:sectPr\b[^>]*?)/\s*>', target_sectpr, flags=re.S)
    if self_closing:
        target_sectpr = f"{self_closing.group(1).rstrip()}></w:sectPr>"

    open_tag_m = re.match(r'(<w:sectPr\b[^>]*>)', target_sectpr)
    close_tag = "</w:sectPr>"
    if not open_tag_m or not target_sectpr.endswith(close_tag):
        return target_sectpr

    open_tag = open_tag_m.group(1)
    inner = target_sectpr[len(open_tag):-len(close_tag)]

    # Remove managed tags from target and prepare source replacements.
    for tag in MANAGED_LAYOUT_TAGS:
        inner = strip_tag_block(inner, tag)
    children = extract_sectpr_children(inner)

    managed_children = {
        tag: extract_tag_block(source_sectpr, tag)
        for tag in MANAGED_LAYOUT_TAGS
    }
    index_by_tag = {tag: idx for idx, tag in enumerate(_CANONICAL_SECTPR_ORDER)}

    for tag in MANAGED_LAYOUT_TAGS:
        block = managed_children.get(tag)
        if not block:
            continue
        target_order = index_by_tag[tag]

        insert_at = len(children)
        for i, child in enumerate(children):
            child_tag = child_tag_name(child)
            if child_tag and index_by_tag.get(child_tag, 10_000) > target_order:
                insert_at = i
                break
        children.insert(insert_at, block)

    return f"{open_tag}{''.join(children)}{close_tag}"


def _ensure_body_level_sectpr(document_xml: str, log: List[str]) -> str:
    """Append a final body-level sectPr when the target does not have one."""
    if has_body_level_sectpr(document_xml):
        return document_xml

    body_close = document_xml.rfind("</w:body>")
    if body_close < 0:
        raise ValueError(
            "Target word/document.xml has no closing w:body; cannot create section properties"
        )
    log.append("Created missing final body-level sectPr in target document.xml")
    return document_xml[:body_close] + "<w:sectPr></w:sectPr>" + document_xml[body_close:]

def _choose_layout_sources(target_count: int, page_layout: Dict[str, Any], log: List[str]) -> List[str]:
    sources = choose_section_sources(target_count, page_layout, require_default=True, log=log)
    out: List[str] = []
    for source in sources:
        out.append(source.get("sectPr") if isinstance(source, dict) else "")
    return out

def apply_page_layout(target_extract_dir: Path, registry: Dict[str, Any], log: List[str]) -> None:
    page_layout = registry.get("page_layout")
    if not isinstance(page_layout, dict):
        raise ValueError("Template registry missing page_layout required for Phase 2 page layout sync")

    doc_path = target_extract_dir / "word" / "document.xml"
    if not doc_path.exists():
        log.append("WARNING: No word/document.xml found; skipping page layout sync")
        return

    doc_xml = _ensure_body_level_sectpr(read_xml_text(doc_path), log)
    target_sectprs = extract_all_sectpr_blocks(doc_xml)
    if not target_sectprs:  # Defensive: _ensure_body_level_sectpr must create one.
        raise ValueError("Could not create target section properties for page layout sync")

    sources = _choose_layout_sources(len(target_sectprs), page_layout, log)
    updated_xml = doc_xml

    for idx, (target_sectpr, source_sectpr) in enumerate(zip(target_sectprs, sources)):
        if not source_sectpr:
            continue
        before_sig = _extract_layout_signature(target_sectpr)
        merged = _merge_managed_layout_tags(target_sectpr, source_sectpr)
        after_sig = _extract_layout_signature(merged)
        updated_xml = replace_nth_sectpr_block(updated_xml, idx, merged)
        log.append(f"Patched sectPr[{idx}] layout signature: {before_sig} -> {after_sig}")

    write_xml_text(doc_path, updated_xml)

def apply_environment_to_target(
    target_extract_dir: Path,
    registry: Dict[str, Any],
    log: List[str],
    apply_theme_flag: bool = True,
    apply_settings_flag: bool = True,
    apply_doc_defaults_flag: bool = True,
    apply_fonts_flag: bool = True,
    apply_headers_footers_flag: bool = True,
    registry_dir: Optional[Path] = None,
) -> Dict[str, Any]:
    """
    Apply the formatting environment from arch_template_registry to target.
    
    Application order (deterministic):
    1. Theme (font/color definitions)
    2. Settings + compat (rendering behavior)
    3. Font table (font declarations)
    4. docDefaults in styles.xml (baseline formatting)
    5. page layout managed tags in document.xml sectPr (pgSz/pgMar/cols/docGrid)
    6. headers/footers
    
    Args:
        target_extract_dir: Extracted target document folder
        registry: Loaded arch_template_registry.json
        log: List to append log messages
        apply_*: Flags to selectively disable parts of application
    """
    target_extract_dir = Path(target_extract_dir)
    
    log.append("=" * 60)
    log.append("BEGIN ENVIRONMENT APPLICATION")
    log.append("=" * 60)
    
    # 1. Theme
    if apply_theme_flag:
        log.append("\n[1/6] Applying theme...")
        apply_theme(target_extract_dir, registry, log)
    else:
        log.append("\n[1/6] Theme application skipped")
    
    # 2. Settings/compat
    if apply_settings_flag:
        log.append("\n[2/6] Applying settings/compat...")
        apply_settings(target_extract_dir, registry, log, registry_dir=registry_dir)
    else:
        log.append("\n[2/6] Settings application skipped")
    
    # 3. Font table
    if apply_fonts_flag:
        log.append("\n[3/6] Applying font table...")
        apply_font_table(target_extract_dir, registry, log)
    else:
        log.append("\n[3/6] Font table application skipped")
    
    # 4. docDefaults in styles.xml
    if apply_doc_defaults_flag:
        log.append("\n[4/6] Applying docDefaults...")
        styles_path = target_extract_dir / "word" / "styles.xml"
        if styles_path.exists():
            styles_xml = read_xml_text(styles_path)
            styles_xml = apply_doc_defaults(styles_xml, registry, log)
            write_xml_text(styles_path, styles_xml)
        else:
            log.append("WARNING: No styles.xml in target; cannot apply docDefaults")
    else:
        log.append("\n[4/6] docDefaults application skipped")

    log.append("\n[5/6] Applying page layout managed tags...")
    apply_page_layout(target_extract_dir, registry, log)

    hf_result: Dict[str, Any] = {
        "part_names": set(),
        "rels_names": set(),
        "media_names": set(),
        "removed_part_names": set(),
        "removed_rels_names": set(),
        "style_ids": set(),
        "direct_num_ids": set(),
    }
    if apply_headers_footers_flag:
        log.append("\n[6/6] Applying headers/footers...")
        from .header_footer_importer import import_headers_footers
        imported = import_headers_footers(target_extract_dir, registry, log)
        hf_result = {
            "part_names": set(imported.part_names),
            "rels_names": set(imported.rels_names),
            "media_names": set(imported.media_names),
            "removed_part_names": set(imported.removed_part_names),
            "removed_rels_names": set(imported.removed_rels_names),
            "style_ids": set(imported.style_ids),
            "direct_num_ids": set(imported.direct_num_ids),
        }
    else:
        log.append("\n[6/6] Headers/footers application skipped")
    
    log.append("\n" + "=" * 60)
    log.append("END ENVIRONMENT APPLICATION")
    log.append("=" * 60)
    return {"header_footer_import": hf_result}
