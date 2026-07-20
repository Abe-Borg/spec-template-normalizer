#!/usr/bin/env python3
"""
arch_env_extractor.py — Phase 1 Environment Capture

Extracts the complete formatting environment from an architect's Word template
and produces arch_template_registry.json — the "VM snapshot" that Phase 2 uses
to recreate the rendering context.

This captures everything Word uses to render a document BEYOND style definitions:
- docDefaults (default rPr/pPr)
- Theme (fonts, colors)
- Settings + compat flags
- Numbering definitions
- Page layout (sectPr, margins)
- Headers/footers
- Font table

This module is imported as a library by gui.py and phase1_smoke_test.py.
It has no CLI entry point.

The output JSON follows the arch_template_registry schema (v1.0.0).
"""

from __future__ import annotations

import base64
import hashlib
import posixpath
import re
from datetime import datetime, timezone
from pathlib import Path, PurePosixPath, PureWindowsPath
from typing import Any, Dict, List, Optional, Tuple
from urllib.parse import unquote, urlsplit
import xml.etree.ElementTree as ET

from ooxml_text import prepare_xml_text_for_utf8, read_xml_text
from docx_decomposer import extract_document_sectpr_blocks


# ─────────────────────────────────────────────────────────────────────────────
# Constants / Namespaces
# ─────────────────────────────────────────────────────────────────────────────

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"
R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
CONTENT_TYPES_NS = "http://schemas.openxmlformats.org/package/2006/content-types"

SCHEMA_VERSION = "1.0.0"

# Header/footer images should be small (logos, rules, and similar assets).  The
# limits protect extraction from a relationship that points at an unexpectedly
# large file without constraining normal templates.
MAX_HEADER_FOOTER_MEDIA_BYTES = 16 * 1024 * 1024
MAX_HEADER_FOOTER_MEDIA_TOTAL_BYTES = 64 * 1024 * 1024


# ─────────────────────────────────────────────────────────────────────────────
# Utility functions
# ─────────────────────────────────────────────────────────────────────────────

def sha256_bytes(b: bytes) -> str:
    return hashlib.sha256(b).hexdigest()


def sha256_file(p: Path) -> str:
    return sha256_bytes(p.read_bytes())


def _read_xml_part(extract_dir: Path, internal_path: str) -> Optional[str]:
    """Read an XML part from extracted DOCX folder, return None if missing."""
    p = extract_dir / internal_path
    if p.exists():
        return read_xml_text(p)
    return None


def _read_xml_part_bytes(extract_dir: Path, internal_path: str) -> Optional[bytes]:
    """Read raw bytes from extracted DOCX folder."""
    p = extract_dir / internal_path
    if p.exists():
        return p.read_bytes()
    return None


def _parse_relationships_xml(xml_text: str, part_name: str) -> List[ET.Element]:
    """Parse and minimally validate an OPC relationships part.

    Relationship XML controls which files the extractor may read.  Treating a
    malformed part as though it were empty hides package corruption and can
    make the resulting registry incomplete, so callers receive a clear error.
    """
    if "<!DOCTYPE" in xml_text.upper():
        raise ValueError(f"{part_name} must not contain a DOCTYPE declaration")

    try:
        root = ET.fromstring(xml_text)
    except ET.ParseError as exc:
        raise ValueError(f"Malformed relationships XML in {part_name}: {exc}") from exc

    relationships_tag = f"{{{REL_NS}}}Relationships"
    relationship_tag = f"{{{REL_NS}}}Relationship"
    if root.tag != relationships_tag:
        raise ValueError(
            f"Invalid relationships root in {part_name}: expected {relationships_tag!r}"
        )

    relationships: List[ET.Element] = []
    seen_ids = set()
    for child in root:
        if child.tag != relationship_tag:
            raise ValueError(f"Unexpected element {child.tag!r} in {part_name}")

        rel_id = child.get("Id")
        rel_type = child.get("Type")
        target = child.get("Target")
        if not rel_id or not rel_type or not target:
            raise ValueError(
                f"Relationship in {part_name} must have non-empty Id, Type, and Target"
            )
        if rel_id in seen_ids:
            raise ValueError(f"Duplicate relationship Id {rel_id!r} in {part_name}")
        seen_ids.add(rel_id)

        target_mode = child.get("TargetMode")
        if target_mode not in (None, "Internal", "External"):
            raise ValueError(
                f"Invalid TargetMode {target_mode!r} for relationship {rel_id!r} "
                f"in {part_name}"
            )
        relationships.append(child)

    return relationships


def _relationship_references(xml_text: str, part_name: str) -> List[Tuple[str, str]]:
    """Return ``(relationship_id, attribute_name)`` references in an XML part.

    Header/footer content can refer to relationships through ``r:id``,
    ``r:embed``, or ``r:link``.  Keeping this check next to relationship
    capture prevents a registry from carrying XML that Phase 2 cannot wire
    back to a real relationship.
    """
    if "<!DOCTYPE" in xml_text.upper():
        raise ValueError(f"{part_name} must not contain a DOCTYPE declaration")
    try:
        root = ET.fromstring(xml_text)
    except ET.ParseError as exc:
        raise ValueError(f"Malformed XML in {part_name}: {exc}") from exc

    references: List[Tuple[str, str]] = []
    relationship_attributes = {
        f"{{{R_NS}}}id": "r:id",
        f"{{{R_NS}}}embed": "r:embed",
        f"{{{R_NS}}}link": "r:link",
    }
    for element in root.iter():
        for attribute, display_name in relationship_attributes.items():
            rel_id = element.get(attribute)
            if rel_id is not None:
                if not rel_id.strip():
                    raise ValueError(
                        f"Empty {display_name} relationship reference in {part_name}"
                    )
                references.append((rel_id, display_name))
    return references


def _resolve_internal_relationship_target(
    extract_dir: Path,
    source_part: str,
    target: str,
) -> Tuple[Path, str]:
    """Resolve an internal relationship target without leaving *extract_dir*.

    OPC targets are URI paths. Absolute paths, Windows drive/UNC paths,
    backslash-separated paths, URI schemes, and targets that escape the
    package root are rejected before touching the filesystem. Legal parent
    segments remain supported for nested owners (for example a custom header
    at ``word/headers/default.xml`` targeting ``../media/logo.png``).
    ``Path.resolve`` then catches escapes via symlinks already present in the
    extraction tree.
    """
    if not target or "\x00" in target:
        raise ValueError(f"Unsafe empty or NUL-containing relationship target {target!r}")

    decoded_target = unquote(target)
    if re.search(r"%2e", target, flags=re.IGNORECASE):
        raise ValueError(f"Unsafe encoded-dot relationship target {target!r}")
    parsed = urlsplit(decoded_target)
    if parsed.scheme or parsed.netloc or parsed.query or parsed.fragment:
        raise ValueError(f"Unsafe URI relationship target {target!r}")
    if "\\" in decoded_target:
        raise ValueError(f"Unsafe backslash relationship target {target!r}")
    if decoded_target.startswith(("/", "//")):
        raise ValueError(f"Unsafe absolute relationship target {target!r}")

    windows_target = PureWindowsPath(decoded_target)
    if windows_target.drive or windows_target.is_absolute():
        raise ValueError(f"Unsafe drive or UNC relationship target {target!r}")

    source_path = PurePosixPath(source_part)
    if (
        source_path.is_absolute()
        or ".." in source_path.parts
        or "\\" in source_part
        or any(":" in part for part in source_path.parts)
    ):
        raise ValueError(f"Unsafe source part name {source_part!r}")

    joined = posixpath.normpath(
        posixpath.join(source_path.parent.as_posix(), decoded_target)
    )
    if joined in {"", ".", ".."} or joined.startswith("../"):
        raise ValueError(f"Unsafe traversing relationship target {target!r}")
    package_path = PurePosixPath(joined)
    if package_path.is_absolute() or any(
        part in {"", ".", ".."} or ":" in part for part in package_path.parts
    ):
        raise ValueError(f"Unsafe relationship target {target!r}")

    extraction_root = extract_dir.resolve()
    candidate = extraction_root.joinpath(*package_path.parts).resolve(strict=False)
    try:
        candidate.relative_to(extraction_root)
    except ValueError as exc:
        raise ValueError(
            f"Relationship target {target!r} escapes extraction root via a symlink"
        ) from exc

    return candidate, package_path.as_posix()


def _relationship_part_name(source_part: str) -> str:
    """Return the OPC relationships-part name adjacent to *source_part*.

    Relationship parts live in an ``_rels`` directory beside their owning
    part.  Building this name from the complete owner path is important for
    legal custom locations such as ``word/layout/headerFirst.xml``; using only
    the basename would incorrectly look under ``word/_rels``.
    """
    source_path = PurePosixPath(source_part)
    if (
        source_path.is_absolute()
        or not source_path.name
        or ".." in source_path.parts
        or "\\" in source_part
    ):
        raise ValueError(f"Unsafe source part name {source_part!r}")
    return (
        source_path.parent
        / "_rels"
        / f"{source_path.name}.rels"
    ).as_posix()


def _valid_content_type(value: Optional[str]) -> bool:
    """Return whether *value* is safe to emit as a MIME content type."""
    if not value or len(value) > 255 or any(ord(ch) < 32 for ch in value):
        return False
    return re.fullmatch(
        r"[A-Za-z0-9!#$&^_.+-]+/[A-Za-z0-9!#$&^_.+-]+",
        value,
    ) is not None


def _load_content_types(extract_dir: Path) -> Tuple[Dict[str, str], Dict[str, str]]:
    """Load OPC content-type defaults and overrides, falling back safely."""
    xml_text = _read_xml_part(extract_dir, "[Content_Types].xml")
    if not xml_text or "<!DOCTYPE" in xml_text.upper():
        return {}, {}

    try:
        root = ET.fromstring(xml_text)
    except ET.ParseError:
        return {}, {}
    if root.tag != f"{{{CONTENT_TYPES_NS}}}Types":
        return {}, {}

    defaults: Dict[str, str] = {}
    overrides: Dict[str, str] = {}
    for child in root:
        if child.tag == f"{{{CONTENT_TYPES_NS}}}Default":
            extension = (child.get("Extension") or "").lower().lstrip(".")
            content_type = child.get("ContentType")
            if extension and _valid_content_type(content_type):
                defaults[extension] = content_type  # type: ignore[assignment]
        elif child.tag == f"{{{CONTENT_TYPES_NS}}}Override":
            part_name = (child.get("PartName") or "").replace("\\", "/")
            content_type = child.get("ContentType")
            if part_name.startswith("/") and _valid_content_type(content_type):
                overrides[part_name] = content_type  # type: ignore[assignment]
    return defaults, overrides


def _content_type_for_part(
    package_part: str,
    defaults: Dict[str, str],
    overrides: Dict[str, str],
) -> str:
    """Return the package-declared MIME type or a conservative known fallback."""
    override = overrides.get(f"/{package_part}")
    if override:
        return override

    extension = PurePosixPath(package_part).suffix.lower().lstrip(".")
    declared_default = defaults.get(extension)
    if declared_default:
        return declared_default

    safe_fallbacks = {
        "png": "image/png",
        "jpg": "image/jpeg",
        "jpeg": "image/jpeg",
        "gif": "image/gif",
        "bmp": "image/bmp",
        "tif": "image/tiff",
        "tiff": "image/tiff",
        "svg": "image/svg+xml",
        "emf": "image/x-emf",
        "wmf": "image/x-wmf",
    }
    return safe_fallbacks.get(extension, "application/octet-stream")


def _on_off_value(xml_text: str, tag: str) -> bool:
    """Read an OOXML on/off element, including explicit false values."""
    match = re.search(rf"<w:{re.escape(tag)}\b(?P<attrs>[^>]*)>", xml_text)
    if not match:
        return False
    value_match = re.search(
        r"(?<![\w:])w:val\s*=\s*(['\"])(?P<value>.*?)\1",
        match.group("attrs"),
        flags=re.IGNORECASE,
    )
    if not value_match:
        return True
    return value_match.group("value").strip().lower() not in {"0", "false", "off"}


def _on_off_attribute(attrs: str, name: str, default: bool = False) -> bool:
    """Read an OOXML on/off attribute from an opening tag's attribute text."""
    value_match = re.search(
        rf"(?<![\w:])w:{re.escape(name)}\s*=\s*(['\"])(?P<value>.*?)\1",
        attrs,
        flags=re.IGNORECASE,
    )
    if not value_match:
        return default
    return value_match.group("value").strip().lower() not in {"0", "false", "off"}


def _qualified_attribute(attrs: str, name: str) -> Optional[str]:
    """Read a qualified XML attribute using either XML quote style."""
    value_match = re.search(
        rf"(?<![\w:]){re.escape(name)}\s*=\s*(['\"])(?P<value>.*?)\1",
        attrs,
    )
    return value_match.group("value") if value_match else None


def _extract_block(xml_text: str, tag: str, ns_prefix: str = "w",
                    _start: int = 0) -> Optional[Tuple[str, int]]:
    """
    Extract the first occurrence of <ns:tag>...</ns:tag> from *xml_text*
    starting at *_start*.  Returns ``(block_text, end_position)`` so that
    :func:`_extract_all_blocks` can call repeatedly, or ``None`` when no
    match is found.

    Uses a forward token-scanner with depth tracking instead of a single
    regex so that paired elements (e.g. ``<w:compat>…</w:compat>``) are
    never truncated to just the opening tag.
    """
    qname = f"{ns_prefix}:{tag}"
    open_token = f"<{qname}"
    close_token = f"</{qname}"

    # -- locate the opening tag -------------------------------------------
    pos = xml_text.find(open_token, _start)
    if pos == -1:
        return None

    # The character after the tag name must be whitespace, '/', or '>'
    # to avoid matching a longer tag name (e.g. <w:stylePaneSortMethod>
    # when searching for <w:style>).
    after = pos + len(open_token)
    if after < len(xml_text) and xml_text[after] not in (" ", "\t", "\n",
                                                          "\r", ">", "/"):
        # Not a true match — skip past and retry
        return _extract_block(xml_text, tag, ns_prefix, _start=after)

    # -- find end of the opening tag ('>' character) ----------------------
    gt = xml_text.find(">", after)
    if gt == -1:
        return None  # malformed XML

    # -- self-closing? ----------------------------------------------------
    if xml_text[gt - 1] == "/":
        block = xml_text[pos:gt + 1]
        return (block, gt + 1)

    # -- paired tag: scan forward tracking depth --------------------------
    depth = 1
    cursor = gt + 1
    while depth > 0:
        next_open = xml_text.find(open_token, cursor)
        next_close = xml_text.find(close_token, cursor)

        if next_close == -1:
            return None  # no matching closer — malformed

        # If there is a nested opener *before* the next closer, handle it
        # only when the opener is a true tag (word-boundary check).
        if next_open != -1 and next_open < next_close:
            after_open = next_open + len(open_token)
            if after_open < len(xml_text) and xml_text[after_open] in (
                    " ", "\t", "\n", "\r", ">", "/"):
                # genuine nested opener
                depth += 1
            cursor = after_open
            continue

        # Process the closer
        end = xml_text.find(">", next_close + len(close_token))
        if end == -1:
            return None  # malformed
        depth -= 1
        cursor = end + 1

    block = xml_text[pos:cursor]
    return (block, cursor)


def _extract_first_block(xml_text: str, tag: str,
                         ns_prefix: str = "w") -> Optional[str]:
    """Public convenience: return only the block text (or ``None``)."""
    result = _extract_block(xml_text, tag, ns_prefix)
    return result[0] if result else None


def _extract_all_blocks(xml_text: str, tag: str,
                        ns_prefix: str = "w") -> List[str]:
    """Extract all occurrences of a namespaced tag as raw XML strings."""
    blocks: List[str] = []
    start = 0
    while True:
        result = _extract_block(xml_text, tag, ns_prefix, _start=start)
        if result is None:
            break
        block_text, end_pos = result
        blocks.append(block_text)
        start = end_pos
    return blocks


def _strip_rsids(xml_text: str) -> str:
    """Remove all rsid* attributes (revision save IDs) for cleaner diffs."""
    # Matches w:rsidR="..." w:rsidRPr="..." etc.
    return re.sub(r'\s+w:rsid\w*="[^"]*"', '', xml_text)


def _strip_proofing(xml_text: str) -> str:
    """Remove proofErr elements."""
    return re.sub(r'<w:proofErr[^>]*/>', '', xml_text)


def _canonicalize(xml_text: str, strip_rsid: bool = True, strip_proof: bool = True) -> str:
    """Apply standard cleanup and make registry XML UTF-8 serializable.

    Source-exact artifacts are copied as bytes by :mod:`phase1_pipeline` and
    never pass through this helper.  Every XML string carried inside the JSON
    registry does, so a UTF-16 or legacy declaration must be changed to match
    the UTF-8 JSON file that contains it.
    """
    if strip_rsid:
        xml_text = _strip_rsids(xml_text)
    if strip_proof:
        xml_text = _strip_proofing(xml_text)
    return prepare_xml_text_for_utf8(xml_text)


# ─────────────────────────────────────────────────────────────────────────────
# Extraction functions for each registry section
# ─────────────────────────────────────────────────────────────────────────────

def extract_package_inventory(extract_dir: Path) -> Dict[str, bool]:
    """Check which optional parts exist in the package."""
    return {
        "has_theme": (extract_dir / "word" / "theme" / "theme1.xml").exists(),
        "has_settings": (extract_dir / "word" / "settings.xml").exists(),
        "has_numbering": (extract_dir / "word" / "numbering.xml").exists(),
        "has_styles": (extract_dir / "word" / "styles.xml").exists(),
        "has_footnotes": (extract_dir / "word" / "footnotes.xml").exists(),
        "has_endnotes": (extract_dir / "word" / "endnotes.xml").exists(),
        "has_header_parts": any((extract_dir / "word").glob("header*.xml")),
        "has_footer_parts": any((extract_dir / "word").glob("footer*.xml")),
    }


def extract_doc_defaults(styles_xml: str) -> Dict[str, Any]:
    """
    Extract <w:docDefaults> which contains default rPr and pPr.
    These are the baseline formatting that all styles inherit from.
    """
    result = {
        "default_run_props": {"rPr": None},
        "default_paragraph_props": {"pPr": None},
    }
    
    doc_defaults = _extract_first_block(styles_xml, "docDefaults")
    if not doc_defaults:
        return result
    
    # Extract rPrDefault/rPr
    rpr_default = _extract_first_block(doc_defaults, "rPrDefault")
    if rpr_default:
        rpr = _extract_first_block(rpr_default, "rPr")
        if rpr:
            result["default_run_props"]["rPr"] = _canonicalize(rpr)
    
    # Extract pPrDefault/pPr
    ppr_default = _extract_first_block(doc_defaults, "pPrDefault")
    if ppr_default:
        ppr = _extract_first_block(ppr_default, "pPr")
        if ppr:
            result["default_paragraph_props"]["pPr"] = _canonicalize(ppr)
    
    return result


def extract_style_defs(styles_xml: str) -> List[Dict[str, Any]]:
    """
    Extract all style definitions with their raw pPr/rPr blocks.
    We store the raw XML to preserve exact formatting.
    """
    style_defs = []
    
    # Find all <w:style> blocks
    style_blocks = _extract_all_blocks(styles_xml, "style")
    
    for block in style_blocks:
        # Parse key attributes
        style_id_m = re.search(r'w:styleId="([^"]+)"', block)
        style_type_m = re.search(r'w:type="([^"]+)"', block)
        
        if not style_id_m:
            continue
            
        style_id = style_id_m.group(1)
        style_type = style_type_m.group(1) if style_type_m else None
        
        # Extract name
        name_m = re.search(r'<w:name\s+w:val="([^"]+)"', block)
        name = name_m.group(1) if name_m else None
        
        # Extract basedOn
        based_on_m = re.search(r'<w:basedOn\s+w:val="([^"]+)"', block)
        based_on = based_on_m.group(1) if based_on_m else None
        
        # Extract next style
        next_m = re.search(r'<w:next\s+w:val="([^"]+)"', block)
        next_style = next_m.group(1) if next_m else None
        
        # Extract link (for paragraph/character style pairs)
        link_m = re.search(r'<w:link\s+w:val="([^"]+)"', block)
        link = link_m.group(1) if link_m else None
        
        # UI properties
        ui_priority_m = re.search(r'<w:uiPriority\s+w:val="([^"]+)"', block)
        qformat = _on_off_value(block, "qFormat")
        semi_hidden = _on_off_value(block, "semiHidden")
        unhide_when_used = _on_off_value(block, "unhideWhenUsed")
        locked = _on_off_value(block, "locked")
        
        # Extract raw pPr and rPr blocks
        pPr = _extract_first_block(block, "pPr")
        rPr = _extract_first_block(block, "rPr")
        tblPr = _extract_first_block(block, "tblPr")
        trPr = _extract_first_block(block, "trPr")
        tcPr = _extract_first_block(block, "tcPr")
        
        style_defs.append({
            "style_id": style_id,
            "name": name,
            "type": style_type,
            "based_on": based_on,
            "next": next_style,
            "link": link,
            "ui_priority": int(ui_priority_m.group(1)) if ui_priority_m else None,
            "qformat": qformat,
            "semi_hidden": semi_hidden,
            "unhide_when_used": unhide_when_used,
            "locked": locked,
            "rsid": None,  # Stripped
            "raw_style_xml": _canonicalize(block),
            "pPr": _canonicalize(pPr) if pPr else None,
            "rPr": _canonicalize(rPr) if rPr else None,
            "tblPr": _canonicalize(tblPr) if tblPr else None,
            "trPr": _canonicalize(trPr) if trPr else None,
            "tcPr": _canonicalize(tcPr) if tcPr else None,
            "notes": {
                "normalized": True,
                "source_xpath": f"/w:styles/w:style[@w:styleId='{style_id}']"
            }
        })
    
    return style_defs


def extract_latent_styles(styles_xml: str) -> Dict[str, Any]:
    """Extract the latentStyles block (defines hidden/default style behaviors)."""
    latent = _extract_first_block(styles_xml, "latentStyles")
    return {
        "latentStyles_xml": _canonicalize(latent) if latent else None
    }


def extract_table_styles(style_defs: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
    """Filter table styles from the full style list."""
    return [s for s in style_defs if s.get("type") == "table"]


def extract_styles_section(styles_xml: str) -> Dict[str, Any]:
    """Build the complete styles section of the registry."""
    style_defs = extract_style_defs(styles_xml)
    return {
        "style_defs": style_defs,
        "latent_styles": extract_latent_styles(styles_xml),
        "table_styles": extract_table_styles(style_defs),
    }


def extract_theme(extract_dir: Path) -> Dict[str, Any]:
    """
    Extract theme/theme1.xml which defines:
    - majorFont / minorFont (heading vs body fonts)
    - Color scheme
    - Font mappings for different scripts
    """
    theme_xml = _read_xml_part(extract_dir, "word/theme/theme1.xml")
    
    result = {
        "theme1_xml": None,
        "notes": {
            "affects": ["themeFonts", "themeColors", "majorMinorFontMapping"]
        }
    }
    
    if theme_xml:
        result["theme1_xml"] = _canonicalize(theme_xml, strip_rsid=False)
    
    return result


def extract_settings(extract_dir: Path) -> Dict[str, Any]:
    """
    Extract settings.xml which includes:
    - Compatibility flags (w:compat)
    - Document-level settings
    """
    settings_xml = _read_xml_part(extract_dir, "word/settings.xml")
    
    result = {
        "settings_xml": None,
        "compat": {
            "compat_xml": None,
            "important_flags": []
        }
    }
    
    if not settings_xml:
        return result
    
    result["settings_xml"] = _canonicalize(settings_xml)
    
    # Extract compat block specifically
    compat = _extract_first_block(settings_xml, "compat")
    if compat:
        result["compat"]["compat_xml"] = _canonicalize(compat)
        
        # Note important flags that affect rendering
        important = []
        if "useWord2013TrackBottomHyphenation" in compat:
            important.append("useWord2013TrackBottomHyphenation")
        if "doNotExpandShiftReturn" in compat:
            important.append("doNotExpandShiftReturn")
        if "<w:compatSetting" in compat:
            important.append("compatSetting:*")
        result["compat"]["important_flags"] = important
    
    return result


def extract_page_layout(document_xml: str, extract_dir: Path) -> Dict[str, Any]:
    """
    Extract page layout information from sectPr blocks.
    Documents can have multiple sections with different layouts.
    """
    result = {
        "default_section": None,
        "section_chain": []
    }
    
    # Find all sectPr blocks in document.xml
    sect_blocks = extract_document_sectpr_blocks(document_xml)
    
    # Parse relationships for header/footer refs
    rels_part_name = "word/_rels/document.xml.rels"
    rels_xml = _read_xml_part(extract_dir, rels_part_name)
    document_relationships = (
        _parse_relationships_xml(rels_xml, rels_part_name) if rels_xml else []
    )
    relationships_by_id = {rel.get("Id"): rel for rel in document_relationships}
    
    for idx, sect in enumerate(sect_blocks):
        sect_info = _parse_sectpr(sect, rels_xml, idx)
        for kind, reference_group in (
            ("header", sect_info["header_refs"]),
            ("footer", sect_info["footer_refs"]),
        ):
            for reference_type, rel_id in reference_group.items():
                if rel_id is None:
                    continue
                rel = relationships_by_id.get(rel_id)
                if rel is None:
                    raise ValueError(
                        f"Section {idx} {kind} reference {reference_type!r} uses "
                        f"missing relationship {rel_id!r} in {rels_part_name}"
                    )
                rel_type = rel.get("Type") or ""
                if rel.get("TargetMode") == "External" or not rel_type.endswith(f"/{kind}"):
                    raise ValueError(
                        f"Section {idx} {kind} reference {reference_type!r} uses "
                        f"relationship {rel_id!r} with unsupported Type={rel_type!r} "
                        f"or TargetMode={rel.get('TargetMode')!r}"
                    )
                target = rel.get("Target")
                assert target is not None  # validated by _parse_relationships_xml
                target_path, package_part = _resolve_internal_relationship_target(
                    extract_dir,
                    "word/document.xml",
                    target,
                )
                if not target_path.is_file():
                    raise ValueError(
                        f"Section {idx} {kind} relationship {rel_id!r} targets "
                        f"missing package part {package_part!r}"
                    )
        result["section_chain"].append(sect_info)
    
    # The last sectPr is typically the "default" section
    if result["section_chain"]:
        result["default_section"] = result["section_chain"][-1]
    
    return result


def _parse_sectpr(sect_xml: str, rels_xml: Optional[str], section_index: int) -> Dict[str, Any]:
    """Parse a single sectPr block into structured data."""
    info = {
        "section_index": section_index,
        "sectPr": _canonicalize(sect_xml),
        "page_size": {},
        "page_margins": {},
        "columns": {},
        "doc_grid": None,
        "header_refs": {"default": None, "first": None, "even": None},
        "footer_refs": {"default": None, "first": None, "even": None},
    }
    
    # Page size
    pg_sz_m = re.search(r'<w:pgSz\s+([^>]+)', sect_xml)
    if pg_sz_m:
        attrs = pg_sz_m.group(1)
        w_m = re.search(r'w:w="(\d+)"', attrs)
        h_m = re.search(r'w:h="(\d+)"', attrs)
        orient_m = re.search(r'w:orient="([^"]+)"', attrs)
        info["page_size"] = {
            "w": int(w_m.group(1)) if w_m else None,
            "h": int(h_m.group(1)) if h_m else None,
            "orient": orient_m.group(1) if orient_m else "portrait"
        }
    
    # Page margins
    pg_mar_m = re.search(r'<w:pgMar\s+([^>]+)', sect_xml)
    if pg_mar_m:
        attrs = pg_mar_m.group(1)
        for margin in ["top", "right", "bottom", "left", "header", "footer", "gutter"]:
            m = re.search(rf'w:{margin}="([+-]?\d+)"', attrs)
            info["page_margins"][margin] = int(m.group(1)) if m else 0
    
    # Columns
    cols_m = re.search(r'<w:cols\s+([^/]*)', sect_xml)
    if cols_m:
        attrs = cols_m.group(1)
        num_m = re.search(r'w:num="(\d+)"', attrs)
        space_m = re.search(r'w:space="(\d+)"', attrs)
        info["columns"] = {
            "num": int(num_m.group(1)) if num_m else 1,
            "space": int(space_m.group(1)) if space_m else 720,
            "sep": _on_off_attribute(attrs, "sep"),
        }
    
    # Document grid
    doc_grid = _extract_first_block(sect_xml, "docGrid")
    if doc_grid:
        info["doc_grid"] = _canonicalize(doc_grid)
    
    # Header/footer reference elements are normally self-closing, but paired
    # empty forms are also legal XML and Word accepts both quote styles.  Only
    # the opening token carries data, so this deliberately recognizes either
    # serialization without trying to parse a namespace-less XML fragment.
    for kind in ("header", "footer"):
        refs = info[f"{kind}_refs"]
        for match in re.finditer(rf"<w:{kind}Reference\b([^>]*)>", sect_xml):
            attrs = match.group(1)
            reference_type = _qualified_attribute(attrs, "w:type")
            rel_id = _qualified_attribute(attrs, "r:id")
            if reference_type in refs and rel_id:
                refs[reference_type] = rel_id
    
    return info


def extract_headers_footers(extract_dir: Path) -> Dict[str, Any]:
    """Extract header/footer parts named by document relationships.

    OPC does not require these parts to be named ``word/header1.xml`` or to
    live directly under ``word``.  The document relationships are therefore
    authoritative; filesystem globs can both miss legal custom paths and pick
    up orphan parts that the document never uses.
    """
    extract_dir = Path(extract_dir)
    content_type_defaults, content_type_overrides = _load_content_types(extract_dir)
    total_media_bytes = 0

    def _read_bounded_media(media_path: Path, relationship_label: str) -> bytes:
        nonlocal total_media_bytes

        size = media_path.stat().st_size
        if size > MAX_HEADER_FOOTER_MEDIA_BYTES:
            raise ValueError(
                f"Header/footer media {relationship_label} is {size} bytes; "
                f"limit is {MAX_HEADER_FOOTER_MEDIA_BYTES} bytes"
            )
        if total_media_bytes + size > MAX_HEADER_FOOTER_MEDIA_TOTAL_BYTES:
            raise ValueError(
                "Total header/footer media exceeds "
                f"{MAX_HEADER_FOOTER_MEDIA_TOTAL_BYTES} bytes"
            )

        with media_path.open("rb") as stream:
            data = stream.read(MAX_HEADER_FOOTER_MEDIA_BYTES + 1)
        if len(data) > MAX_HEADER_FOOTER_MEDIA_BYTES:
            raise ValueError(
                f"Header/footer media {relationship_label} grew beyond "
                f"{MAX_HEADER_FOOTER_MEDIA_BYTES} bytes while being read"
            )
        if total_media_bytes + len(data) > MAX_HEADER_FOOTER_MEDIA_TOTAL_BYTES:
            raise ValueError(
                "Total header/footer media exceeds "
                f"{MAX_HEADER_FOOTER_MEDIA_TOTAL_BYTES} bytes"
            )
        total_media_bytes += len(data)
        return data

    def _extract_media_from_rels(
        part_name: str,
        rels_part_name: str,
        rels_xml: Optional[str],
        part_xml: str,
    ) -> List[Dict[str, str]]:
        references = _relationship_references(part_xml, part_name)
        if not rels_xml:
            if references:
                missing = sorted({rel_id for rel_id, _attribute in references})
                raise ValueError(
                    f"{part_name} references missing relationships {missing}; "
                    f"{rels_part_name} does not exist"
                )
            return []

        relationships = _parse_relationships_xml(rels_xml, rels_part_name)
        relationships_by_id = {rel.get("Id"): rel for rel in relationships}
        missing = sorted(
            {
                rel_id
                for rel_id, _attribute in references
                if rel_id not in relationships_by_id
            }
        )
        if missing:
            raise ValueError(
                f"{part_name} references relationship IDs absent from "
                f"{rels_part_name}: {missing}"
            )

        for rel_id, attribute_name in references:
            rel = relationships_by_id[rel_id]
            rel_type = rel.get("Type") or ""
            if attribute_name in {"r:embed", "r:link"} and not rel_type.endswith("/image"):
                raise ValueError(
                    f"Unsupported {attribute_name} relationship {rel_id!r} in {part_name}: "
                    f"expected an image relationship, got {rel_type!r}"
                )

        media_items: List[Dict[str, str]] = []
        for rel in relationships:
            # External image links may point to file://, UNC, or network URLs.
            # Preserve the relationship XML, but never dereference the target.
            if rel.get("TargetMode") == "External":
                continue

            rel_type = rel.get("Type") or ""
            if not rel_type.endswith("/image"):
                raise ValueError(
                    f"Unsupported internal header/footer relationship in {rels_part_name}: "
                    f"Id={rel.get('Id')!r}, Type={rel_type!r}. "
                    "Only internal images and external relationships are supported."
                )

            rel_id = rel.get("Id")
            target = rel.get("Target")
            assert rel_id is not None and target is not None  # validated above
            media_path, package_part = _resolve_internal_relationship_target(
                extract_dir,
                part_name,
                target,
            )
            if not media_path.exists() or not media_path.is_file():
                raise ValueError(
                    f"Header/footer image relationship {rel_id!r} in {rels_part_name} "
                    f"targets missing package part {package_part!r}"
                )

            word_root = PurePosixPath("word")
            package_part_path = PurePosixPath(package_part)
            try:
                registry_target = package_part_path.relative_to(word_root).as_posix()
            except ValueError:
                registry_target = package_part

            data = _read_bounded_media(
                media_path,
                f"{rels_part_name}:{rel_id}",
            )
            media_items.append(
                {
                    "rel_id": rel_id,
                    "target": registry_target,
                    "content_type": _content_type_for_part(
                        package_part,
                        content_type_defaults,
                        content_type_overrides,
                    ),
                    "data_base64": base64.b64encode(data).decode("ascii"),
                }
            )

        return media_items

    result = {"headers": [], "footers": []}
    header_footer_media = set()

    # Resolve each part from the validated document relationship.  Preserve
    # one registry entry per relationship ID so every section reference can be
    # mapped even when a package legally has two IDs for the same part.
    rels_xml = _read_xml_part(extract_dir, "word/_rels/document.xml.rels")
    document_parts: List[Tuple[str, str, str, Path]] = []
    if rels_xml:
        for rel in _parse_relationships_xml(
            rels_xml,
            "word/_rels/document.xml.rels",
        ):
            rel_type = rel.get("Type") or ""
            if not rel_type.endswith(("/header", "/footer")):
                continue

            rel_id = rel.get("Id")
            target = rel.get("Target")
            assert rel_id is not None and target is not None  # validated above
            kind = "header" if rel_type.endswith("/header") else "footer"
            if rel.get("TargetMode") == "External":
                raise ValueError(
                    f"Unsupported external {kind} relationship {rel_id!r} in "
                    "word/_rels/document.xml.rels"
                )
            part_path, package_part = _resolve_internal_relationship_target(
                extract_dir,
                "word/document.xml",
                target,
            )
            if not part_path.is_file():
                raise ValueError(
                    f"{kind.title()} relationship {rel_id!r} targets missing "
                    f"package part {package_part!r}"
                )
            if PurePosixPath(package_part).suffix.lower() != ".xml":
                raise ValueError(
                    f"Unsupported {kind} part name {package_part!r}: expected an XML part"
                )
            document_parts.append((kind, rel_id, package_part, part_path))

    kinds_by_part: Dict[str, str] = {}
    for kind, rel_id, part_name, _part_path in document_parts:
        prior_kind = kinds_by_part.setdefault(part_name.casefold(), kind)
        if prior_kind != kind:
            raise ValueError(
                f"Package part {part_name!r} is targeted as both a header and footer "
                f"(including relationship {rel_id!r})"
            )

    captured_parts: Dict[str, Dict[str, Any]] = {}
    for kind, rel_id, part_name, part_path in sorted(
        document_parts,
        key=lambda item: (item[0], item[2].casefold(), item[1]),
    ):
        captured = captured_parts.get(part_name)
        if captured is None:
            xml_content = read_xml_text(part_path)
            rels_part_name = _relationship_part_name(part_name)
            rels_content = _read_xml_part(extract_dir, rels_part_name)
            media_entries = _extract_media_from_rels(
                part_name,
                rels_part_name,
                rels_content,
                xml_content,
            )
            for media_entry in media_entries:
                header_footer_media.add(media_entry["target"])
            captured = {
                "part_name": part_name,
                "xml": _canonicalize(xml_content),
                "rels_part_name": rels_part_name if rels_content else None,
                "rels_xml": (
                    _canonicalize(rels_content, strip_rsid=False)
                    if rels_content
                    else None
                ),
                "media": media_entries,
            }
            captured_parts[part_name] = captured

        entry = dict(captured)
        entry["media"] = [dict(item) for item in captured["media"]]
        entry["rel_id"] = rel_id
        result[f"{kind}s"].append(entry)

    result["header_footer_media"] = sorted(header_footer_media)
    
    return result


def extract_numbering(extract_dir: Path) -> Dict[str, Any]:
    """
    Extract numbering.xml which defines list behaviors.
    This is critical for correct list continuation.
    """
    numbering_xml = _read_xml_part(extract_dir, "word/numbering.xml")
    
    result = {
        "numbering_xml": None,
        "abstract_nums": [],
        "nums": []
    }
    
    if not numbering_xml:
        return result
    
    result["numbering_xml"] = _canonicalize(numbering_xml)
    
    # Extract individual abstractNum definitions using depth-tracking scanner
    for block in _extract_all_blocks(numbering_xml, "abstractNum"):
        id_m = re.search(r'w:abstractNumId="(\d+)"', block)
        result["abstract_nums"].append({
            "abstractNumId": int(id_m.group(1)) if id_m else -1,
            "xml": _canonicalize(block)
        })

    # Extract num definitions (map numId -> abstractNumId)
    for block in _extract_all_blocks(numbering_xml, "num"):
        id_m = re.search(r'w:numId="(\d+)"', block)
        num_id = int(id_m.group(1)) if id_m else -1

        abs_m = re.search(r'<w:abstractNumId\s+w:val="(\d+)"', block)
        abs_id = int(abs_m.group(1)) if abs_m else None

        result["nums"].append({
            "numId": num_id,
            "abstractNumId": abs_id,
            "xml": _canonicalize(block)
        })
    
    return result


def extract_fonts(extract_dir: Path) -> Dict[str, Any]:
    """Extract declared fonts while rejecting unsupported embedded font data.

    The Phase 1 bundle intentionally carries only ``fontTable.xml``.  Emitting
    ``w:embed*`` references without the relationship part and obfuscated font
    payload would create dangling references in Phase 2, so such a source must
    fail during capture instead of producing a superficially valid bundle.
    """
    font_xml = _read_xml_part(extract_dir, "word/fontTable.xml")
    if font_xml:
        references = _relationship_references(font_xml, "word/fontTable.xml")
        embed_elements = re.findall(
            r"<w:(embedRegular|embedBold|embedItalic|embedBoldItalic)\b",
            font_xml,
        )
        if embed_elements or references:
            details = sorted(set(embed_elements)) or [name for _rid, name in references]
            raise ValueError(
                "Embedded fonts are unsupported: word/fontTable.xml contains "
                f"relationship-backed font declarations {details}"
            )

    font_rels_part = "word/_rels/fontTable.xml.rels"
    font_rels_xml = _read_xml_part(extract_dir, font_rels_part)
    if font_rels_xml:
        relationships = _parse_relationships_xml(font_rels_xml, font_rels_part)
        if relationships:
            relationship_ids = sorted(rel.get("Id") for rel in relationships)
            raise ValueError(
                f"Embedded fonts are unsupported: {font_rels_part} contains "
                f"relationships {relationship_ids}"
            )

    return {
        "font_table_xml": _canonicalize(font_xml) if font_xml else None,
        "notes": {
            "captures_declared_fonts_only": True
        }
    }


def extract_relationships(extract_dir: Path) -> Dict[str, Any]:
    """Extract key relationship files."""
    result = {
        "relationships": [],
        "other_parts_passthrough": []
    }
    
    # document.xml.rels
    doc_rels = _read_xml_part(extract_dir, "word/_rels/document.xml.rels")
    if doc_rels:
        _parse_relationships_xml(doc_rels, "word/_rels/document.xml.rels")
        result["relationships"].append({
            "part": "word/document.xml",
            "rels_xml": _canonicalize(doc_rels, strip_rsid=False)
        })
    
    # Note other parts we preserve but don't parse
    if (extract_dir / "docProps" / "core.xml").exists():
        result["other_parts_passthrough"].append("docProps/core.xml")
    if (extract_dir / "docProps" / "app.xml").exists():
        result["other_parts_passthrough"].append("docProps/app.xml")
    
    return result


# ─────────────────────────────────────────────────────────────────────────────
# Main extraction orchestrator
# ─────────────────────────────────────────────────────────────────────────────

def extract_arch_template_registry(
    extract_dir: Path,
    source_docx_path: Optional[Path] = None,
) -> Dict[str, Any]:
    """
    Build the complete arch_template_registry.json from an extracted DOCX folder.
    
    Args:
        extract_dir: Path to extracted DOCX folder (contains word/, _rels/, etc.)
        source_docx_path: Optional path to original .docx for metadata
    
    Returns:
        Complete registry dict following the arch_template_registry schema
    """
    extract_dir = Path(extract_dir)
    
    # Read required parts
    styles_xml = _read_xml_part(extract_dir, "word/styles.xml")
    document_xml = _read_xml_part(extract_dir, "word/document.xml")
    
    if not styles_xml:
        raise FileNotFoundError(f"word/styles.xml not found in {extract_dir}")
    if not document_xml:
        raise FileNotFoundError(f"word/document.xml not found in {extract_dir}")
    
    # Build metadata
    meta = {
        "schema_version": SCHEMA_VERSION,
        "source_docx": {
            "filename": source_docx_path.name if source_docx_path else extract_dir.name,
            "sha256": sha256_file(source_docx_path) if source_docx_path and source_docx_path.exists() else None,
            "extracted_utc": datetime.now(timezone.utc).isoformat()
        }
    }
    
    # Build complete registry
    registry = {
        "meta": meta,
        "package_inventory": extract_package_inventory(extract_dir),
        "doc_defaults": extract_doc_defaults(styles_xml),
        "styles": extract_styles_section(styles_xml),
        "theme": extract_theme(extract_dir),
        "settings": extract_settings(extract_dir),
        "page_layout": extract_page_layout(document_xml, extract_dir),
        "headers_footers": extract_headers_footers(extract_dir),
        "numbering": extract_numbering(extract_dir),
        "fonts": extract_fonts(extract_dir),
        "custom_xml": extract_relationships(extract_dir),
        "capture_policy": {
            # XML fragments are source-derived but deliberately normalized by
            # removing volatile rsid attributes and proofing markers.
            "store_raw_xml_blocks": False,
            "store_normalized_xml_blocks": True,
            "canonicalize_whitespace": False,
            "strip_rsids": True,
            "strip_proofing": True
        }
    }
    
    return registry

