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

Usage:
    python arch_env_extractor.py ARCH_TEMPLATE.docx
    python arch_env_extractor.py ARCH_TEMPLATE.docx --output arch_template_registry.json
    python arch_env_extractor.py --extract-dir MySpec_extracted

The output JSON follows the arch_template_registry schema (v1.0.0).
"""

from __future__ import annotations

import argparse
import hashlib
import json
import re
import zipfile
from dataclasses import dataclass, field
from datetime import datetime, timezone
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple
import xml.etree.ElementTree as ET


# ─────────────────────────────────────────────────────────────────────────────
# Constants / Namespaces
# ─────────────────────────────────────────────────────────────────────────────

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"
R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"

SCHEMA_VERSION = "1.0.0"


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
        return p.read_text(encoding="utf-8")
    return None


def _read_xml_part_bytes(extract_dir: Path, internal_path: str) -> Optional[bytes]:
    """Read raw bytes from extracted DOCX folder."""
    p = extract_dir / internal_path
    if p.exists():
        return p.read_bytes()
    return None


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
    """Apply standard cleanup to XML text."""
    if strip_rsid:
        xml_text = _strip_rsids(xml_text)
    if strip_proof:
        xml_text = _strip_proofing(xml_text)
    return xml_text


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
        qformat = "<w:qFormat" in block
        semi_hidden = "<w:semiHidden" in block
        unhide_when_used = "<w:unhideWhenUsed" in block
        locked = "<w:locked" in block
        
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
    sect_blocks = _extract_all_blocks(document_xml, "sectPr")
    
    # Parse relationships for header/footer refs
    rels_xml = _read_xml_part(extract_dir, "word/_rels/document.xml.rels")
    
    for idx, sect in enumerate(sect_blocks):
        sect_info = _parse_sectpr(sect, rels_xml, idx)
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
            m = re.search(rf'w:{margin}="(\d+)"', attrs)
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
            "sep": "w:sep" in attrs
        }
    
    # Document grid
    doc_grid = _extract_first_block(sect_xml, "docGrid")
    if doc_grid:
        info["doc_grid"] = _canonicalize(doc_grid)
    
    # Header/footer references
    for hdr_m in re.finditer(r'<w:headerReference\s+([^>]+)/>', sect_xml):
        attrs = hdr_m.group(1)
        type_m = re.search(r'w:type="([^"]+)"', attrs)
        rid_m = re.search(r'r:id="([^"]+)"', attrs)
        if type_m and rid_m:
            hdr_type = type_m.group(1)
            if hdr_type in info["header_refs"]:
                info["header_refs"][hdr_type] = rid_m.group(1)
    
    for ftr_m in re.finditer(r'<w:footerReference\s+([^>]+)/>', sect_xml):
        attrs = ftr_m.group(1)
        type_m = re.search(r'w:type="([^"]+)"', attrs)
        rid_m = re.search(r'r:id="([^"]+)"', attrs)
        if type_m and rid_m:
            ftr_type = type_m.group(1)
            if ftr_type in info["footer_refs"]:
                info["footer_refs"][ftr_type] = rid_m.group(1)
    
    return info


def extract_headers_footers(extract_dir: Path) -> Dict[str, Any]:
    """Extract all header and footer parts."""
    result = {
        "headers": [],
        "footers": []
    }
    
    word_dir = extract_dir / "word"
    
    # Read relationships to map rId to part names
    rels_xml = _read_xml_part(extract_dir, "word/_rels/document.xml.rels")
    rid_to_target = {}
    if rels_xml:
        for m in re.finditer(r'<Relationship[^>]+Id="([^"]+)"[^>]+Target="([^"]+)"', rels_xml):
            rid_to_target[m.group(1)] = m.group(2)
    
    # Collect headers
    for hdr_path in sorted(word_dir.glob("header*.xml")):
        part_name = f"word/{hdr_path.name}"
        xml_content = hdr_path.read_text(encoding="utf-8")
        
        # Find the rId for this part
        rel_id = None
        for rid, target in rid_to_target.items():
            if target == hdr_path.name or target == part_name:
                rel_id = rid
                break
        
        result["headers"].append({
            "part_name": part_name,
            "rel_id": rel_id,
            "xml": _canonicalize(xml_content)
        })
    
    # Collect footers
    for ftr_path in sorted(word_dir.glob("footer*.xml")):
        part_name = f"word/{ftr_path.name}"
        xml_content = ftr_path.read_text(encoding="utf-8")
        
        rel_id = None
        for rid, target in rid_to_target.items():
            if target == ftr_path.name or target == part_name:
                rel_id = rid
                break
        
        result["footers"].append({
            "part_name": part_name,
            "rel_id": rel_id,
            "xml": _canonicalize(xml_content)
        })
    
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
    """Extract fontTable.xml which declares fonts used in the document."""
    font_xml = _read_xml_part(extract_dir, "word/fontTable.xml")
    
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
            "store_raw_xml_blocks": True,
            "canonicalize_whitespace": True,
            "strip_rsids": True,
            "strip_proofing": True
        }
    }
    
    return registry


def extract_docx_to_dir(docx_path: Path, extract_dir: Path) -> None:
    """Extract a .docx file to a directory."""
    import shutil
    
    if extract_dir.exists():
        shutil.rmtree(extract_dir)
    extract_dir.mkdir(parents=True, exist_ok=True)
    
    with zipfile.ZipFile(docx_path, "r") as zf:
        zf.extractall(extract_dir)


# ─────────────────────────────────────────────────────────────────────────────
# CLI
# ─────────────────────────────────────────────────────────────────────────────

def main():
    parser = argparse.ArgumentParser(
        description="Extract arch_template_registry.json from an architect Word template"
    )
    parser.add_argument(
        "input",
        nargs="?",
        help="Path to .docx file OR extracted folder"
    )
    parser.add_argument(
        "--extract-dir",
        help="Use existing extracted folder (skip extraction)"
    )
    parser.add_argument(
        "--output", "-o",
        default=None,
        help="Output path for arch_template_registry.json (default: <extract_dir>/arch_template_registry.json)"
    )
    parser.add_argument(
        "--pretty",
        action="store_true",
        default=True,
        help="Pretty-print JSON output (default: True)"
    )
    
    args = parser.parse_args()
    
    # Determine extraction directory
    source_docx = None
    
    if args.extract_dir:
        extract_dir = Path(args.extract_dir)
        if not extract_dir.exists():
            raise FileNotFoundError(f"Extract directory not found: {extract_dir}")
    elif args.input:
        input_path = Path(args.input)
        if input_path.is_file() and input_path.suffix.lower() == ".docx":
            source_docx = input_path
            extract_dir = Path(f"{input_path.stem}_extracted")
            print(f"Extracting {input_path} to {extract_dir}...")
            extract_docx_to_dir(input_path, extract_dir)
        elif input_path.is_dir():
            extract_dir = input_path
        else:
            raise ValueError(f"Input must be a .docx file or extracted folder: {input_path}")
    else:
        parser.print_help()
        return
    
    # Validate extraction directory
    if not (extract_dir / "word" / "styles.xml").exists():
        raise FileNotFoundError(
            f"Invalid extraction directory: {extract_dir}\n"
            "Expected word/styles.xml to exist."
        )
    
    # Extract registry
    print(f"Extracting environment from: {extract_dir}")
    registry = extract_arch_template_registry(extract_dir, source_docx)

    # Validate before writing
    from phase1_validator import validate_template_registry
    print("Validating template registry...")
    validate_template_registry(registry)

    # Determine output path
    if args.output:
        out_path = Path(args.output)
    else:
        out_path = extract_dir / "arch_template_registry.json"
    
    # Write output
    out_path.parent.mkdir(parents=True, exist_ok=True)
    indent = 2 if args.pretty else None
    out_path.write_text(json.dumps(registry, indent=indent), encoding="utf-8")
    
    print(f"arch_template_registry.json written: {out_path}")
    
    # Summary
    inv = registry["package_inventory"]
    print("\nPackage inventory:")
    for k, v in inv.items():
        status = "✓" if v else "✗"
        print(f"  {status} {k}")
    
    n_styles = len(registry["styles"]["style_defs"])
    n_abstract = len(registry["numbering"]["abstract_nums"])
    n_nums = len(registry["numbering"]["nums"])
    n_headers = len(registry["headers_footers"]["headers"])
    n_footers = len(registry["headers_footers"]["footers"])
    
    print(f"\nCaptured:")
    print(f"  {n_styles} style definitions")
    print(f"  {n_abstract} abstract numbering definitions")
    print(f"  {n_nums} numbering instances")
    print(f"  {n_headers} header parts")
    print(f"  {n_footers} footer parts")


if __name__ == "__main__":
    main()