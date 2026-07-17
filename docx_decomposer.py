"""
PHASE 1 — Architect Template Normalization (Lean)

Library module providing pipeline functions for the Phase 1 DOCX CSI Normalizer.
Called by gui.py — no CLI entry point.

Purpose
- Run ONLY on an architect's DOCX template (e.g. MySpec.docx).
- Produce formal contract artifacts for Phase 2:
  - arch_style_registry.json
  - arch_template_registry.json

Hard invariants (enforced):
- Pixel-identical output (we DO NOT emit a reconstructed docx in Phase 1)
- Never modify headers/footers
- Never modify w:sectPr
- Never modify numbering definitions (word/numbering.xml)
- Never reconstruct DOCX XML (no parse/re-serialize of document.xml)
- When applying, only insert/replace <w:pStyle> in document.xml
- Styles are derived locally from exemplar paragraphs (LLM never specifies pPr/rPr)

Note:
- This module intentionally does NOT generate an "analysis markdown" report and does NOT reconstruct a docx.
"""

from __future__ import annotations

import hashlib
import json
import re
import shutil
import stat
import zipfile
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Dict, List, Optional, Set, Tuple
import html

import xml.etree.ElementTree as ET  # only used for styles.xml name lookup + optional catalogs
from paragraph_rules import (
    RE_SECTION_WITH_TITLE,
    compute_skip_reason,
    detect_numbering_role,
    detect_role_signal,
    infer_expected_roles,
    is_classifiable_paragraph,
    is_role_candidate_paragraph,
)
from ooxml_text import prepare_xml_text_for_utf8, read_xml_text


# -----------------------------
# Utilities
# -----------------------------

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"

ROLE_TO_ARCH_STYLE: Dict[str, str] = {
    "SectionID": "CSI_SectionID__ARCH",
    "SectionTitle": "CSI_SectionTitle__ARCH",
    "PART": "CSI_Part__ARCH",
    "ARTICLE": "CSI_Article__ARCH",
    "PARAGRAPH": "CSI_Paragraph__ARCH",
    "SUBPARAGRAPH": "CSI_Subparagraph__ARCH",
    "SUBSUBPARAGRAPH": "CSI_Subsubparagraph__ARCH",
    "END_OF_SECTION": "CSI_EndOfSection__ARCH",
}
ALLOWED_ROLES: Set[str] = set(ROLE_TO_ARCH_STYLE)

def sha256_bytes(b: bytes) -> str:
    return hashlib.sha256(b).hexdigest()

def sha256_text(s: str) -> str:
    return sha256_bytes(s.encode("utf-8"))

def _q(tag: str) -> str:
    return f"{{{W_NS}}}{tag}"

def _get_attr(elem: ET.Element, local_name: str) -> Optional[str]:
    return elem.get(f"{{{W_NS}}}{local_name}")


# -----------------------------
# Phase 1 stability snapshot
# -----------------------------

@dataclass(frozen=True)
class StabilitySnapshot:
    header_footer_hashes: Dict[str, str]
    sectpr_hash: str
    doc_rels_hash: str


@dataclass(frozen=True)
class ParagraphContext:
    in_table: bool
    contains_sectPr: bool


@dataclass(frozen=True)
class StructureSignal:
    paragraph_index: int
    text_role: Optional[str]
    numbering_source: str
    num_id: Optional[str]
    abstract_num_id: Optional[str]
    ilvl: Optional[str]
    num_fmt: Optional[str]
    lvl_text: Optional[str]
    family_key: Optional[Tuple[Optional[str], Optional[str]]]


def snapshot_headers_footers(extract_dir: Path) -> Dict[str, str]:
    wf = extract_dir / "word"
    hashes: Dict[str, str] = {}
    for p in sorted(wf.glob("header*.xml")) + sorted(wf.glob("footer*.xml")):
        rel = str(p.relative_to(extract_dir)).replace("\\", "/")
        hashes[rel] = sha256_bytes(p.read_bytes())
    return hashes


def snapshot_doc_rels_hash(extract_dir: Path) -> str:
    rels_path = extract_dir / "word" / "_rels" / "document.xml.rels"
    if not rels_path.exists():
        return ""
    return sha256_bytes(rels_path.read_bytes())


def extract_sectpr_block(document_xml: str) -> str:
    # pragmatic stability check: exact raw text blocks
    blocks = re.findall(r"(<w:sectPr[\s\S]*?</w:sectPr>)", document_xml)
    return "\n".join(blocks)


def snapshot_stability(extract_dir: Path) -> StabilitySnapshot:
    doc_path = extract_dir / "word" / "document.xml"
    doc_text = read_xml_text(doc_path)
    sectpr = extract_sectpr_block(doc_text)
    return StabilitySnapshot(
        header_footer_hashes=snapshot_headers_footers(extract_dir),
        sectpr_hash=sha256_text(sectpr),
        doc_rels_hash=snapshot_doc_rels_hash(extract_dir),
    )


def verify_stability(extract_dir: Path, snap: StabilitySnapshot) -> None:
    current_hf = snapshot_headers_footers(extract_dir)
    if current_hf != snap.header_footer_hashes:
        changed = []
        all_keys = set(current_hf.keys()) | set(snap.header_footer_hashes.keys())
        for k in sorted(all_keys):
            if current_hf.get(k) != snap.header_footer_hashes.get(k):
                changed.append(k)
        raise ValueError(f"Header/footer stability check FAILED. Changed: {changed}")

    doc_text = read_xml_text(extract_dir / "word" / "document.xml")
    current_sectpr = extract_sectpr_block(doc_text)
    if sha256_text(current_sectpr) != snap.sectpr_hash:
        raise ValueError("Section properties (w:sectPr) stability check FAILED.")

    current_rels = snapshot_doc_rels_hash(extract_dir)
    if current_rels != snap.doc_rels_hash:
        raise ValueError("document.xml.rels stability check FAILED (can break header/footer binding).")


# -----------------------------
# DOCX extraction (workspace only)
# -----------------------------

MAX_PACKAGE_ENTRIES = 10_000
MAX_PACKAGE_UNCOMPRESSED_BYTES = 512 * 1024 * 1024
MAX_PACKAGE_PART_BYTES = 128 * 1024 * 1024
MAX_COMPRESSION_RATIO = 1_000


def _safe_package_member_path(extract_dir: Path, member_name: str) -> Path:
    if not member_name or "\x00" in member_name or "\\" in member_name:
        raise ValueError(f"Unsafe DOCX package member name: {member_name!r}")
    if member_name.startswith(("/", "//")) or re.match(r"^[A-Za-z]:", member_name):
        raise ValueError(f"Unsafe absolute DOCX package member: {member_name!r}")
    parts = member_name.rstrip("/").split("/")
    if any(part in {"", ".", ".."} for part in parts):
        raise ValueError(f"Unsafe DOCX package member traversal: {member_name!r}")
    reserved_windows_names = {"CON", "PRN", "AUX", "NUL"} | {
        f"{prefix}{number}" for prefix in ("COM", "LPT") for number in range(1, 10)
    }
    for part in parts:
        stem = part.split(".", 1)[0].upper()
        if ":" in part or part.endswith((" ", ".")) or stem in reserved_windows_names:
            raise ValueError(f"Unsafe Windows DOCX package member: {member_name!r}")
    destination = (extract_dir / Path(*parts)).resolve()
    root = extract_dir.resolve()
    try:
        destination.relative_to(root)
    except ValueError as exc:
        raise ValueError(f"DOCX package member escapes extraction root: {member_name!r}") from exc
    return destination


def extract_docx(docx_path: Path, extract_dir: Path, *, overwrite: bool = False) -> None:
    if extract_dir.exists():
        if not overwrite:
            raise FileExistsError(f"Extraction target already exists: {extract_dir}")
        # OneDrive-safe deletion: retry with delay if locked
        import time
        max_retries = 3
        for attempt in range(max_retries):
            try:
                shutil.rmtree(extract_dir)
                break
            except PermissionError:
                if attempt < max_retries - 1:
                    print(f"Folder locked (OneDrive?), retrying in 2s... ({attempt + 1}/{max_retries})")
                    time.sleep(2)
                else:
                    # Last resort: rename instead of delete
                    import uuid
                    backup = extract_dir.with_name(f"{extract_dir.name}_old_{uuid.uuid4().hex[:8]}")
                    print(f"Cannot delete {extract_dir}, renaming to {backup}")
                    extract_dir.rename(backup)
    
    extract_dir.mkdir(parents=True, exist_ok=True)
    try:
        with zipfile.ZipFile(docx_path, "r") as archive:
            entries = archive.infolist()
            if len(entries) > MAX_PACKAGE_ENTRIES:
                raise ValueError(
                    f"DOCX package has {len(entries)} entries; limit is {MAX_PACKAGE_ENTRIES}"
                )
            total_size = sum(entry.file_size for entry in entries)
            if total_size > MAX_PACKAGE_UNCOMPRESSED_BYTES:
                raise ValueError(
                    f"DOCX package expands to {total_size} bytes; limit is {MAX_PACKAGE_UNCOMPRESSED_BYTES}"
                )

            seen_names: Set[str] = set()
            for entry in entries:
                # ZipInfo.filename may normalize backslashes; orig_filename is
                # the only trustworthy value for rejecting a crafted archive.
                raw_member_name = entry.orig_filename
                _safe_package_member_path(extract_dir, raw_member_name)
                normalized_name = entry.filename.casefold()
                if normalized_name in seen_names:
                    raise ValueError(f"DOCX package contains duplicate member: {entry.filename!r}")
                seen_names.add(normalized_name)
                destination = _safe_package_member_path(extract_dir, entry.filename)
                unix_mode = (entry.external_attr >> 16) & 0xFFFF
                if stat.S_ISLNK(unix_mode):
                    raise ValueError(f"DOCX package contains a symbolic link: {entry.filename!r}")
                if entry.file_size > MAX_PACKAGE_PART_BYTES:
                    raise ValueError(
                        f"DOCX package member {entry.filename!r} is {entry.file_size} bytes; "
                        f"per-part limit is {MAX_PACKAGE_PART_BYTES}"
                    )
                if entry.file_size and entry.compress_size == 0:
                    raise ValueError(f"DOCX package member has invalid compressed size: {entry.filename!r}")
                if entry.compress_size and entry.file_size / entry.compress_size > MAX_COMPRESSION_RATIO:
                    raise ValueError(f"DOCX package member has suspicious compression ratio: {entry.filename!r}")

                if entry.is_dir():
                    destination.mkdir(parents=True, exist_ok=True)
                    continue
                destination.parent.mkdir(parents=True, exist_ok=True)
                with archive.open(entry, "r") as source, destination.open("xb") as target:
                    shutil.copyfileobj(source, target, length=1024 * 1024)

        required_parts = [
            extract_dir / "[Content_Types].xml",
            extract_dir / "_rels" / ".rels",
            extract_dir / "word" / "document.xml",
            extract_dir / "word" / "styles.xml",
        ]
        missing_parts = [str(path.relative_to(extract_dir)) for path in required_parts if not path.is_file()]
        if missing_parts:
            raise ValueError(f"DOCX package is missing required parts: {missing_parts}")
    except Exception:
        # Never leave a partially extracted tree that a later run could mistake
        # for a valid template.
        shutil.rmtree(extract_dir, ignore_errors=True)
        raise


# -----------------------------
# Slim bundle construction (LLM input)
# -----------------------------

def iter_paragraph_xml_blocks(document_xml_text: str):
    # Match both paired <w:p>...</w:p> and self-closing <w:p ... />
    for m in re.finditer(r"(<w:p\b(?:[^>]*/>|[\s\S]*?</w:p>))", document_xml_text):
        yield m.start(), m.end(), m.group(1)


def paragraph_text_from_block(p_xml: str) -> str:
    # Deleted/moved-from text and field instructions are not visible paragraph
    # content and must not influence semantic classification.
    visible = re.sub(r"<w:(?:del|moveFrom)\b[^>]*>[\s\S]*?</w:(?:del|moveFrom)>", "", p_xml)
    visible = re.sub(r"<w:instrText\b[^>]*>[\s\S]*?</w:instrText>", "", visible)
    # Tabs and explicit line breaks separate words in Word even though they do
    # not live inside w:t nodes.  The old extractor concatenated either side.
    separator_token = "\ue000"
    visible = re.sub(r"<w:(?:tab|br|cr)\b[^>]*/>", separator_token, visible)
    visible = re.sub(r"<w:noBreakHyphen\b[^>]*/>", "-", visible)
    visible = re.sub(r"<w:softHyphen\b[^>]*/>", "", visible)
    pieces = re.findall(rf"<w:t\b[^>]*>([\s\S]*?)</w:t>|({separator_token})", visible)
    if not pieces:
        return ""
    joined = html.unescape("".join(text if text else " " for text, _separator in pieces))
    joined = re.sub(r"\s+", " ", joined).strip()
    return joined


def paragraph_contains_sectpr(p_xml: str) -> bool:
    return "<w:sectPr" in p_xml


def paragraph_pstyle_from_block(p_xml: str) -> Optional[str]:
    m = re.search(r"<w:pStyle\b[^>]*w:val=\"([^\"]+)\"", p_xml)
    return m.group(1) if m else None


def paragraph_numpr_from_block(p_xml: str) -> Dict[str, Optional[str]]:
    numId = None
    ilvl = None
    m1 = re.search(r"<w:numId\b[^>]*w:val=\"([^\"]+)\"", p_xml)
    m2 = re.search(r"<w:ilvl\b[^>]*w:val=\"([^\"]+)\"", p_xml)
    if m1:
        numId = m1.group(1)
    if m2:
        ilvl = m2.group(1)
    return {"numId": numId, "ilvl": ilvl}


def paragraph_ppr_hints_from_block(p_xml: str) -> Dict[str, Any]:
    # lightweight hints (alignment + ind + spacing)
    hints: Dict[str, Any] = {}
    m = re.search(r"<w:jc\b[^>]*w:val=\"([^\"]+)\"", p_xml)
    if m:
        hints["jc"] = m.group(1)

    ind = {}
    for k in ["left", "right", "firstLine", "hanging"]:
        m2 = re.search(rf"<w:ind\b[^>]*w:{k}=\"([^\"]+)\"", p_xml)
        if m2:
            ind[k] = m2.group(1)
    if ind:
        hints["ind"] = ind

    spacing = {}
    for k in ["before", "after", "line", "lineRule"]:
        m3 = re.search(rf"<w:spacing\b[^>]*w:{k}=\"([^\"]+)\"", p_xml)
        if m3:
            spacing[k] = m3.group(1)
    if spacing:
        hints["spacing"] = spacing

    return hints


def _read_on_off_tag(xml: str, tag: str) -> Optional[bool]:
    """Read WordprocessingML on/off property semantics for a given tag.

    Returns:
      - None if the tag is absent
      - True/False when present, honoring explicit w:val values
    """
    m = re.search(rf"<w:{tag}\b([^>]*)/?>[\s\S]*?(?:</w:{tag}>)?", xml)
    if not m:
        return None

    attrs = m.group(1) or ""
    val_match = re.search(r'w:val="([^"]+)"', attrs)
    if not val_match:
        return True

    norm = val_match.group(1).strip().lower()
    if norm in {"false", "0", "off", "none"}:
        return False
    return True


def paragraph_rpr_hints_from_block(p_xml: str) -> Dict[str, Any]:
    """Extract lightweight run-property signals for classification."""
    rpr_inner = extract_paragraph_rpr_inner(p_xml)
    if not rpr_inner:
        return {}

    hints: Dict[str, Any] = {}

    for key, tag in [("bold", "b"), ("italic", "i"), ("caps", "caps"), ("underline", "u")]:
        val = _read_on_off_tag(rpr_inner, tag)
        if val is not None:
            hints[key] = val

    sz = re.search(r"<w:sz\b[^>]*w:val=\"([^\"]+)\"", rpr_inner)
    if sz:
        hints["sz"] = sz.group(1)

    font = re.search(r"<w:rFonts\b[^>]*w:ascii=\"([^\"]+)\"", rpr_inner)
    if font:
        hints["font"] = font.group(1)

    return hints


def _localname(tag: str) -> str:
    return tag.rsplit('}', 1)[-1]


def collect_paragraph_contexts(document_xml_path: Path) -> List[ParagraphContext]:
    tree = ET.parse(document_xml_path)
    root = tree.getroot()
    out: List[ParagraphContext] = []

    def walk(node: ET.Element, in_table: bool = False) -> None:
        local = _localname(node.tag)
        next_in_table = in_table or (local == "tbl")
        if local == "p":
            has_sect = node.find(f".//{_q('sectPr')}") is not None
            out.append(ParagraphContext(in_table=next_in_table, contains_sectPr=has_sect))
        for child in list(node):
            walk(child, in_table=next_in_table)

    walk(root)
    return out


def build_style_catalog(styles_xml_path: Path, used_style_ids: Set[str]) -> Dict[str, Any]:
    # Compact info for all paragraph styles (used + unused).
    tree = ET.parse(styles_xml_path)
    root = tree.getroot()

    styles_by_id: Dict[str, ET.Element] = {}
    for st in root.findall(f".//{_q('style')}"):
        sid = _get_attr(st, "styleId")
        if sid:
            styles_by_id[sid] = st

    catalog: Dict[str, Any] = {}
    for sid in sorted(styles_by_id.keys()):
        st = styles_by_id.get(sid)
        if st is None:
            continue
        if _get_attr(st, "type") != "paragraph":
            continue
        name_el = st.find(_q("name"))
        based_el = st.find(_q("basedOn"))
        catalog[sid] = {
            "styleId": sid,
            "type": "paragraph",
            "name": _get_attr(name_el, "val") if name_el is not None else None,
            "basedOn": _get_attr(based_el, "val") if based_el is not None else None,
            "default": str(_get_attr(st, "default") or "").lower() in {"1", "true", "on"},
            "in_use": sid in used_style_ids,
            "resolved_numPr": _find_style_chain_numpr(sid, styles_by_id),
        }
    return catalog


def build_numbering_catalog(numbering_xml_path: Path, used_num_ids: Set[str]) -> Dict[str, Any]:
    # Read-only; we never edit numbering.xml. Provide minimal mapping to LLM.
    if not numbering_xml_path.exists():
        return {"nums": {}, "abstracts": {}}

    tree = ET.parse(numbering_xml_path)
    root = tree.getroot()

    num_map: Dict[str, str] = {}
    num_overrides: Dict[str, List[Dict[str, Any]]] = {}
    for num in root.findall(f".//{_q('num')}"):
        numId = _get_attr(num, "numId")
        abs_el = num.find(_q("abstractNumId"))
        if numId and abs_el is not None:
            absId = _get_attr(abs_el, "val")
            if absId:
                num_map[numId] = absId
        if numId:
            overrides: List[Dict[str, Any]] = []
            for override in num.findall(_q("lvlOverride")):
                item: Dict[str, Any] = {"ilvl": _get_attr(override, "ilvl")}
                start = override.find(_q("startOverride"))
                if start is not None and _get_attr(start, "val") is not None:
                    item["startOverride"] = _get_attr(start, "val")
                override_lvl = override.find(_q("lvl"))
                if override_lvl is not None:
                    num_fmt = override_lvl.find(_q("numFmt"))
                    lvl_text = override_lvl.find(_q("lvlText"))
                    if num_fmt is not None and _get_attr(num_fmt, "val") is not None:
                        item["numFmt"] = _get_attr(num_fmt, "val")
                    if lvl_text is not None and _get_attr(lvl_text, "val") is not None:
                        item["lvlText"] = _get_attr(lvl_text, "val")
                overrides.append(item)
            if overrides:
                num_overrides[numId] = overrides

    abs_needed = {num_map[n] for n in used_num_ids if n in num_map}

    # abstractNum patterns (light)
    abstracts: Dict[str, Any] = {}
    for absn in root.findall(f".//{_q('abstractNum')}"):
        absId = _get_attr(absn, "abstractNumId")
        if not absId or absId not in abs_needed:
            continue
        lvls = []
        for lvl in absn.findall(_q("lvl")):
            ilvl = _get_attr(lvl, "ilvl")
            numFmt = lvl.find(_q("numFmt"))
            lvlText = lvl.find(_q("lvlText"))
            lvls.append({
                "ilvl": ilvl,
                "numFmt": _get_attr(numFmt, "val") if numFmt is not None else None,
                "lvlText": _get_attr(lvlText, "val") if lvlText is not None else None,
            })
        abstracts[absId] = {"abstractNumId": absId, "levels": lvls}

    nums: Dict[str, Any] = {}
    for numId in sorted(used_num_ids):
        item: Dict[str, Any] = {"numId": numId, "abstractNumId": num_map.get(numId)}
        if numId in num_overrides:
            item["levelOverrides"] = num_overrides[numId]
        nums[numId] = item

    return {"nums": nums, "abstracts": abstracts}


def build_slim_bundle(extract_dir: Path) -> Dict[str, Any]:
    snap = snapshot_stability(extract_dir)

    doc_path = extract_dir / "word" / "document.xml"
    doc_text = read_xml_text(doc_path)

    paragraphs = []
    used_style_ids: Set[str] = set()
    used_num_ids: Set[str] = set()

    contexts = collect_paragraph_contexts(doc_path)
    blocks = list(iter_paragraph_xml_blocks(doc_text))
    if len(contexts) != len(blocks):
        raise ValueError(
            f"Paragraph extraction mismatch: xml_contexts={len(contexts)} blocks={len(blocks)}"
        )

    for idx, (_s, _e, p_xml) in enumerate(blocks):
        ctx = contexts[idx]
        raw_text = paragraph_text_from_block(p_xml)
        pStyle = paragraph_pstyle_from_block(p_xml)
        numpr = paragraph_numpr_from_block(p_xml)
        hints = paragraph_ppr_hints_from_block(p_xml)
        rpr_hints = paragraph_rpr_hints_from_block(p_xml)
        has_direct_ppr = bool(extract_paragraph_ppr_inner(p_xml))
        has_uniform_direct_rpr = bool(extract_paragraph_rpr_inner(p_xml))
        contains_sect = ctx.contains_sectPr
        in_table = ctx.in_table

        if pStyle:
            used_style_ids.add(pStyle)
        if numpr.get("numId"):
            used_num_ids.add(numpr["numId"])

        display_text = raw_text
        if len(raw_text) > 200:
            display_text = raw_text[:200] + "…"
        skip_reason = compute_skip_reason(raw_text, contains_sect, in_table)

        paragraphs.append({
            "paragraph_index": idx,
            "text": display_text,
            "pStyle": pStyle,
            "numPr": numpr if (numpr.get("numId") or numpr.get("ilvl")) else None,
            "pPr_hints": hints if hints else None,
            "rPr_hints": rpr_hints if rpr_hints else None,
            "has_direct_pPr": has_direct_ppr,
            "has_uniform_direct_rPr": has_uniform_direct_rpr,
            "contains_sectPr": contains_sect,
            "in_table": in_table,
            "skip_reason": skip_reason,
            "text_was_truncated": len(raw_text) > 200,
        })

    style_catalog: Dict[str, Any] = {}
    styles_path = extract_dir / "word" / "styles.xml"
    if styles_path.exists():
        style_catalog = build_style_catalog(styles_path, used_style_ids)

    # Numbering is frequently inherited solely through a paragraph style.  It
    # still has to be included in the numbering catalog and exposed per
    # paragraph; otherwise a fully auto-numbered spec appears unnumbered.
    for paragraph in paragraphs:
        pstyle = paragraph.get("pStyle")
        inherited = style_catalog.get(pstyle, {}).get("resolved_numPr") if pstyle else None
        direct = paragraph.get("numPr")
        effective: Optional[Dict[str, Any]] = dict(inherited) if isinstance(inherited, dict) else None
        if isinstance(direct, dict) and (direct.get("numId") or direct.get("ilvl")):
            effective = effective or {}
            effective.update({key: value for key, value in direct.items() if value is not None})
        paragraph["effective_numPr"] = effective if effective else None
        if isinstance(effective, dict) and effective.get("numId"):
            used_num_ids.add(effective["numId"])

    numbering_catalog = build_numbering_catalog(extract_dir / "word" / "numbering.xml", used_num_ids)
    for paragraph in paragraphs:
        paragraph["numbering_role"] = detect_numbering_role(paragraph, numbering_catalog)

    return {
        "stability": {
            "header_footer_hashes": snap.header_footer_hashes,
            "sectPr_hash": snap.sectpr_hash,
        },
        "paragraphs": paragraphs,
        "style_catalog": style_catalog,
        "numbering_catalog": numbering_catalog,
    }


# -----------------------------
# Applying LLM instructions (pStyle only)
# -----------------------------

def strip_pstyle_from_paragraph(p_xml: str) -> str:
    result = re.sub(r"<w:pStyle\b[^>]*/>", "", p_xml)
    # Also remove empty pPr that might result
    result = re.sub(r"<w:pPr>\s*</w:pPr>", "", result)
    result = re.sub(r"<w:pPr\s*/>", "", result)
    return result

def ppr_without_pstyle(p_xml: str) -> str:
    m = re.search(r"<w:pPr\b[\s\S]*?</w:pPr>", p_xml)
    if not m:
        return ""
    ppr = m.group(0)
    # Remove pStyle
    ppr = re.sub(r"<w:pStyle\b[^>]*/>", "", ppr)
    # If pPr is now empty, return empty string
    inner = re.sub(r"<w:pPr\b[^>]*>([\s\S]*)</w:pPr>", r"\1", ppr, flags=re.S)
    if not inner.strip():
        return ""
    return ppr


def xml_escape(s: str) -> str:
    return (s.replace("&", "&amp;")
             .replace("<", "&lt;")
             .replace(">", "&gt;")
             .replace('"', "&quot;")
             .replace("'", "&apos;"))


def _strip_rsids_for_cmp(xml_text: str) -> str:
    """Remove rsid* attributes for comparison purposes."""
    return re.sub(r'\s+w:rsid\w*="[^"]*"', '', xml_text)


def _strip_proofing_for_cmp(xml_text: str) -> str:
    """Remove proofErr elements for comparison purposes."""
    return re.sub(r'<w:proofErr[^>]*/>', '', xml_text)


def extract_paragraph_ppr_inner(p_xml: str) -> str:
    if re.search(r"<w:pPr\b[^>]*/>", p_xml):
        return ""
    m = re.search(r"<w:pPr\b[^>]*>(.*?)</w:pPr>", p_xml, flags=re.S)
    if not m:
        return ""
    inner = m.group(1)
    inner = re.sub(r"<w:pStyle\b[^>]*/>", "", inner)
    return inner.strip()


def _xml_element_signature(element: ET.Element) -> Tuple[Any, ...]:
    """Return an attribute-order-independent signature for a small XML tree."""
    return (
        element.tag,
        tuple(sorted(element.attrib.items())),
        (element.text or "").strip(),
        tuple(_xml_element_signature(child) for child in list(element)),
    )


def _run_property_fragments(rpr_inner: str) -> List[Tuple[Tuple[Any, ...], str]]:
    """Split direct w:rPr children while preserving their original XML bytes."""
    fragments: List[Tuple[Tuple[Any, ...], str]] = []
    pattern = re.compile(r"(<w:([A-Za-z_][\w.-]*)\b[^>]*(?:/>|>[\s\S]*?</w:\2>))")
    for match in pattern.finditer(rpr_inner):
        raw = match.group(1).strip()
        try:
            wrapper = ET.fromstring(f'<root xmlns:w="{W_NS}">{raw}</root>')
            if not list(wrapper):
                continue
            signature = _xml_element_signature(list(wrapper)[0])
        except ET.ParseError:
            # Unknown extension prefixes are uncommon in run properties.  Keep
            # a conservative normalized signature instead of throwing away a
            # property that is demonstrably identical in every run.
            signature = (re.sub(r"\s+", " ", raw).strip(),)
        fragments.append((signature, raw))
    return fragments


def extract_paragraph_rpr_inner(p_xml: str) -> str:
    """Return only direct run properties common to all visible text runs.

    Promoting the most common *non-empty* run formatting to a paragraph style
    made an entire role bold when the exemplar merely began with a bold label.
    An absent rPr is therefore a real empty set, and mixed formatting yields
    only the exact properties shared by every visible run.
    """
    visible_xml = re.sub(r"<w:(?:del|moveFrom)\b[^>]*>[\s\S]*?</w:(?:del|moveFrom)>", "", p_xml)
    run_properties: List[List[Tuple[Tuple[Any, ...], str]]] = []

    for rm in re.finditer(r"<w:r\b[^>]*>(.*?)</w:r>", visible_xml, flags=re.S):
        run_inner = rm.group(1)
        text_nodes = re.findall(r"<w:t\b[^>]*>([\s\S]*?)</w:t>", run_inner)
        has_visible_control = bool(re.search(r"<w:(?:tab|br|cr)\b", run_inner))
        if not text_nodes and not has_visible_control:
            continue
        run_text = html.unescape("".join(text_nodes))
        if not run_text.strip() and not has_visible_control:
            continue
        m = re.search(r"<w:rPr\b[^>]*>(.*?)</w:rPr>", run_inner, flags=re.S)
        if not m:
            run_properties.append([])
            continue
        raw = m.group(1).strip()
        if not raw:
            run_properties.append([])
            continue
        cleaned = _strip_proofing_for_cmp(_strip_rsids_for_cmp(raw)).strip()
        run_properties.append(_run_property_fragments(cleaned))

    if not run_properties:
        return ""

    common = {signature for signature, _raw in run_properties[0]}
    for properties in run_properties[1:]:
        common &= {signature for signature, _raw in properties}
    if not common:
        return ""
    return "".join(raw for signature, raw in run_properties[0] if signature in common)


def derive_style_def_from_paragraph(styleId: str, name: str, p_xml: str, based_on: Optional[str] = None) -> Dict[str, Any]:
    # A paragraph's current style is part of its effective appearance.  New
    # portable styles must inherit from it; trusting an LLM-supplied "Normal"
    # here silently discarded the architect's inherited formatting.
    source_style = paragraph_pstyle_from_block(p_xml)
    return {
        "styleId": styleId,
        "name": name,
        "type": "paragraph",
        "basedOn": source_style or based_on,
        "pPr_inner": extract_paragraph_ppr_inner(p_xml),
        "rPr_inner": extract_paragraph_rpr_inner(p_xml),
    }


def _default_paragraph_style_id(styles_xml_text: str) -> Optional[str]:
    try:
        root = ET.fromstring(styles_xml_text)
    except ET.ParseError:
        return None
    for style in root.findall(f".//{_q('style')}"):
        if _get_attr(style, "type") != "paragraph":
            continue
        if str(_get_attr(style, "default") or "").lower() in {"1", "true", "on"}:
            return _get_attr(style, "styleId")
    return "Normal" if re.search(r'w:styleId="Normal"', styles_xml_text) else None


def build_portable_styles_xml(extract_dir: Path, instructions: Dict[str, Any]) -> str:
    """Build the portable stylesheet without mutating the extracted package."""
    slim_bundle = build_slim_bundle(extract_dir)
    validate_instructions(instructions, slim_bundle=slim_bundle)

    styles_path = extract_dir / "word" / "styles.xml"
    if not styles_path.exists():
        raise FileNotFoundError(f"DOCX is missing required part: {styles_path}")
    styles_text = read_xml_text(styles_path)
    default_style = _default_paragraph_style_id(styles_text)
    doc_text = read_xml_text(extract_dir / "word" / "document.xml")
    paragraph_blocks = [block for _start, _end, block in iter_paragraph_xml_blocks(doc_text)]

    derived_blocks: List[str] = []
    for style_spec in instructions.get("create_styles", []) or []:
        source_index = int(style_spec["derive_from_paragraph_index"])
        if source_index >= len(paragraph_blocks):
            raise ValueError(
                f"Style {style_spec['styleId']}: derive_from_paragraph_index out of range: {source_index}"
            )
        exemplar = paragraph_blocks[source_index]
        source_style = paragraph_pstyle_from_block(exemplar)
        derived = derive_style_def_from_paragraph(
            style_spec["styleId"],
            style_spec.get("name") or style_spec["styleId"],
            exemplar,
            based_on=source_style or default_style,
        )
        derived_blocks.append(build_style_xml_block(derived))

    portable = insert_styles_into_styles_xml(styles_text, derived_blocks)
    style_ids = set(re.findall(r'w:styleId="([^"]+)"', portable))
    for item in instructions.get("apply_pStyle", []) or []:
        if item["styleId"] not in style_ids:
            raise ValueError(f"apply_pStyle references unknown styleId: {item['styleId']}")
    return prepare_xml_text_for_utf8(portable)


def build_style_xml_block(style_def: Dict[str, Any]) -> str:
    sid = style_def.get("styleId")
    name = style_def.get("name") or sid
    based_on = style_def.get("basedOn")
    stype = style_def.get("type") or "paragraph"
    ppr_inner = style_def.get("pPr_inner") or ""
    rpr_inner = style_def.get("rPr_inner") or ""

    if not sid or not isinstance(sid, str):
        raise ValueError("styleId is required")
    if stype != "paragraph":
        raise ValueError("Only paragraph styles are supported")

    parts: List[str] = []
    parts.append(f'<w:style w:type="{stype}" w:styleId="{sid}">')
    parts.append(f'  <w:name w:val="{xml_escape(name)}"/>')
    if based_on:
        parts.append(f'  <w:basedOn w:val="{xml_escape(based_on)}"/>')
    parts.append('  <w:qFormat/>')

    if ppr_inner.strip():
        parts.append('  <w:pPr>')
        parts.append(ppr_inner.strip())
        parts.append('  </w:pPr>')

    if rpr_inner.strip():
        parts.append('  <w:rPr>')
        parts.append(rpr_inner.strip())
        parts.append('  </w:rPr>')

    parts.append('</w:style>')
    return "\n".join(parts) + "\n"


def insert_styles_into_styles_xml(styles_xml_text: str, style_blocks: List[str]) -> str:
    if not style_blocks:
        return styles_xml_text

    existing = set(re.findall(r'w:styleId="([^"]+)"', styles_xml_text))
    filtered: List[str] = []
    for sb in style_blocks:
        m = re.search(r'w:styleId="([^"]+)"', sb)
        if not m:
            raise ValueError("Style block missing w:styleId")
        sid = m.group(1)
        if sid in existing:
            continue
        filtered.append(sb)

    if not filtered:
        return styles_xml_text

    insert_point = styles_xml_text.rfind("</w:styles>")
    if insert_point == -1:
        raise ValueError("styles.xml does not contain </w:styles>")
    insertion = "\n" + "\n".join(filtered) + "\n"
    return styles_xml_text[:insert_point] + insertion + styles_xml_text[insert_point:]


def apply_pstyle_to_paragraph_block(p_xml: str, styleId: str) -> str:
    if "<w:sectPr" in p_xml:
        return p_xml

    if re.search(r"<w:pStyle\b", p_xml):
        return re.sub(
            r'(<w:pStyle\b[^>]*w:val=")([^"]+)(")',
            rf'\g<1>{styleId}\g<3>',
            p_xml,
            count=1
        )

    if re.search(r"<w:pPr\b[^>]*/>", p_xml):
        return re.sub(
            r"<w:pPr\b[^>]*/>",
            rf'<w:pPr><w:pStyle w:val="{styleId}"/></w:pPr>',
            p_xml,
            count=1
        )

    if "<w:pPr" in p_xml:
        return re.sub(
            r'(<w:pPr\b[^>]*>)',
            rf'\1<w:pStyle w:val="{styleId}"/>',
            p_xml,
            count=1
        )

    return re.sub(
        r'(<w:p\b[^>]*>)',
        rf'\1<w:pPr><w:pStyle w:val="{styleId}"/></w:pPr>',
        p_xml,
        count=1
    )


def validate_instructions(instructions: Dict[str, Any], slim_bundle: Optional[Dict[str, Any]] = None) -> None:
    if not isinstance(instructions, dict):
        raise ValueError("Instructions must be a JSON object")
    # Keep the runtime and the published JSON schema on one shared structural
    # contract (including rejecting bool-as-int indices and null arrays).
    from phase1_validator import validate_instruction_contract

    validate_instruction_contract(instructions)
    allowed_keys = {"create_styles", "apply_pStyle", "ignored_paragraphs", "roles", "notes"}
    extra = set(instructions.keys()) - allowed_keys
    if extra:
        raise ValueError(f"Invalid instruction keys: {extra}")

    for list_key in ("create_styles", "apply_pStyle", "ignored_paragraphs", "notes"):
        value = instructions.get(list_key, [])
        if value is not None and not isinstance(value, list):
            raise ValueError(f"{list_key} must be an array")

    allowed_new_style_ids: Set[str] = set(ROLE_TO_ARCH_STYLE.values())

    created_style_src_idx: Dict[str, int] = {}
    created_style_ids_seen: Set[str] = set()
    for sd in instructions.get("create_styles", []) or []:
        if not isinstance(sd, dict):
            raise ValueError("create_styles entries must be objects")
        sid = sd.get("styleId")
        if not sid or not isinstance(sid, str):
            raise ValueError("create_styles entries must have styleId (string)")
        if not (sid.startswith("CSI_") and sid.endswith("__ARCH")):
            raise ValueError(f"Style {sid}: create_styles styleId must be namespaced CSI_*__ARCH")
        if sid not in allowed_new_style_ids:
            raise ValueError(f"Style {sid}: styleId is not allowed")
        if sid in created_style_ids_seen:
            raise ValueError(f"Duplicate create_styles styleId: {sid}")
        created_style_ids_seen.add(sid)
        if any(k in sd for k in ("pPr", "rPr", "pPr_inner", "rPr_inner")):
            raise ValueError(f"Style {sid}: LLM formatting fields are forbidden. Use derive_from_paragraph_index only.")
        allowed_style_fields = {"styleId", "name", "type", "derive_from_paragraph_index", "basedOn", "role"}
        extra_fields = set(sd.keys()) - allowed_style_fields
        if extra_fields:
            raise ValueError(f"Style {sid}: invalid fields: {extra_fields}")
        if sd.get("type", "paragraph") != "paragraph":
            raise ValueError(f"Style {sid}: only paragraph styles are supported")
        src = sd.get("derive_from_paragraph_index")
        if src is None or not isinstance(src, int) or src < 0:
            raise ValueError(f"Style {sid}: derive_from_paragraph_index must be non-negative int")
        created_style_src_idx[sid] = src

    seen_para: Set[int] = set()
    for ap in instructions.get("apply_pStyle", []) or []:
        if not isinstance(ap, dict):
            raise ValueError("apply_pStyle entries must be objects")
        idx = ap.get("paragraph_index")
        sid = ap.get("styleId")
        if not isinstance(idx, int) or idx < 0:
            raise ValueError(f"Invalid paragraph_index: {idx}")
        if not isinstance(sid, str) or not sid:
            raise ValueError(f"Invalid styleId for paragraph {idx}: {sid}")
        extra_fields = set(ap.keys()) - {"paragraph_index", "styleId"}
        if extra_fields:
            raise ValueError(f"apply_pStyle[{idx}] has invalid fields: {extra_fields}")
        if idx in seen_para:
            raise ValueError(f"Duplicate paragraph_index in apply_pStyle: {idx}")
        seen_para.add(idx)

    ignored_indices: Set[int] = set()
    for ignored in instructions.get("ignored_paragraphs", []) or []:
        if not isinstance(ignored, dict):
            raise ValueError("ignored_paragraphs entries must be objects")
        extra_fields = set(ignored) - {"paragraph_index", "reason"}
        if extra_fields:
            raise ValueError(f"ignored_paragraphs entry has invalid fields: {extra_fields}")
        idx = ignored.get("paragraph_index")
        reason = ignored.get("reason")
        if not isinstance(idx, int) or idx < 0:
            raise ValueError(f"Invalid ignored paragraph_index: {idx}")
        if not isinstance(reason, str) or not reason.strip():
            raise ValueError(f"ignored_paragraphs[{idx}].reason must be a non-empty string")
        if idx in ignored_indices:
            raise ValueError(f"Duplicate paragraph_index in ignored_paragraphs: {idx}")
        if idx in seen_para:
            raise ValueError(f"Paragraph {idx} cannot be both styled and ignored")
        ignored_indices.add(idx)

    roles = instructions.get("roles")
    if roles is None or not isinstance(roles, dict):
        raise ValueError("Missing/invalid required key: roles")

    for role, spec in roles.items():
        if role not in ALLOWED_ROLES:
            raise ValueError(f"Unknown role '{role}'")
        if not isinstance(spec, dict):
            raise ValueError(f"roles['{role}'] must be an object")
        extra_fields = set(spec.keys()) - {"styleId", "exemplar_paragraph_index", "style_name"}
        if extra_fields:
            raise ValueError(f"roles['{role}'] has invalid fields: {extra_fields}")
        sid = spec.get("styleId")
        ex = spec.get("exemplar_paragraph_index")
        if not isinstance(sid, str) or not sid:
            raise ValueError(f"roles['{role}'].styleId must be a non-empty string")
        if not isinstance(ex, int) or ex < 0:
            raise ValueError(f"roles['{role}'].exemplar_paragraph_index must be non-negative int")
        if sid in created_style_src_idx and ex != created_style_src_idx[sid]:
            raise ValueError(
                f"roles['{role}'] exemplar_paragraph_index ({ex}) must equal derive_from_paragraph_index "
                f"({created_style_src_idx[sid]}) for created styleId '{sid}'"
            )

    if slim_bundle is None:
        return

    paragraphs = slim_bundle.get("paragraphs", [])
    paragraph_count = len(paragraphs)
    allowed_existing_style_ids = set(slim_bundle.get("style_catalog", {}).keys())
    default_style_ids = {
        style_id
        for style_id, info in slim_bundle.get("style_catalog", {}).items()
        if isinstance(info, dict) and info.get("default") is True
    }
    if not default_style_ids and "Normal" in allowed_existing_style_ids:
        default_style_ids.add("Normal")
    # A reserved name is not a style definition.  It may only be referenced if
    # it already exists in the source or is actually declared in create_styles.
    all_allowed_style_ids = allowed_existing_style_ids | created_style_ids_seen

    for sd in instructions.get("create_styles", []) or []:
        sid = sd["styleId"]
        if sid in allowed_new_style_ids and sid in allowed_existing_style_ids:
            raise ValueError(
                f"Reserved ARCH style collision: '{sid}' already exists in styles.xml; refuse silent reuse"
            )
        source_idx = int(sd["derive_from_paragraph_index"])
        if source_idx >= paragraph_count:
            raise ValueError(
                f"Style {sid}: derive_from_paragraph_index out of range: {source_idx} >= {paragraph_count}"
            )
        source_style = paragraphs[source_idx].get("pStyle")
        requested_base = sd.get("basedOn")
        if "basedOn" in sd and requested_base != source_style:
            if source_style:
                raise ValueError(
                    f"Style {sid}: basedOn must preserve exemplar source pStyle '{source_style}', got '{requested_base}'"
                )
            defaults = [
                style_id
                for style_id, info in slim_bundle.get("style_catalog", {}).items()
                if isinstance(info, dict) and info.get("default") is True
            ]
            allowed_fallbacks = set(defaults) | ({"Normal"} if "Normal" in allowed_existing_style_ids else set())
            if requested_base not in allowed_fallbacks:
                raise ValueError(
                    f"Style {sid}: basedOn '{requested_base}' is not the source default paragraph style"
                )
        if requested_base and requested_base not in allowed_existing_style_ids:
            raise ValueError(f"Style {sid}: basedOn references unknown source style '{requested_base}'")

    for role, spec in roles.items():
        sid = spec["styleId"]
        ex = int(spec["exemplar_paragraph_index"])
        if ex >= paragraph_count:
            raise ValueError(f"roles['{role}'].exemplar_paragraph_index out of range: {ex} >= {paragraph_count}")
        para = paragraphs[ex]
        text = (para.get("text") or "").strip()
        if not text:
            raise ValueError(f"roles['{role}'] exemplar paragraph {ex} is blank")
        skip_reason = para.get("skip_reason")
        if skip_reason is None:
            skip_reason = compute_skip_reason(text, bool(para.get("contains_sectPr")), bool(para.get("in_table")))
        if skip_reason == "sectPr":
            raise ValueError(f"roles['{role}'] exemplar paragraph {ex} contains sectPr")
        if skip_reason == "in_table":
            raise ValueError(f"roles['{role}'] exemplar paragraph {ex} is inside a table")
        if skip_reason == "editor_note":
            raise ValueError(f"roles['{role}'] exemplar paragraph {ex} cannot be an editor note")
        if skip_reason:
            raise ValueError(f"roles['{role}'] exemplar paragraph {ex} is non-classifiable (skip_reason={skip_reason})")
        if sid in allowed_existing_style_ids and (
            para.get("has_direct_pPr")
            or para.get("has_uniform_direct_rPr")
            or para.get("pPr_hints")
            or para.get("rPr_hints")
            or para.get("numPr")
        ):
            raise ValueError(
                f"roles['{role}'] exemplar paragraph {ex} contains direct formatting; "
                "derive the canonical CSI_*__ARCH style from the exemplar instead of reusing its pStyle"
            )

    for ignored_index in ignored_indices:
        if ignored_index >= paragraph_count:
            continue
        paragraph = paragraphs[ignored_index]
        semantic_exclusion = paragraph.get("skip_reason") in {
            "editor_note",
            "specifier_note",
            "copyright_notice",
        }
        if not semantic_exclusion and (
            paragraph.get("effective_numPr") or paragraph.get("numPr") or paragraph.get("numbering_role")
        ):
            raise ValueError(
                f"ignored_paragraphs[{ignored_index}] is numbered structural content and cannot be ignored"
            )

    for paragraph in paragraphs:
        if paragraph.get("skip_reason") not in {
            "editor_note",
            "specifier_note",
            "copyright_notice",
        }:
            continue
        index = int(paragraph["paragraph_index"])
        if index in seen_para:
            raise ValueError(
                f"apply_pStyle[{index}] targets detected non-CSI content ({paragraph['skip_reason']}); "
                "use ignored_paragraphs"
            )

    required_apply_indices = [int(p["paragraph_index"]) for p in paragraphs if is_classifiable_paragraph(p)]

    apply_indices = [int(ap["paragraph_index"]) for ap in instructions.get("apply_pStyle", []) or []]
    apply_set = set(apply_indices)
    required_set = set(required_apply_indices)
    covered_set = apply_set | ignored_indices
    missing = sorted(required_set - covered_set)
    unexpected = sorted(covered_set - required_set)
    if missing or unexpected:
        raise ValueError(
            "classification coverage mismatch; "
            f"missing={missing}, unexpected={unexpected}"
        )

    validate_semantic_structure(instructions, slim_bundle)

    declared_role_style_ids = {
        spec["styleId"] for spec in roles.values() if isinstance(spec, dict)
    }
    undeclared_apply_styles = sorted(
        {
            item["styleId"]
            for item in instructions.get("apply_pStyle", []) or []
            if item["styleId"] not in declared_role_style_ids
        }
    )
    if undeclared_apply_styles:
        raise ValueError(
            "apply_pStyle may use only styleIds declared in roles; "
            f"undeclared={undeclared_apply_styles}"
        )
    unused_created_styles = sorted(created_style_ids_seen - declared_role_style_ids)
    if unused_created_styles:
        raise ValueError(f"create_styles contains styles unused by roles: {unused_created_styles}")

    # Validate style identity only after structural and coverage diagnostics so
    # callers receive the most actionable document error first.
    for ap in instructions.get("apply_pStyle", []) or []:
        idx = ap["paragraph_index"]
        sid = ap["styleId"]
        if idx >= paragraph_count:
            raise ValueError(f"apply_pStyle paragraph_index out of range: {idx} >= {paragraph_count}")
        if sid not in all_allowed_style_ids:
            raise ValueError(f"apply_pStyle[{idx}] uses disallowed styleId '{sid}'")
        source_style = paragraphs[idx].get("pStyle")
        if sid in allowed_existing_style_ids:
            expected_source_styles = {source_style} if source_style else default_style_ids
            if sid not in expected_source_styles:
                source_description = source_style or f"default style ({sorted(default_style_ids)})"
                raise ValueError(
                    f"apply_pStyle[{idx}] reuses styleId '{sid}' but source paragraph uses {source_description}"
                )

    for role, spec in roles.items():
        sid = spec["styleId"]
        ex = int(spec["exemplar_paragraph_index"])
        if sid not in all_allowed_style_ids:
            raise ValueError(f"roles['{role}'] uses disallowed styleId '{sid}'")
        para = paragraphs[ex]
        text = (para.get("text") or "").strip()
        source_style = para.get("pStyle")
        if sid in allowed_existing_style_ids:
            expected_source_styles = {source_style} if source_style else default_style_ids
            if sid not in expected_source_styles:
                source_description = source_style or f"default style ({sorted(default_style_ids)})"
                raise ValueError(
                    f"roles['{role}'] styleId '{sid}' does not match exemplar paragraph {ex} source {source_description}"
                )
        combined_section_style = (
            role == "SectionTitle"
            and sid == ROLE_TO_ARCH_STYLE["SectionID"]
            and bool(RE_SECTION_WITH_TITLE.match(text))
            and isinstance(roles.get("SectionID"), dict)
            and roles["SectionID"].get("styleId") == sid
            and int(roles["SectionID"].get("exemplar_paragraph_index", -1)) == ex
        )
        if sid in created_style_ids_seen and sid != ROLE_TO_ARCH_STYLE[role] and not combined_section_style:
            raise ValueError(
                f"roles['{role}'] must use its canonical generated styleId '{ROLE_TO_ARCH_STYLE[role]}', got '{sid}'"
            )

def validate_semantic_structure(instructions: Dict[str, Any], slim_bundle: Dict[str, Any]) -> None:
    paragraphs = slim_bundle.get("paragraphs", [])
    roles = instructions.get("roles", {}) or {}
    apply_map = {item["paragraph_index"]: item["styleId"] for item in instructions.get("apply_pStyle", []) or []}
    role_style = {role: spec.get("styleId") for role, spec in roles.items() if isinstance(spec, dict)}
    expected_roles, strong_hits = infer_expected_roles(
        paragraphs,
        numbering_catalog=slim_bundle.get("numbering_catalog"),
    )
    strong_roles_by_index: Dict[int, Set[str]] = {}
    for strong_role, indices in strong_hits.items():
        for index in indices:
            strong_roles_by_index.setdefault(int(index), set()).add(strong_role)

    nums = slim_bundle.get("numbering_catalog", {}).get("nums", {})
    abstracts = slim_bundle.get("numbering_catalog", {}).get("abstracts", {})
    style_catalog = slim_bundle.get("style_catalog", {})

    classifiable = [p for p in paragraphs if is_role_candidate_paragraph(p)]

    signals: Dict[int, StructureSignal] = {}
    numbered_families: Set[Tuple[Optional[str], Optional[str]]] = set()
    for p in classifiable:
        idx = int(p["paragraph_index"])
        text_roles = strong_roles_by_index.get(idx, set())
        text_role = next(iter(sorted(text_roles)), None)

        num_id: Optional[str] = None
        ilvl: Optional[str] = None
        source = "none"
        direct = p.get("numPr") if isinstance(p.get("numPr"), dict) else None
        effective = p.get("effective_numPr") if isinstance(p.get("effective_numPr"), dict) else None
        style_numpr = style_catalog.get(p.get("pStyle"), {}).get("resolved_numPr") if p.get("pStyle") else None

        if isinstance(direct, dict) and (direct.get("numId") or direct.get("ilvl")):
            source = "direct_numpr"
            merged = effective or direct
            num_id, ilvl = merged.get("numId"), merged.get("ilvl")
        elif isinstance(style_numpr, dict) and (style_numpr.get("numId") or style_numpr.get("ilvl")):
            source = "style_numpr"
            num_id, ilvl = style_numpr.get("numId"), style_numpr.get("ilvl")
        elif text_role in {"ARTICLE", "PARAGRAPH", "SUBPARAGRAPH", "SUBSUBPARAGRAPH"}:
            source = "text_literal"

        abstract_num_id = nums.get(num_id, {}).get("abstractNumId") if num_id else None
        num_fmt = None
        lvl_text = None
        pattern = _extract_numbering_pattern_from_numpr(
            {"numId": num_id, "ilvl": ilvl} if num_id or ilvl else None,
            slim_bundle.get("numbering_catalog", {}),
        )
        if pattern:
            abstract_num_id = pattern.get("abstractNumId", abstract_num_id)
            num_fmt = pattern.get("numFmt")
            lvl_text = pattern.get("lvlText")

        family_key = None
        if source in {"style_numpr", "direct_numpr"} and (num_id or ilvl):
            family_key = (abstract_num_id if abstract_num_id else num_id, ilvl)
            numbered_families.add(family_key)

        signals[idx] = StructureSignal(
            paragraph_index=idx,
            text_role=text_role,
            numbering_source=source,
            num_id=num_id,
            abstract_num_id=abstract_num_id,
            ilvl=ilvl,
            num_fmt=num_fmt,
            lvl_text=lvl_text,
            family_key=family_key,
        )

    missing_roles = sorted(role for role in expected_roles if role not in roles)
    if missing_roles:
        raise ValueError(f"Semantic validation failed: missing expected roles {missing_roles}")

    if numbered_families and not roles:
        raise ValueError("Semantic validation failed: numbered structure detected but roles is empty")

    # Numbering definitions are often duplicated to restart a sequence. Compare
    # the rendered numbering signature rather than numId, but require every
    # distinct level/pattern to have an explicit role exemplar. The former
    # "deepest role absorbs everything below it" rule hid lost levels.
    required_numbering_signatures: Set[Tuple[Optional[str], Optional[str], Optional[str]]] = {
        (signal.ilvl, signal.num_fmt, signal.lvl_text)
        for signal in signals.values()
        if signal.family_key is not None
    }
    covered_numbering_signatures: Set[Tuple[Optional[str], Optional[str], Optional[str]]] = set()
    for spec in roles.values():
        if not isinstance(spec, dict):
            continue
        s = signals.get(int(spec["exemplar_paragraph_index"]))
        if s and s.family_key is not None:
            covered_numbering_signatures.add((s.ilvl, s.num_fmt, s.lvl_text))

    missing_numbering = required_numbering_signatures - covered_numbering_signatures
    if missing_numbering:
        rendered = sorted(
            [f"ilvl={ilvl!r}, numFmt={num_fmt!r}, lvlText={lvl_text!r}" for ilvl, num_fmt, lvl_text in missing_numbering]
        )
        raise ValueError(
            "Semantic validation failed: missing numbered hierarchy coverage: " + "; ".join(rendered)
        )

    for role, spec in roles.items():
        exemplar_idx = int(spec["exemplar_paragraph_index"])
        candidates = strong_roles_by_index.get(exemplar_idx, set())
        if candidates and role not in candidates:
            raise ValueError(
                f"Semantic validation failed: roles['{role}'] exemplar paragraph {exemplar_idx} looks like {sorted(candidates)}"
            )

    # A shared style is safe only when its source exemplars have the same
    # observable source profile. This permits genuinely identical PART/ARTICLE
    # styles but rejects role collapse onto an unrelated comment/default style.
    roles_by_style: Dict[str, List[Tuple[str, int]]] = {}
    for role, spec in roles.items():
        roles_by_style.setdefault(spec["styleId"], []).append(
            (role, int(spec["exemplar_paragraph_index"]))
        )
    for style_id, role_exemplars in roles_by_style.items():
        if len(role_exemplars) < 2:
            continue
        profiles: Set[str] = set()
        for _role, exemplar_idx in role_exemplars:
            paragraph = paragraphs[exemplar_idx]
            profile = {
                "pStyle": paragraph.get("pStyle"),
                "numPr": paragraph.get("effective_numPr") or paragraph.get("numPr"),
                "pPr_hints": paragraph.get("pPr_hints"),
                "rPr_hints": paragraph.get("rPr_hints"),
            }
            profiles.add(json.dumps(profile, sort_keys=True, separators=(",", ":")))
        if len(profiles) > 1:
            names = [role for role, _idx in role_exemplars]
            raise ValueError(
                f"Semantic validation failed: roles {names} collapse onto styleId '{style_id}' "
                "despite different source formatting/numbering profiles"
            )

    for role, indices in strong_hits.items():
        expected_style = role_style.get(role)
        if not expected_style:
            continue
        mismatches = [idx for idx in indices if apply_map.get(idx) != expected_style]
        if mismatches:
            samples = ", ".join(str(i) for i in mismatches[:5])
            raise ValueError(
                f"Semantic validation failed: role {role} expects styleId {expected_style}; mismatched paragraph indices [{samples}]"
            )




def apply_instructions(extract_dir: Path, instructions: Dict[str, Any]) -> None:
    slim_bundle = build_slim_bundle(extract_dir)
    validate_instructions(instructions, slim_bundle=slim_bundle)

    snap = snapshot_stability(extract_dir)

    styles_path = extract_dir / "word" / "styles.xml"
    styles_text = read_xml_text(styles_path)

    doc_path = extract_dir / "word" / "document.xml"
    doc_text = read_xml_text(doc_path)

    blocks = list(iter_paragraph_xml_blocks(doc_text))
    para_blocks = [b[2] for b in blocks]
    original_para_blocks = list(para_blocks)

    # 1) Build the portable stylesheet entirely in memory.  This preserves
    # source-style inheritance and avoids a partial write if later validation
    # fails.
    style_defs = instructions.get("create_styles") or []
    styles_new = build_portable_styles_xml(extract_dir, instructions)

    styles_text_final = styles_new
    style_ids_in_styles = set(re.findall(r'w:styleId="([^"]+)"', styles_text_final))

    # ensure referenced styles exist
    for sd in style_defs:
        if sd["styleId"] not in style_ids_in_styles:
            raise ValueError(f"create_styles styleId not found in styles.xml after insertion: {sd['styleId']}")
    for item in (instructions.get("apply_pStyle") or []):
        if item["styleId"] not in style_ids_in_styles:
            raise ValueError(f"apply_pStyle references unknown styleId: {item['styleId']}")
    for role, spec in (instructions.get("roles") or {}).items():
        sid = spec["styleId"]
        ex = int(spec["exemplar_paragraph_index"])
        if sid not in style_ids_in_styles:
            raise ValueError(f"roles['{role}'] references unknown styleId: {sid}")
        if ex < 0 or ex >= len(para_blocks):
            raise ValueError(f"roles['{role}'] exemplar_paragraph_index out of range: {ex}")
        if paragraph_contains_sectpr(para_blocks[ex]):
            raise ValueError(f"roles['{role}'] exemplar paragraph {ex} contains sectPr; refuse.")

    # 2) Apply paragraph styles by index (pStyle insertion ONLY)
    idx_map: Dict[int, str] = {}
    for item in (instructions.get("apply_pStyle") or []):
        idx_map[int(item["paragraph_index"])] = item["styleId"]

    original_ppr = {i: ppr_without_pstyle(pb) for i, pb in enumerate(para_blocks)}

    for idx, sid in idx_map.items():
        if idx < 0 or idx >= len(para_blocks):
            raise ValueError(f"paragraph_index out of range: {idx}")
        if paragraph_contains_sectpr(para_blocks[idx]):
            raise ValueError(f"Refusing to apply style to paragraph {idx} because it contains sectPr.")
        para_blocks[idx] = apply_pstyle_to_paragraph_block(para_blocks[idx], sid)

    # drift checks: only pStyle may differ
    for idx in idx_map.keys():
        before = strip_pstyle_from_paragraph(original_para_blocks[idx])
        after = strip_pstyle_from_paragraph(para_blocks[idx])
        if before != after:
            print(f"=== BEFORE (paragraph {idx}) ===")
            print(before[:2000])
            print(f"=== AFTER (paragraph {idx}) ===")
            print(after[:2000])
            raise ValueError(f"Paragraph drift detected at index {idx}: changes beyond <w:pStyle>.")
    for i, pb in enumerate(para_blocks):
        if original_ppr[i] != ppr_without_pstyle(pb):
            raise ValueError(f"Paragraph properties drift detected at index {i} (beyond w:pStyle).")

    # reassemble document.xml
    out_parts: List[str] = []
    last_end = 0
    for i, (s, e, _p) in enumerate(blocks):
        out_parts.append(doc_text[last_end:s])
        out_parts.append(para_blocks[i])
        last_end = e
    out_parts.append(doc_text[last_end:])
    doc_new = "".join(out_parts)

    styles_original_bytes = styles_path.read_bytes()
    doc_original_bytes = doc_path.read_bytes()
    styles_tmp = styles_path.with_name(styles_path.name + ".phase1.tmp")
    doc_tmp = doc_path.with_name(doc_path.name + ".phase1.tmp")
    try:
        styles_tmp.write_bytes(styles_new.encode("utf-8"))
        doc_tmp.write_bytes(prepare_xml_text_for_utf8(doc_new).encode("utf-8"))
        styles_tmp.replace(styles_path)
        doc_tmp.replace(doc_path)
        verify_stability(extract_dir, snap)
    except Exception:
        styles_path.write_bytes(styles_original_bytes)
        doc_path.write_bytes(doc_original_bytes)
        raise
    finally:
        for temporary in (styles_tmp, doc_tmp):
            if temporary.exists():
                temporary.unlink()


# -----------------------------
# Phase 1 contract output
# -----------------------------

def _build_style_name_map(styles_xml_path: Path) -> Dict[str, str]:
    if not styles_xml_path.exists():
        return {}
    out: Dict[str, str] = {}
    tree = ET.parse(styles_xml_path)
    root = tree.getroot()
    for st in root.findall(f".//{_q('style')}"):
        sid = _get_attr(st, "styleId")
        if not sid:
            continue
        name_el = st.find(_q("name"))
        if name_el is not None:
            nm = _get_attr(name_el, "val")
            if nm:
                out[sid] = nm
    return out


def _build_styles_by_id(styles_xml_path: Path) -> Dict[str, ET.Element]:
    if not styles_xml_path.exists():
        return {}
    tree = ET.parse(styles_xml_path)
    root = tree.getroot()
    styles_by_id: Dict[str, ET.Element] = {}
    for st in root.findall(f".//{_q('style')}"):
        sid = _get_attr(st, "styleId")
        if sid:
            styles_by_id[sid] = st
    return styles_by_id


def _find_style_chain_numpr(style_id: str, styles_by_id: Dict[str, ET.Element]) -> Optional[Dict[str, str]]:
    """Resolve style + basedOn chain looking for w:numPr and return numId/ilvl when found."""
    visited: Set[str] = set()
    current = style_id
    while current and current not in visited:
        visited.add(current)
        st = styles_by_id.get(current)
        if st is None:
            break

        ppr = st.find(_q("pPr"))
        if ppr is not None:
            numpr = ppr.find(_q("numPr"))
            if numpr is not None:
                num_id_el = numpr.find(_q("numId"))
                ilvl_el = numpr.find(_q("ilvl"))
                out: Dict[str, str] = {}
                num_id = _get_attr(num_id_el, "val") if num_id_el is not None else None
                ilvl = _get_attr(ilvl_el, "val") if ilvl_el is not None else None
                if num_id:
                    out["numId"] = num_id
                if ilvl:
                    out["ilvl"] = ilvl
                return out

        based = st.find(_q("basedOn"))
        current = _get_attr(based, "val") if based is not None else None

    return None


def _determine_numbering_provenance(
    style_id: str,
    exemplar_idx: int,
    paragraphs: List[Dict],
    styles_xml_path: Path,
) -> str:
    """Determine whether numbering is from style, direct paragraph, literal text, or absent."""
    exemplar = paragraphs[exemplar_idx] if 0 <= exemplar_idx < len(paragraphs) else {}
    numpr = exemplar.get("numPr") if isinstance(exemplar, dict) else None
    if isinstance(numpr, dict) and (numpr.get("numId") or numpr.get("ilvl")):
        return "direct_numpr"

    styles_by_id = _build_styles_by_id(styles_xml_path)
    if _find_style_chain_numpr(style_id, styles_by_id) is not None:
        return "style_numpr"

    txt = (exemplar.get("text") or "") if isinstance(exemplar, dict) else ""
    marker_patterns = [
        r"^\s*\d+\.\d{2,}\s+",
        r"^\s*[A-Z]\.\s+",
        r"^\s*\d+\.\s+",
        r"^\s*[a-z]\.\s+",
    ]
    if any(re.match(pat, txt) for pat in marker_patterns):
        return "text_literal"
    return "none"


def _extract_numbering_pattern_from_numpr(
    numpr: Optional[Dict[str, str]],
    numbering_catalog: Dict[str, Any],
) -> Optional[Dict[str, str]]:
    if numpr is None:
        return None

    pattern: Dict[str, str] = {}
    num_id = numpr.get("numId")
    ilvl = numpr.get("ilvl")
    if num_id:
        pattern["numId"] = num_id
    if ilvl:
        pattern["ilvl"] = ilvl

    nums = numbering_catalog.get("nums", {}) if isinstance(numbering_catalog, dict) else {}
    abstracts = numbering_catalog.get("abstracts", {}) if isinstance(numbering_catalog, dict) else {}

    abstract_num_id = nums.get(num_id, {}).get("abstractNumId") if num_id else None
    if abstract_num_id:
        pattern["abstractNumId"] = abstract_num_id
        levels = abstracts.get(abstract_num_id, {}).get("levels", [])
        level = next((lvl for lvl in levels if lvl.get("ilvl") == ilvl), None)
        if level is None and levels:
            level = levels[0]
        if isinstance(level, dict):
            if level.get("numFmt"):
                pattern["numFmt"] = level["numFmt"]
            if level.get("lvlText"):
                pattern["lvlText"] = level["lvlText"]

    if num_id:
        overrides = nums.get(num_id, {}).get("levelOverrides", [])
        override = next((item for item in overrides if item.get("ilvl") == ilvl), None)
        if isinstance(override, dict):
            for key in ("numFmt", "lvlText", "startOverride"):
                if override.get(key) is not None:
                    pattern[key] = override[key]

    return pattern if pattern else None


def _extract_numbering_pattern(
    style_id: str,
    styles_xml_path: Path,
    numbering_catalog: Dict[str, Any],
) -> Optional[Dict[str, str]]:
    styles_by_id = _build_styles_by_id(styles_xml_path)
    return _extract_numbering_pattern_from_numpr(
        _find_style_chain_numpr(style_id, styles_by_id), numbering_catalog
    )


def build_style_registry_dict(
    extract_dir: Path,
    source_docx_name: str,
    instructions: Dict[str, Any],
    pre_apply_bundle: Optional[Dict[str, Any]] = None,
    styles_xml_path: Optional[Path] = None,
    source_sha256: Optional[str] = None,
) -> Dict[str, Any]:
    """Build the arch_style_registry payload dict without writing to disk."""
    if source_sha256 is None:
        raise ValueError("source_sha256 is required for the version 2 style registry")
    if not re.fullmatch(r"[0-9a-f]{64}", source_sha256):
        raise ValueError("source_sha256 must be a lowercase 64-character SHA-256 hex digest")
    roles = instructions.get("roles") or {}
    styles_path = styles_xml_path or (extract_dir / "word" / "styles.xml")
    name_map = _build_style_name_map(styles_path)
    bundle = pre_apply_bundle if pre_apply_bundle is not None else build_slim_bundle(extract_dir)
    paragraphs = bundle.get("paragraphs", [])
    created_style_ids = {sd.get("styleId") for sd in (instructions.get("create_styles") or []) if isinstance(sd, dict)}
    styles_by_id = _build_styles_by_id(styles_path)

    used_num_ids: Set[str] = set()
    for spec in roles.values():
        style_id = spec.get("styleId")
        if not style_id:
            continue
        numpr = _find_style_chain_numpr(style_id, styles_by_id)
        if numpr and numpr.get("numId"):
            used_num_ids.add(numpr["numId"])
        exemplar_idx = int(spec.get("exemplar_paragraph_index", -1))
        if 0 <= exemplar_idx < len(paragraphs):
            direct_numpr = paragraphs[exemplar_idx].get("effective_numPr") or paragraphs[exemplar_idx].get("numPr")
            if isinstance(direct_numpr, dict) and direct_numpr.get("numId"):
                used_num_ids.add(direct_numpr["numId"])
    numbering_catalog = build_numbering_catalog(extract_dir / "word" / "numbering.xml", used_num_ids)

    out_roles: Dict[str, Any] = {}
    for role, spec in roles.items():
        style_id = spec["styleId"]
        exemplar_idx = int(spec["exemplar_paragraph_index"])
        entry: Dict[str, Any] = {
            "style_id": style_id,
            "exemplar_paragraph_index": exemplar_idx,
        }
        style_name = name_map.get(style_id)
        if style_name:
            entry["style_name"] = style_name

        if 0 <= exemplar_idx < len(paragraphs):
            exemplar = paragraphs[exemplar_idx]
            resolved: Dict[str, Any] = {}
            if exemplar.get("pPr_hints"):
                resolved["pPr_hints"] = exemplar["pPr_hints"]
            if exemplar.get("rPr_hints"):
                resolved["rPr_hints"] = exemplar["rPr_hints"]
            if resolved:
                entry["resolved_formatting"] = resolved
        entry["numbering_provenance"] = _determine_numbering_provenance(
            style_id, exemplar_idx, paragraphs, styles_path
        )
        pattern = None
        if entry["numbering_provenance"] == "style_numpr":
            pattern = _extract_numbering_pattern(style_id, styles_path, numbering_catalog)
        elif entry["numbering_provenance"] == "direct_numpr" and 0 <= exemplar_idx < len(paragraphs):
            pattern = _extract_numbering_pattern_from_numpr(
                paragraphs[exemplar_idx].get("effective_numPr") or paragraphs[exemplar_idx].get("numPr"),
                numbering_catalog,
            )
        if pattern:
            entry["numbering_pattern"] = pattern

        out_roles[role] = entry

    source_tokens: Dict[str, str] = {}
    for role_name in ("SectionID", "SectionTitle"):
        role_spec = roles.get(role_name)
        if not isinstance(role_spec, dict):
            continue
        exemplar_idx = int(role_spec.get("exemplar_paragraph_index", -1))
        if 0 <= exemplar_idx < len(paragraphs):
            text = str(paragraphs[exemplar_idx].get("text", "")).strip()
            if text:
                source_tokens[role_name] = text

    # Combined SECTION-number/title paragraphs legitimately use one style and
    # one exemplar.  Preserve distinct tokens when possible for Phase 2's
    # header/footer substitution logic.
    if "SectionID" in source_tokens:
        original_section_token = source_tokens["SectionID"]
        combined = RE_SECTION_WITH_TITLE.match(original_section_token)
        if combined and combined.group(2).strip():
            source_tokens["SectionID"] = combined.group(1).strip()
            if source_tokens.get("SectionTitle") in {None, original_section_token}:
                source_tokens["SectionTitle"] = combined.group(2).strip(" -–—:")

    payload: Dict[str, Any] = {
        "version": 2,
        "source_docx": source_docx_name,
        "source_tokens": source_tokens,
        "roles": out_roles,
    }
    payload["source_sha256"] = source_sha256
    return payload


def emit_arch_style_registry(
    extract_dir: Path,
    source_docx_name: str,
    instructions: Dict[str, Any],
    out_path: Optional[Path] = None,
    *,
    source_sha256: Optional[str] = None,
) -> Path:
    payload = build_style_registry_dict(
        extract_dir,
        source_docx_name,
        instructions,
        source_sha256=source_sha256,
    )

    if out_path is None:
        out_path = extract_dir / "arch_style_registry.json"
    out_path.write_text(json.dumps(payload, indent=2), encoding="utf-8")
    return out_path
