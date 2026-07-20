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
import posixpath
import re
import shutil
import stat
import zipfile
from dataclasses import dataclass
from pathlib import Path
from pathlib import PurePosixPath, PureWindowsPath
from typing import Any, Dict, List, Optional, Set, Tuple
import html
from urllib.parse import unquote, urlsplit

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
PKG_REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"

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
    contains_drawing: bool = False
    contains_textbox: bool = False


@dataclass(frozen=True)
class _ScannedParagraph:
    start: int
    end: int
    xml: str
    context: ParagraphContext


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
    """Hash header/footer owners named by document relationships."""
    hashes: Dict[str, str] = {}
    rels_path = extract_dir / "word" / "_rels" / "document.xml.rels"
    if not rels_path.is_file():
        return hashes
    root = ET.fromstring(rels_path.read_bytes())
    for rel in root.findall(f"{{{PKG_REL_NS}}}Relationship"):
        rel_type = rel.attrib.get("Type", "")
        if not rel_type.endswith(("/header", "/footer")):
            continue
        if rel.attrib.get("TargetMode") == "External":
            continue
        target = rel.attrib.get("Target", "")
        decoded = unquote(target)
        parsed = urlsplit(decoded)
        windows_target = PureWindowsPath(decoded)
        if (
            not decoded
            or "%2e" in target.casefold()
            or parsed.scheme
            or parsed.netloc
            or parsed.query
            or parsed.fragment
            or "\\" in decoded
            or decoded.startswith("/")
            or windows_target.drive
            or windows_target.is_absolute()
        ):
            raise ValueError(f"Unsafe target header/footer relationship: {target!r}")
        part_name = posixpath.normpath(posixpath.join("word", decoded))
        if part_name in {"", ".", ".."} or part_name.startswith("../"):
            raise ValueError(f"Unsafe target header/footer relationship: {target!r}")
        part_path = extract_dir / Path(*PurePosixPath(part_name).parts)
        if not part_path.is_file():
            raise ValueError(f"Referenced header/footer part is missing: {part_name}")
        hashes[part_name] = sha256_bytes(part_path.read_bytes())
        owner = PurePosixPath(part_name)
        rels_name = (
            owner.parent / "_rels" / f"{owner.name}.rels"
        ).as_posix()
        owner_rels = extract_dir / Path(*PurePosixPath(rels_name).parts)
        if owner_rels.is_file():
            hashes[rels_name] = sha256_bytes(owner_rels.read_bytes())
    return hashes


def snapshot_doc_rels_hash(extract_dir: Path) -> str:
    rels_path = extract_dir / "word" / "_rels" / "document.xml.rels"
    if not rels_path.exists():
        return ""
    return sha256_bytes(rels_path.read_bytes())


def extract_document_sectpr_blocks(document_xml: str) -> List[str]:
    """Return real document section properties outside tables/drawings."""
    ranges: List[Tuple[int, int]] = []
    for qname in (
        "w:tbl",
        "w:drawing",
        "w:pict",
        "w:object",
        "w:txbxContent",
        "v:textbox",
        "wps:txbx",
    ):
        ranges.extend(_named_xml_block_ranges(document_xml, qname))
    merged: List[Tuple[int, int]] = []
    for start, end in sorted(ranges):
        if merged and start <= merged[-1][1]:
            merged[-1] = (merged[-1][0], max(merged[-1][1], end))
        else:
            merged.append((start, end))
    pieces: List[str] = []
    cursor = 0
    for start, end in merged:
        pieces.append(document_xml[cursor:start])
        cursor = end
    pieces.append(document_xml[cursor:])
    analysis_xml = "".join(pieces)
    return _extract_named_xml_blocks(analysis_xml, "w:sectPr")


def extract_sectpr_block(document_xml: str) -> str:
    # Preserve exact paired and self-closing in-scope blocks for the stability hash.
    return "\n".join(extract_document_sectpr_blocks(document_xml))


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

_OUT_OF_SCOPE_CONTAINER_NAMES = {
    "w:drawing",
    "w:pict",
    "w:object",
    "w:txbxContent",
    "v:textbox",
    "wps:txbx",
}


def _element_token_info(token: str) -> Optional[Tuple[str, bool, bool]]:
    """Return ``(qname, closing, self_closing)`` for an element token."""
    if token.startswith(("<!--", "<![CDATA[", "<?", "<!")):
        return None
    closing_match = re.match(r"</\s*([A-Za-z_][\w.-]*(?::[A-Za-z_][\w.-]*)?)\s*>", token)
    if closing_match:
        return closing_match.group(1), True, False
    opening_match = re.match(
        r"<\s*([A-Za-z_][\w.-]*(?::[A-Za-z_][\w.-]*)?)(?=\s|/?>)",
        token,
    )
    if not opening_match:
        return None
    return opening_match.group(1), False, bool(re.search(r"/\s*>$", token))


def _iter_markup_tokens(xml_text: str):
    position = 0
    while True:
        start = xml_text.find("<", position)
        if start == -1:
            return
        end = _xml_markup_end(xml_text, start)
        if end <= start or end > len(xml_text) or not xml_text[end - 1 : end] in {">"}:
            raise ValueError(f"Malformed XML markup beginning at character {start}")
        yield start, end, xml_text[start:end]
        position = end


def _named_xml_block_ranges(xml_text: str, qname: str) -> List[Tuple[int, int]]:
    """Return outermost paired/self-closing *qname* ranges."""
    ranges: List[Tuple[int, int]] = []
    depth = 0
    block_start: Optional[int] = None
    for start, end, token in _iter_markup_tokens(xml_text):
        info = _element_token_info(token)
        if info is None:
            continue
        token_qname, closing, self_closing = info
        if token_qname != qname:
            continue
        if closing:
            if depth == 0 or block_start is None:
                raise ValueError(f"Unexpected closing tag </{qname}>")
            depth -= 1
            if depth == 0:
                ranges.append((block_start, end))
                block_start = None
            continue
        if depth == 0:
            block_start = start
        if self_closing:
            if depth == 0 and block_start is not None:
                ranges.append((block_start, end))
                block_start = None
        else:
            depth += 1
    if depth or block_start is not None:
        raise ValueError(f"Unclosed <{qname}> element")
    return ranges


def _extract_named_xml_blocks(xml_text: str, qname: str) -> List[str]:
    """Extract paired and self-closing *qname* blocks with depth tracking."""
    return [xml_text[start:end] for start, end in _named_xml_block_ranges(xml_text, qname)]


def _scan_paragraph_xml_blocks(document_xml_text: str) -> List[_ScannedParagraph]:
    """Lexically scan non-overlapping document paragraphs.

    Paragraphs nested inside another paragraph occur in drawing text boxes.
    They are deliberately not assigned Phase 1 paragraph indices; their host
    paragraph is marked out of scope instead.  Table depth is tracked with a
    real element stack, so a nested table cannot make later paragraphs appear
    to be outside their containing table.
    """
    scanned: List[_ScannedParagraph] = []
    stack: List[Tuple[str, bool]] = []
    table_depth = 0
    excluded_depth = 0
    paragraph_start: Optional[int] = None
    paragraph_stack_depth: Optional[int] = None
    paragraph_in_table = False
    paragraph_contains_sectpr = False
    paragraph_contains_drawing = False
    paragraph_contains_textbox = False

    for start, end, token in _iter_markup_tokens(document_xml_text):
        info = _element_token_info(token)
        if info is None:
            continue
        qname, closing, self_closing = info

        if closing:
            if not stack or stack[-1][0] != qname:
                raise ValueError(f"Mismatched XML closing tag </{qname}>")
            _open_qname, opened_excluded = stack.pop()

            if (
                qname == "w:p"
                and paragraph_start is not None
                and paragraph_stack_depth == len(stack)
            ):
                scanned.append(
                    _ScannedParagraph(
                        start=paragraph_start,
                        end=end,
                        xml=document_xml_text[paragraph_start:end],
                        context=ParagraphContext(
                            in_table=paragraph_in_table,
                            contains_sectPr=paragraph_contains_sectpr,
                            contains_drawing=paragraph_contains_drawing,
                            contains_textbox=paragraph_contains_textbox,
                        ),
                    )
                )
                paragraph_start = None
                paragraph_stack_depth = None

            if qname == "w:tbl":
                table_depth -= 1
            if opened_excluded:
                excluded_depth -= 1
            continue

        starts_excluded = qname in _OUT_OF_SCOPE_CONTAINER_NAMES
        if paragraph_start is not None:
            # Section properties in a nested drawing/text-box story do not
            # belong to the visible host paragraph.
            if qname == "w:sectPr" and excluded_depth == 0:
                paragraph_contains_sectpr = True
            if qname in {"w:drawing", "w:pict", "w:object"}:
                paragraph_contains_drawing = True
            if qname in {"w:txbxContent", "v:textbox", "wps:txbx"}:
                paragraph_contains_textbox = True

        if qname == "w:p" and paragraph_start is None and excluded_depth == 0:
            paragraph_start = start
            paragraph_stack_depth = len(stack)
            paragraph_in_table = table_depth > 0
            paragraph_contains_sectpr = False
            paragraph_contains_drawing = False
            paragraph_contains_textbox = False
            if self_closing:
                scanned.append(
                    _ScannedParagraph(
                        start=start,
                        end=end,
                        xml=document_xml_text[start:end],
                        context=ParagraphContext(
                            in_table=paragraph_in_table,
                            contains_sectPr=False,
                        ),
                    )
                )
                paragraph_start = None
                paragraph_stack_depth = None
                continue

        if self_closing:
            continue
        stack.append((qname, starts_excluded))
        if qname == "w:tbl":
            table_depth += 1
        if starts_excluded:
            excluded_depth += 1

    if stack:
        raise ValueError(f"Unclosed XML element <{stack[-1][0]}>")
    if paragraph_start is not None:
        raise ValueError("Unclosed w:p paragraph")
    return scanned


def iter_paragraph_xml_blocks(document_xml_text: str):
    for paragraph in _scan_paragraph_xml_blocks(document_xml_text):
        yield paragraph.start, paragraph.end, paragraph.xml


def _strip_out_of_scope_subtrees(xml_text: str) -> str:
    """Remove drawing/text-box subtrees from an analysis-only XML string."""
    spans: List[Tuple[int, int]] = []
    stack: List[Tuple[str, bool, Optional[int]]] = []
    excluded_depth = 0
    for start, end, token in _iter_markup_tokens(xml_text):
        info = _element_token_info(token)
        if info is None:
            continue
        qname, closing, self_closing = info
        if closing:
            if not stack or stack[-1][0] != qname:
                raise ValueError(f"Mismatched XML closing tag </{qname}>")
            _name, opened_excluded, span_start = stack.pop()
            if opened_excluded:
                excluded_depth -= 1
                if excluded_depth == 0 and span_start is not None:
                    spans.append((span_start, end))
            continue

        starts_excluded = qname in _OUT_OF_SCOPE_CONTAINER_NAMES
        if self_closing:
            if starts_excluded and excluded_depth == 0:
                spans.append((start, end))
            continue
        span_start = start if starts_excluded and excluded_depth == 0 else None
        stack.append((qname, starts_excluded, span_start))
        if starts_excluded:
            excluded_depth += 1

    if not spans:
        return xml_text
    pieces: List[str] = []
    last_end = 0
    for start, end in spans:
        pieces.append(xml_text[last_end:start])
        last_end = end
    pieces.append(xml_text[last_end:])
    return "".join(pieces)


def _mask_text_box_subtrees(p_xml: str) -> str:
    """Hide text-box story XML while preserving every original index.

    Metadata reads and paragraph-level edits use the masked string to locate
    only the host paragraph's properties. Splices are then made against the
    original XML, so the complete ``w:txbxContent`` subtree remains byte exact.
    """
    pieces: List[str] = []
    cursor = 0
    for start, end in _named_xml_block_ranges(p_xml, "w:txbxContent"):
        pieces.append(p_xml[cursor:start])
        pieces.append(" " * (end - start))
        cursor = end
    if cursor == 0:
        return p_xml
    pieces.append(p_xml[cursor:])
    return "".join(pieces)


def paragraph_text_from_block(p_xml: str) -> str:
    # Deleted/moved-from text, field instructions, and drawing/text-box content
    # are not visible in-scope paragraph content for Phase 1 classification.
    visible = _strip_out_of_scope_subtrees(p_xml)
    visible = re.sub(r"<w:(?:del|moveFrom)\b[^>]*>[\s\S]*?</w:(?:del|moveFrom)>", "", visible)
    visible = re.sub(r"<w:instrText\b[^>]*>[\s\S]*?</w:instrText>", "", visible)
    # Tabs and explicit line breaks separate words in Word even though they do
    # not live inside w:t nodes.  Keep non-breaking hyphens semantically a
    # hyphen and omit optional soft hyphens.
    separator_token = "\ue000"
    no_break_hyphen_token = "\ue001"
    # Empty OOXML controls may be serialized either as ``<w:tab/>`` or as a
    # paired empty element such as ``<w:tab></w:tab>``.  Normalize both forms
    # before collecting text so they have identical visible semantics.
    visible = re.sub(
        r"<w:(tab|br|cr)\b[^>]*>\s*</w:\1\s*>",
        separator_token,
        visible,
    )
    visible = re.sub(r"<w:(?:tab|br|cr)\b[^>]*/\s*>", separator_token, visible)
    visible = re.sub(
        r"<w:noBreakHyphen\b[^>]*>\s*</w:noBreakHyphen\s*>",
        no_break_hyphen_token,
        visible,
    )
    visible = re.sub(
        r"<w:noBreakHyphen\b[^>]*/\s*>",
        no_break_hyphen_token,
        visible,
    )
    visible = re.sub(
        r"<w:softHyphen\b[^>]*>\s*</w:softHyphen\s*>",
        "",
        visible,
    )
    visible = re.sub(r"<w:softHyphen\b[^>]*/\s*>", "", visible)
    pieces = re.findall(
        rf"<w:t\b[^>]*>([\s\S]*?)</w:t>|({separator_token})|({no_break_hyphen_token})",
        visible,
    )
    if not pieces:
        return ""
    joined = html.unescape(
        "".join(
            text if text else (" " if separator else "\u2011")
            for text, separator, _hyphen in pieces
        )
    )
    joined = re.sub(r"\s+", " ", joined).strip()
    return joined


def paragraph_contains_sectpr(p_xml: str) -> bool:
    return "<w:sectPr" in _mask_text_box_subtrees(p_xml)


def paragraph_pstyle_from_block(p_xml: str) -> Optional[str]:
    m = re.search(r"<w:pStyle\b[^>]*w:val=\"([^\"]+)\"", _mask_text_box_subtrees(p_xml))
    return m.group(1) if m else None


def paragraph_numpr_from_block(p_xml: str) -> Dict[str, Optional[str]]:
    numId = None
    ilvl = None
    outer_xml = _mask_text_box_subtrees(p_xml)
    m1 = re.search(r"<w:numId\b[^>]*w:val=\"([^\"]+)\"", outer_xml)
    m2 = re.search(r"<w:ilvl\b[^>]*w:val=\"([^\"]+)\"", outer_xml)
    if m1:
        numId = m1.group(1)
    if m2:
        ilvl = m2.group(1)
    return {"numId": numId, "ilvl": ilvl}


def paragraph_ppr_hints_from_block(p_xml: str) -> Dict[str, Any]:
    # lightweight hints (alignment + ind + spacing)
    outer_xml = _mask_text_box_subtrees(p_xml)
    hints: Dict[str, Any] = {}
    m = re.search(r"<w:jc\b[^>]*w:val=\"([^\"]+)\"", outer_xml)
    if m:
        hints["jc"] = m.group(1)

    ind = {}
    for k in ["left", "right", "firstLine", "hanging"]:
        m2 = re.search(rf"<w:ind\b[^>]*w:{k}=\"([^\"]+)\"", outer_xml)
        if m2:
            ind[k] = m2.group(1)
    if ind:
        hints["ind"] = ind

    spacing = {}
    for k in ["before", "after", "line", "lineRule"]:
        m3 = re.search(rf"<w:spacing\b[^>]*w:{k}=\"([^\"]+)\"", outer_xml)
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


def collect_paragraph_contexts(document_xml_path: Path) -> List[ParagraphContext]:
    # ElementTree supplies a strict well-formedness check without rewriting
    # the source bytes; the lexical scanner supplies exact byte ranges.
    ET.parse(document_xml_path)
    document_xml_text = read_xml_text(document_xml_path)
    return [item.context for item in _scan_paragraph_xml_blocks(document_xml_text)]


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

    # Validate the complete part, then use one lexical scan for exact ranges
    # and structural context.  This avoids the old regex/DOM index mismatch
    # when a drawing contains nested text-box paragraphs.
    ET.parse(doc_path)
    blocks = _scan_paragraph_xml_blocks(doc_text)

    for idx, block in enumerate(blocks):
        p_xml = block.xml
        ctx = block.context
        analysis_xml = _strip_out_of_scope_subtrees(p_xml)
        raw_text = paragraph_text_from_block(p_xml)
        pStyle = paragraph_pstyle_from_block(analysis_xml)
        numpr = paragraph_numpr_from_block(analysis_xml)
        hints = paragraph_ppr_hints_from_block(analysis_xml)
        rpr_hints = paragraph_rpr_hints_from_block(analysis_xml)
        has_direct_ppr = bool(extract_paragraph_ppr_inner(analysis_xml))
        has_uniform_direct_rpr = bool(extract_paragraph_rpr_inner(analysis_xml))
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
        # A drawing/text box is itself out of scope, but it does not make
        # visible host text out of scope.  Empty host paragraphs remain
        # explicitly reported as drawing/text-box containers.
        if not raw_text.strip():
            if ctx.contains_textbox:
                skip_reason = "text_box"
            elif ctx.contains_drawing:
                skip_reason = "drawing"

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
            "contains_drawing": ctx.contains_drawing,
            "contains_textbox": ctx.contains_textbox,
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
    outer_xml = _mask_text_box_subtrees(p_xml)
    pstyle = re.search(r"<w:pStyle\b[^>]*/\s*>", outer_xml, flags=re.S)
    result = p_xml
    if pstyle:
        result = p_xml[:pstyle.start()] + p_xml[pstyle.end():]

    # Also remove an outer empty pPr that might result. Nested text-box pPr
    # nodes are a separate story and must remain byte-identical.
    outer_result = _mask_text_box_subtrees(result)
    empty_ppr = re.search(
        r"<w:pPr\b[^>]*>\s*</w:pPr\s*>|<w:pPr\b[^>]*/\s*>",
        outer_result,
        flags=re.S,
    )
    if empty_ppr:
        result = result[:empty_ppr.start()] + result[empty_ppr.end():]
    return result

def ppr_without_pstyle(p_xml: str) -> str:
    outer_xml = _mask_text_box_subtrees(p_xml)
    if re.search(r"<w:pPr\b[^>]*/\s*>", outer_xml, flags=re.S):
        return ""
    m = re.search(
        r"<w:pPr\b[^>]*>(?P<inner>.*?)</w:pPr\s*>",
        outer_xml,
        flags=re.S,
    )
    if not m:
        return ""
    ppr = p_xml[m.start():m.end()]
    ppr = re.sub(r"<w:pStyle\b[^>]*/\s*>", "", ppr, count=1)
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
    outer_xml = _mask_text_box_subtrees(p_xml)
    if re.search(r"<w:pPr\b[^>]*/>", outer_xml):
        return ""
    m = re.search(r"<w:pPr\b[^>]*>(.*?)</w:pPr>", outer_xml, flags=re.S)
    if not m:
        return ""
    inner = p_xml[m.start(1):m.end(1)]
    inner = re.sub(r"<w:pStyle\b[^>]*/>", "", inner)
    # A visible paragraph may also carry a section break.  It is part of the
    # document structure, not reusable paragraph formatting: copying it into
    # a derived style would leak document relationship IDs into styles.xml and
    # can make the emitted package invalid.  The original paragraph-level
    # sectPr remains untouched in document.xml.
    for sectpr_block in _extract_named_xml_blocks(inner, "w:sectPr"):
        inner = inner.replace(sectpr_block, "", 1)
    return inner.strip()


def _xml_element_signature(element: ET.Element) -> Tuple[Any, ...]:
    """Return an attribute-order-independent signature for a small XML tree."""
    return (
        element.tag,
        tuple(sorted(element.attrib.items())),
        (element.text or "").strip(),
        tuple(_xml_element_signature(child) for child in list(element)),
    )


_XML_QNAME = r"[A-Za-z_][\w.-]*(?::[A-Za-z_][\w.-]*)?"
_XML_PREFIX_RE = re.compile(
    r"(?:</?\s*|\s)([A-Za-z_][\w.-]*):[A-Za-z_][\w.-]*(?=\s|/?>|=)"
)
_XMLNS_ATTRIBUTE_RE = re.compile(
    r"\sxmlns:(?P<prefix>[A-Za-z_][\w.-]*)\s*=\s*"
    r"(?P<quote>['\"])(?P<uri>.*?)(?P=quote)",
    flags=re.S,
)


def _xml_markup_end(xml_text: str, start: int) -> int:
    """Return the exclusive end of markup beginning at *start*."""
    if xml_text.startswith("<!--", start):
        end = xml_text.find("-->", start + 4)
        return len(xml_text) if end == -1 else end + 3
    if xml_text.startswith("<![CDATA[", start):
        end = xml_text.find("]]>", start + 9)
        return len(xml_text) if end == -1 else end + 3
    if xml_text.startswith("<?", start):
        end = xml_text.find("?>", start + 2)
        return len(xml_text) if end == -1 else end + 2

    quote: Optional[str] = None
    for index in range(start + 1, len(xml_text)):
        char = xml_text[index]
        if quote is not None:
            if char == quote:
                quote = None
            continue
        if char in {'"', "'"}:
            quote = char
        elif char == ">":
            return index + 1
    return len(xml_text)


def _direct_xml_child_fragments(xml_text: str) -> List[str]:
    """Return exact top-level element fragments from an XML inner string.

    Run properties may be in Word extension namespaces (for example w14) and
    may themselves contain nested elements.  A QName-aware lexical walk keeps
    those fragments byte-for-byte without assuming a particular prefix.
    """
    fragments: List[str] = []
    element_stack: List[str] = []
    fragment_start: Optional[int] = None
    position = 0

    while True:
        start = xml_text.find("<", position)
        if start == -1:
            break
        end = _xml_markup_end(xml_text, start)
        if end <= start or end > len(xml_text):
            return []
        token = xml_text[start:end]
        position = end

        if token.startswith(("<!--", "<![CDATA[", "<?", "<!")):
            continue

        if token.startswith("</"):
            match = re.match(rf"</\s*({_XML_QNAME})\s*>", token, flags=re.S)
            if not match or not element_stack or element_stack[-1] != match.group(1):
                return []
            element_stack.pop()
            if not element_stack and fragment_start is not None:
                fragments.append(xml_text[fragment_start:end].strip())
                fragment_start = None
            continue

        match = re.match(rf"<\s*({_XML_QNAME})(?=\s|/?>)", token, flags=re.S)
        if not match:
            return []
        if not element_stack:
            fragment_start = start
        if re.search(r"/\s*>$", token):
            if not element_stack and fragment_start is not None:
                fragments.append(xml_text[fragment_start:end].strip())
                fragment_start = None
        else:
            element_stack.append(match.group(1))

    return fragments if not element_stack else []


def _fragment_namespace_declarations(xml_text: str) -> str:
    """Declare prefixes synthetically so ElementTree can compare fragments."""
    prefixes = set(_XML_PREFIX_RE.findall(xml_text)) - {"xml", "xmlns"}
    declarations = []
    for prefix in sorted(prefixes):
        uri = W_NS if prefix == "w" else f"urn:phase1-ooxml-prefix:{prefix}"
        declarations.append(f'xmlns:{prefix}="{xml_escape(uri)}"')
    return " ".join(declarations)


def _run_property_fragments(rpr_inner: str) -> List[Tuple[Tuple[Any, ...], str]]:
    """Split direct w:rPr children while preserving their original XML bytes."""
    fragments: List[Tuple[Tuple[Any, ...], str]] = []
    for raw in _direct_xml_child_fragments(rpr_inner):
        try:
            declarations = _fragment_namespace_declarations(raw)
            wrapper = ET.fromstring(f"<root {declarations}>{raw}</root>")
            if not list(wrapper):
                continue
            signature = _xml_element_signature(list(wrapper)[0])
        except ET.ParseError:
            # Keep a conservative signature rather than throwing away a
            # property that is demonstrably identical in every visible run.
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
    visible_xml = _mask_text_box_subtrees(p_xml)
    visible_xml = re.sub(r"<w:(?:del|moveFrom)\b[^>]*>[\s\S]*?</w:(?:del|moveFrom)>", "", visible_xml)
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


def _root_opening_tag_span(xml_text: str) -> Optional[Tuple[int, int]]:
    """Locate the document element's opening tag without reserializing XML."""
    position = 0
    while True:
        start = xml_text.find("<", position)
        if start == -1:
            return None
        end = _xml_markup_end(xml_text, start)
        token = xml_text[start:end]
        position = end
        if token.startswith(("<?", "<!--", "<!")):
            continue
        if token.startswith("</") or not re.match(rf"<\s*{_XML_QNAME}(?=\s|/?>)", token):
            return None
        return start, end


def _root_namespace_declarations(xml_text: str) -> Dict[str, str]:
    span = _root_opening_tag_span(xml_text)
    if span is None:
        return {}
    opening_tag = xml_text[span[0]:span[1]]
    return {
        match.group("prefix"): html.unescape(match.group("uri"))
        for match in _XMLNS_ATTRIBUTE_RE.finditer(opening_tag)
    }


def _propagate_style_fragment_namespaces(
    styles_xml_text: str,
    document_xml_text: str,
    style_blocks: List[str],
) -> str:
    """Copy namespaces needed by derived fragments onto the styles root."""
    if not style_blocks:
        return styles_xml_text

    used_prefixes = set(_XML_PREFIX_RE.findall("\n".join(style_blocks))) - {"xml", "xmlns"}
    styles_namespaces = _root_namespace_declarations(styles_xml_text)
    document_namespaces = _root_namespace_declarations(document_xml_text)
    additions = {
        prefix: document_namespaces[prefix]
        for prefix in used_prefixes - set(styles_namespaces)
        if prefix in document_namespaces
    }
    if not additions:
        return styles_xml_text

    span = _root_opening_tag_span(styles_xml_text)
    if span is None:
        raise ValueError("styles.xml does not contain a valid root opening tag")
    opening_tag = styles_xml_text[span[0]:span[1]]
    if re.search(r"/\s*>$", opening_tag):
        raise ValueError("styles.xml root must not be self-closing")
    declarations = "".join(
        f' xmlns:{prefix}="{xml_escape(uri)}"'
        for prefix, uri in sorted(additions.items())
    )
    updated_opening_tag = opening_tag[:-1] + declarations + ">"
    return styles_xml_text[:span[0]] + updated_opening_tag + styles_xml_text[span[1]:]


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

    portable = _propagate_style_fragment_namespaces(
        insert_styles_into_styles_xml(styles_text, derived_blocks),
        doc_text,
        derived_blocks,
    )
    portable = prepare_xml_text_for_utf8(portable)
    try:
        ET.fromstring(portable)
    except ET.ParseError as exc:
        raise ValueError(
            f"Generated portable styles.xml is not well-formed after namespace propagation: {exc}"
        ) from exc
    style_ids = set(re.findall(r'w:styleId="([^"]+)"', portable))
    for item in instructions.get("apply_pStyle", []) or []:
        if item["styleId"] not in style_ids:
            raise ValueError(f"apply_pStyle references unknown styleId: {item['styleId']}")
    return portable


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
    outer_xml = _mask_text_box_subtrees(p_xml)
    escaped_style_id = xml_escape(styleId)
    pstyle = re.search(
        r"<w:pStyle\b[^>]*\bw:val=(?P<quote>['\"])(?P<value>[^'\"]*)(?P=quote)",
        outer_xml,
        flags=re.S,
    )
    if pstyle:
        return p_xml[:pstyle.start("value")] + escaped_style_id + p_xml[pstyle.end("value"):]

    self_closing_ppr = re.search(r"<w:pPr\b[^>]*/\s*>", outer_xml, flags=re.S)
    if self_closing_ppr:
        return (
            p_xml[:self_closing_ppr.start()]
            + f'<w:pPr><w:pStyle w:val="{escaped_style_id}"/></w:pPr>'
            + p_xml[self_closing_ppr.end():]
        )

    opening_ppr = re.search(r"<w:pPr\b[^>]*>", outer_xml, flags=re.S)
    if opening_ppr:
        return (
            p_xml[:opening_ppr.end()]
            + f'<w:pStyle w:val="{escaped_style_id}"/>'
            + p_xml[opening_ppr.end():]
        )

    opening_p = re.search(r"<w:p\b[^>]*>", outer_xml, flags=re.S)
    if not opening_p:
        return p_xml
    return (
        p_xml[:opening_p.end()]
        + f'<w:pPr><w:pStyle w:val="{escaped_style_id}"/></w:pPr>'
        + p_xml[opening_p.end():]
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

    # 2) Apply paragraph styles by index (pStyle insertion ONLY)
    idx_map: Dict[int, str] = {}
    for item in (instructions.get("apply_pStyle") or []):
        idx_map[int(item["paragraph_index"])] = item["styleId"]

    original_ppr = {i: ppr_without_pstyle(pb) for i, pb in enumerate(para_blocks)}

    for idx, sid in idx_map.items():
        if idx < 0 or idx >= len(para_blocks):
            raise ValueError(f"paragraph_index out of range: {idx}")
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
