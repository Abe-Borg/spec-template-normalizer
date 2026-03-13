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
import zipfile
from collections import Counter
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Dict, List, Optional, Set, Tuple
import html

import xml.etree.ElementTree as ET  # only used for styles.xml name lookup + optional catalogs


# -----------------------------
# Utilities
# -----------------------------

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"

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
    doc_text = doc_path.read_text(encoding="utf-8")
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

    doc_text = (extract_dir / "word" / "document.xml").read_text(encoding="utf-8")
    current_sectpr = extract_sectpr_block(doc_text)
    if sha256_text(current_sectpr) != snap.sectpr_hash:
        raise ValueError("Section properties (w:sectPr) stability check FAILED.")

    current_rels = snapshot_doc_rels_hash(extract_dir)
    if current_rels != snap.doc_rels_hash:
        raise ValueError("document.xml.rels stability check FAILED (can break header/footer binding).")


# -----------------------------
# DOCX extraction (workspace only)
# -----------------------------

def extract_docx(docx_path: Path, extract_dir: Path) -> None:
    if extract_dir.exists():
        # OneDrive-safe deletion: retry with delay if locked
        import time
        max_retries = 3
        for attempt in range(max_retries):
            try:
                shutil.rmtree(extract_dir)
                break
            except PermissionError as e:
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
    with zipfile.ZipFile(docx_path, "r") as zin:
        zin.extractall(extract_dir)


# -----------------------------
# Slim bundle construction (LLM input)
# -----------------------------

def iter_paragraph_xml_blocks(document_xml_text: str):
    # Non-greedy paragraph blocks. Avoid full XML parse to preserve indices + raw text.
    for m in re.finditer(r"(<w:p\b[\s\S]*?</w:p>)", document_xml_text):
        yield m.start(), m.end(), m.group(1)


def paragraph_text_from_block(p_xml: str) -> str:
    texts = re.findall(r"<w:t\b[^>]*>([\s\S]*?)</w:t>", p_xml)
    if not texts:
        return ""
    joined = html.unescape("".join(texts))
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


def paragraph_is_in_table(start: int, table_spans: List[Tuple[int, int]]) -> bool:
    return any(t_start <= start < t_end for t_start, t_end in table_spans)


def build_style_catalog(styles_xml_path: Path, used_style_ids: Set[str]) -> Dict[str, Any]:
    # Compact info for used styles + basedOn chain (safe to keep; not required, but useful)
    tree = ET.parse(styles_xml_path)
    root = tree.getroot()

    styles_by_id: Dict[str, ET.Element] = {}
    for st in root.findall(f".//{_q('style')}"):
        sid = _get_attr(st, "styleId")
        if sid:
            styles_by_id[sid] = st

    to_include = set(used_style_ids)
    changed = True
    while changed:
        changed = False
        for sid in list(to_include):
            st = styles_by_id.get(sid)
            if st is None:
                continue
            based = st.find(_q("basedOn"))
            if based is not None:
                base_id = _get_attr(based, "val")
                if base_id and base_id not in to_include:
                    to_include.add(base_id)
                    changed = True

    catalog: Dict[str, Any] = {}
    for sid in sorted(to_include):
        st = styles_by_id.get(sid)
        if st is None:
            continue
        name_el = st.find(_q("name"))
        based_el = st.find(_q("basedOn"))
        catalog[sid] = {
            "styleId": sid,
            "type": _get_attr(st, "type"),
            "name": _get_attr(name_el, "val") if name_el is not None else None,
            "basedOn": _get_attr(based_el, "val") if based_el is not None else None,
        }
    return catalog


def build_numbering_catalog(numbering_xml_path: Path, used_num_ids: Set[str]) -> Dict[str, Any]:
    # Read-only; we never edit numbering.xml. Provide minimal mapping to LLM.
    if not numbering_xml_path.exists():
        return {"nums": {}, "abstracts": {}}

    tree = ET.parse(numbering_xml_path)
    root = tree.getroot()

    num_map: Dict[str, str] = {}
    for num in root.findall(f".//{_q('num')}"):
        numId = _get_attr(num, "numId")
        abs_el = num.find(_q("abstractNumId"))
        if numId and abs_el is not None:
            absId = _get_attr(abs_el, "val")
            if absId:
                num_map[numId] = absId

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
        nums[numId] = {"numId": numId, "abstractNumId": num_map.get(numId)}

    return {"nums": nums, "abstracts": abstracts}


def build_slim_bundle(extract_dir: Path) -> Dict[str, Any]:
    snap = snapshot_stability(extract_dir)

    doc_path = extract_dir / "word" / "document.xml"
    doc_text = doc_path.read_text(encoding="utf-8")

    paragraphs = []
    used_style_ids: Set[str] = set()
    used_num_ids: Set[str] = set()

    table_spans = [(m.start(), m.end()) for m in re.finditer(r"<w:tbl\b[\s\S]*?</w:tbl>", doc_text)]

    for idx, (_s, _e, p_xml) in enumerate(iter_paragraph_xml_blocks(doc_text)):
        txt = paragraph_text_from_block(p_xml)
        pStyle = paragraph_pstyle_from_block(p_xml)
        numpr = paragraph_numpr_from_block(p_xml)
        hints = paragraph_ppr_hints_from_block(p_xml)
        rpr_hints = paragraph_rpr_hints_from_block(p_xml)
        contains_sect = paragraph_contains_sectpr(p_xml)
        in_table = paragraph_is_in_table(_s, table_spans)

        if pStyle:
            used_style_ids.add(pStyle)
        if numpr.get("numId"):
            used_num_ids.add(numpr["numId"])

        if len(txt) > 200:
            txt = txt[:200] + "…"

        paragraphs.append({
            "paragraph_index": idx,
            "text": txt,
            "pStyle": pStyle,
            "numPr": numpr if (numpr.get("numId") or numpr.get("ilvl")) else None,
            "pPr_hints": hints if hints else None,
            "rPr_hints": rpr_hints if rpr_hints else None,
            "contains_sectPr": contains_sect,
            "in_table": in_table,
        })

    style_catalog: Dict[str, Any] = {}
    styles_path = extract_dir / "word" / "styles.xml"
    if styles_path.exists():
        style_catalog = build_style_catalog(styles_path, used_style_ids)

    numbering_catalog = build_numbering_catalog(extract_dir / "word" / "numbering.xml", used_num_ids)

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
    inner = re.sub(r"<w:numPr\b[^>]*>.*?</w:numPr>", "", inner, flags=re.S)
    return inner.strip()


def extract_paragraph_rpr_inner(p_xml: str) -> str:
    """Extract the representative run-level rPr inner XML from a paragraph.

    For multi-run paragraphs, selects the most common non-empty rPr
    (after canonicalization for comparison), breaking ties by first
    occurrence.  Returns the original (non-canonicalized) rPr to
    preserve all formatting information.
    """
    originals: List[str] = []
    canon_keys: List[str] = []

    for rm in re.finditer(r"<w:r\b[^>]*>(.*?)</w:r>", p_xml, flags=re.S):
        run_inner = rm.group(1)
        text_nodes = re.findall(r"<w:t\b[^>]*>([\s\S]*?)</w:t>", run_inner)
        if not text_nodes:
            continue
        run_text = html.unescape("".join(text_nodes))
        if not run_text.strip():
            continue
        m = re.search(r"<w:rPr\b[^>]*>(.*?)</w:rPr>", run_inner, flags=re.S)
        if not m:
            continue
        raw = m.group(1).strip()
        if not raw:
            continue
        canon = _strip_proofing_for_cmp(_strip_rsids_for_cmp(raw)).strip()
        if not canon:
            continue
        originals.append(raw)
        canon_keys.append(canon)

    if not originals:
        return ""

    if len(set(canon_keys)) == 1:
        return originals[0]

    # Multiple distinct forms: pick the most common, tie-break by first occurrence
    counts = Counter(canon_keys)
    max_count = counts.most_common(1)[0][1]
    best_idx = len(originals)  # sentinel
    for canon_val, cnt in counts.items():
        if cnt == max_count:
            first_idx = canon_keys.index(canon_val)
            if first_idx < best_idx:
                best_idx = first_idx

    return originals[best_idx]


def derive_style_def_from_paragraph(styleId: str, name: str, p_xml: str, based_on: Optional[str] = None) -> Dict[str, Any]:
    return {
        "styleId": styleId,
        "name": name,
        "type": "paragraph",
        "basedOn": based_on,
        "pPr_inner": extract_paragraph_ppr_inner(p_xml),
        "rPr_inner": extract_paragraph_rpr_inner(p_xml),
    }


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
    allowed_keys = {"create_styles", "apply_pStyle", "roles", "notes"}
    extra = set(instructions.keys()) - allowed_keys
    if extra:
        raise ValueError(f"Invalid instruction keys: {extra}")

    allowed_new_style_ids: Set[str] = {
        "CSI_SectionTitle__ARCH",
        "CSI_SectionID__ARCH",
        "CSI_Part__ARCH",
        "CSI_Article__ARCH",
        "CSI_Paragraph__ARCH",
        "CSI_Subparagraph__ARCH",
        "CSI_Subsubparagraph__ARCH",
    }

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

    roles = instructions.get("roles")
    if roles is None or not isinstance(roles, dict):
        raise ValueError("Missing/invalid required key: roles")

    allowed_roles = {"SectionID","SectionTitle","PART","ARTICLE","PARAGRAPH","SUBPARAGRAPH","SUBSUBPARAGRAPH"}
    for role, spec in roles.items():
        if role not in allowed_roles:
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
    all_allowed_style_ids = allowed_existing_style_ids | allowed_new_style_ids

    for ap in instructions.get("apply_pStyle", []) or []:
        idx = ap["paragraph_index"]
        sid = ap["styleId"]
        if idx >= paragraph_count:
            raise ValueError(f"apply_pStyle paragraph_index out of range: {idx} >= {paragraph_count}")
        if sid not in all_allowed_style_ids:
            raise ValueError(f"apply_pStyle[{idx}] uses disallowed styleId '{sid}'")

    for role, spec in roles.items():
        sid = spec["styleId"]
        ex = int(spec["exemplar_paragraph_index"])
        if sid not in all_allowed_style_ids:
            raise ValueError(f"roles['{role}'] uses disallowed styleId '{sid}'")
        if ex >= paragraph_count:
            raise ValueError(f"roles['{role}'].exemplar_paragraph_index out of range: {ex} >= {paragraph_count}")
        para = paragraphs[ex]
        text = (para.get("text") or "").strip()
        if not text:
            raise ValueError(f"roles['{role}'] exemplar paragraph {ex} is blank")
        if para.get("contains_sectPr"):
            raise ValueError(f"roles['{role}'] exemplar paragraph {ex} contains sectPr")
        if para.get("in_table"):
            raise ValueError(f"roles['{role}'] exemplar paragraph {ex} is inside a table")
        if text.upper() == "END OF SECTION":
            raise ValueError(f"roles['{role}'] exemplar paragraph {ex} cannot be END OF SECTION")
        if text.startswith("[") and text.endswith("]"):
            raise ValueError(f"roles['{role}'] exemplar paragraph {ex} cannot be an editor note")

    required_apply_indices: List[int] = []
    for p in paragraphs:
        text = (p.get("text") or "").strip()
        if not text or p.get("contains_sectPr") or p.get("in_table"):
            continue
        if text.upper() == "END OF SECTION":
            continue
        if text.startswith("[") and text.endswith("]"):
            continue
        required_apply_indices.append(int(p["paragraph_index"]))

    apply_indices = [int(ap["paragraph_index"]) for ap in instructions.get("apply_pStyle", []) or []]
    apply_set = set(apply_indices)
    required_set = set(required_apply_indices)
    missing = sorted(required_set - apply_set)
    unexpected = sorted(apply_set - required_set)
    if missing or unexpected:
        raise ValueError(
            "apply_pStyle coverage mismatch; "
            f"missing={missing[:20]}{'...' if len(missing) > 20 else ''}, "
            f"unexpected={unexpected[:20]}{'...' if len(unexpected) > 20 else ''}"
        )


def apply_instructions(extract_dir: Path, instructions: Dict[str, Any]) -> None:
    slim_bundle = build_slim_bundle(extract_dir)
    validate_instructions(instructions, slim_bundle=slim_bundle)

    snap = snapshot_stability(extract_dir)

    styles_path = extract_dir / "word" / "styles.xml"
    styles_text = styles_path.read_text(encoding="utf-8")

    doc_path = extract_dir / "word" / "document.xml"
    doc_text = doc_path.read_text(encoding="utf-8")

    blocks = list(iter_paragraph_xml_blocks(doc_text))
    para_blocks = [b[2] for b in blocks]
    original_para_blocks = list(para_blocks)

    # 1) Create derived styles (local formatting capture from exemplar paragraphs)
    style_defs = instructions.get("create_styles") or []
    derived_blocks: List[str] = []
    for sd in style_defs:
        style_id = sd["styleId"]
        style_name = sd.get("name") or style_id
        src_idx = sd["derive_from_paragraph_index"]
        based_on = sd.get("basedOn")
        if src_idx >= len(para_blocks):
            raise ValueError(f"Style {style_id}: derive_from_paragraph_index out of range: {src_idx}")
        exemplar_p = para_blocks[src_idx]
        derived_def = derive_style_def_from_paragraph(style_id, style_name, exemplar_p, based_on=based_on)
        derived_blocks.append(build_style_xml_block(derived_def))

    styles_new = insert_styles_into_styles_xml(styles_text, derived_blocks)
    if styles_new != styles_text:
        styles_path.write_text(styles_new, encoding="utf-8")

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
    doc_path.write_text(doc_new, encoding="utf-8")

    verify_stability(extract_dir, snap)


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


def build_style_registry_dict(extract_dir: Path, source_docx_name: str, instructions: Dict[str, Any]) -> Dict[str, Any]:
    """Build the arch_style_registry payload dict without writing to disk."""
    roles = instructions.get("roles") or {}
    styles_path = extract_dir / "word" / "styles.xml"
    name_map = _build_style_name_map(styles_path)
    bundle = build_slim_bundle(extract_dir)
    paragraphs = bundle.get("paragraphs", [])
    created_style_ids = {sd.get("styleId") for sd in (instructions.get("create_styles") or []) if isinstance(sd, dict)}

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
            if style_id in created_style_ids and exemplar.get("pStyle"):
                entry["warning"] = "Derived style exemplar has existing pStyle; inheritance chain may affect formatting"

        out_roles[role] = entry

    return {
        "version": 1,
        "source_docx": source_docx_name,
        "roles": out_roles,
    }


def emit_arch_style_registry(extract_dir: Path, source_docx_name: str, instructions: Dict[str, Any], out_path: Optional[Path] = None) -> Path:
    payload = build_style_registry_dict(extract_dir, source_docx_name, instructions)

    if out_path is None:
        out_path = extract_dir / "arch_style_registry.json"
    out_path.write_text(json.dumps(payload, indent=2), encoding="utf-8")
    return out_path
