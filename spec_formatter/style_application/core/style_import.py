"""
Style extraction, materialization, and import for Phase 2.

Handles importing architect styles into target documents with
property materialization for cross-document portability.
"""

import re
from pathlib import Path
from typing import Dict, List, Set, Optional

from .ooxml_text import read_xml_text, write_xml_text

# Word built-in styles that exist implicitly in every DOCX.
# These never need to be imported — Word creates them internally
# even when styles.xml has no explicit <w:style> block for them.
WORD_BUILTIN_STYLE_IDS = frozenset({
    "Normal",
    "DefaultParagraphFont",
    "TableNormal",
    "NoList",
})

def _extract_style_block(styles_xml_text: str, style_id: str) -> Optional[str]:
    m = re.search(
        rf'(<w:style\b[^>]*w:styleId="{re.escape(style_id)}"[\s\S]*?</w:style>)',
        styles_xml_text,
        flags=re.S
    )
    return m.group(1) if m else None

def _extract_basedOn(style_block: str) -> Optional[str]:
    m = re.search(r'<w:basedOn\b[^>]*w:val="([^"]+)"', style_block)
    return m.group(1) if m else None

def _extract_numpr_block(style_block: str) -> Optional[str]:
    m = re.search(r'(<w:numPr\b[^>]*>[\s\S]*?</w:numPr>)', style_block, flags=re.S)
    return m.group(1) if m else None

def _find_style_numpr_in_chain(styles_xml_text: str, style_id: str, max_hops: int = 50) -> Optional[str]:
    seen = set()
    cur = style_id
    hops = 0
    while cur and cur not in seen and hops < max_hops:
        seen.add(cur)
        hops += 1
        block = _extract_style_block(styles_xml_text, cur)
        if not block:
            break
        numpr = _extract_numpr_block(block)
        if numpr:
            return numpr
        cur = _extract_basedOn(block)
    return None


def ensure_explicit_numpr_from_current_style(
    paragraph_xml: str,
    current_styles_xml: str,
) -> str:
    """Materialize numbering inherited from the paragraph's current style.

    A pStyle swap otherwise silently drops numbering when the destination style
    has no numbering of its own.  Existing direct numPr and section-property
    paragraphs are deliberately left untouched.
    """
    if "<w:numPr" in paragraph_xml or "<w:sectPr" in paragraph_xml:
        return paragraph_xml
    pstyle_match = re.search(
        r'<w:pStyle\b[^>]*w:val="([^"]+)"',
        paragraph_xml,
    )
    if not pstyle_match:
        return paragraph_xml
    numpr = _find_style_numpr_in_chain(current_styles_xml, pstyle_match.group(1))
    if not numpr:
        return paragraph_xml
    if re.search(r"<w:pPr\b[^>]*>", paragraph_xml):
        return re.sub(
            r"(<w:pPr\b[^>]*>)",
            rf"\1{numpr}",
            paragraph_xml,
            count=1,
        )
    return re.sub(
        r"(<w:p\b[^>]*>)",
        rf"\1<w:pPr>{numpr}</w:pPr>",
        paragraph_xml,
        count=1,
    )

def _strip_pstyle_and_numpr(ppr_inner: str) -> str:
    if not ppr_inner:
        return ""
    out = re.sub(r"<w:pStyle\b[^>]*/>", "", ppr_inner)
    out = re.sub(r"<w:numPr\b[^>]*>[\s\S]*?</w:numPr>", "", out, flags=re.S)
    return out.strip()

def _extract_tag_inner(xml: str, tag: str) -> Optional[str]:
    m = re.search(rf"<{tag}\b[^>]*>([\s\S]*?)</{tag}>", xml, flags=re.S)
    return m.group(1) if m else None

def _docdefaults_rpr_inner(styles_xml_text: str) -> str:
    m = re.search(
        r"<w:docDefaults\b[\s\S]*?<w:rPrDefault\b[\s\S]*?<w:rPr\b[^>]*>([\s\S]*?)</w:rPr>[\s\S]*?</w:rPrDefault>",
        styles_xml_text,
        flags=re.S
    )
    return m.group(1).strip() if m else ""

def _docdefaults_ppr_inner(styles_xml_text: str) -> str:
    m = re.search(
        r"<w:docDefaults\b[\s\S]*?<w:pPrDefault\b[\s\S]*?<w:pPr\b[^>]*>([\s\S]*?)</w:pPr>[\s\S]*?</w:pPrDefault>",
        styles_xml_text,
        flags=re.S
    )
    return _strip_pstyle_and_numpr(m.group(1).strip()) if m else ""

def _effective_rpr_inner_in_arch(arch_styles_xml_text: str, style_id: str) -> str:
    """
    Return a *minimal* effective rPr inner XML for the FORCE typography set only.

    We resolve each child tag independently through the basedOn chain, then fall back
    to docDefaults. This avoids the bug where a derived style contains <w:rPr> but
    doesn't specify (for example) <w:rFonts>, causing inherited font settings to be missed.
    """
    force_tags = ("rFonts", "sz", "szCs", "lang")

    def _extract_child_node(inner_xml: str, tag: str) -> Optional[str]:
        if not inner_xml:
            return None
        # Self-closing: <w:tag .../>
        m = re.search(rf"(<w:{re.escape(tag)}\b[^>]*/>)", inner_xml)
        if m:
            return m.group(1)
        # Paired: <w:tag ...>...</w:tag>
        m = re.search(
            rf"(<w:{re.escape(tag)}\b[^>]*>[\s\S]*?</w:{re.escape(tag)}>)",
            inner_xml,
            flags=re.S
        )
        if m:
            return m.group(1)
        return None

    def _resolve(tag: str) -> Optional[str]:
        seen = set()
        cur = style_id
        hops = 0
        while cur and cur not in seen and hops < 50:
            seen.add(cur)
            hops += 1
            blk = _extract_style_block(arch_styles_xml_text, cur)
            if not blk:
                break
            rpr_inner = _extract_tag_inner(blk, "w:rPr") or ""
            node = _extract_child_node(rpr_inner, tag)
            if node:
                return node
            cur = _extract_basedOn(blk)

        # fall back to docDefaults
        docdef_inner = _docdefaults_rpr_inner(arch_styles_xml_text)
        return _extract_child_node(docdef_inner, tag)

    nodes: List[str] = []
    for t in force_tags:
        node = _resolve(t)
        if node:
            nodes.append(node)

    return "".join(nodes)

def _effective_ppr_inner_in_arch(arch_styles_xml_text: str, style_id: str) -> str:
    seen = set()
    cur = style_id
    hops = 0
    while cur and cur not in seen and hops < 50:
        seen.add(cur); hops += 1
        blk = _extract_style_block(arch_styles_xml_text, cur)
        if not blk:
            break
        inner = _extract_tag_inner(blk, "w:pPr") or ""
        inner = _strip_pstyle_and_numpr(inner)
        if inner:
            return inner
        cur = _extract_basedOn(blk)
    return _docdefaults_ppr_inner(arch_styles_xml_text)

def _rpr_contains_tag(rpr_inner: str, tag: str) -> bool:
    return re.search(rf"<w:{re.escape(tag)}\b", rpr_inner) is not None

def _extract_rpr_inner(style_block: str) -> Optional[str]:
    return _extract_tag_inner(style_block, "w:rPr")

def _inject_missing_rpr_children(style_block: str, missing_children_xml: str) -> str:
    """Insert missing rPr children (already as raw XML) just before </w:rPr>."""
    if not missing_children_xml.strip():
        return style_block
    if "</w:rPr>" not in style_block:
        return style_block
    # Replace only the first closing tag (avoid accidental insertion into nested rPr blocks)
    return style_block.replace("</w:rPr>", f"{missing_children_xml}</w:rPr>", 1)

def _materialize_minimal_typography(style_block: str, style_id: str, arch_styles_xml_text: str) -> str:
    """
    Make imported styles resilient across documents by ensuring a minimal set of
    typography-related rPr children exist (fonts, sizes, language).

    IMPORTANT:
    - Does NOT invent values.
    - Only copies missing nodes from the *effective* arch style chain + docDefaults.
    - Avoids rewriting the whole block.
    """
    eff_rpr = _effective_rpr_inner_in_arch(arch_styles_xml_text, style_id).strip()
    if not eff_rpr:
        return style_block

    # If the style has no rPr at all, inject the minimal effective rPr.
    if "<w:rPr" not in style_block:
        return style_block.replace(
            "</w:style>",
            f"\n  <w:rPr>{eff_rpr}</w:rPr>\n</w:style>"
        )

    # Expand self-closing rPr to open/close so we can inject children.
    if re.search(r"<w:rPr\b[^>]*/>", style_block):
        style_block = re.sub(r"<w:rPr\b[^>]*/>", "<w:rPr></w:rPr>", style_block, count=1)

    cur_rpr = _extract_rpr_inner(style_block) or ""

    missing_nodes: List[str] = []

    def _get_child_node(tag: str) -> Optional[str]:
        # self-closing or paired tags, searched within eff_rpr
        m = re.search(rf"(<w:{tag}\b[^>]*/>)", eff_rpr)
        if m:
            return m.group(1)
        m = re.search(rf"(<w:{tag}\b[^>]*>[\s\S]*?</w:{tag}>)", eff_rpr, flags=re.S)
        if m:
            return m.group(1)
        return None

    for tag in ["rFonts", "sz", "szCs", "lang"]:
        if _rpr_contains_tag(cur_rpr, tag):
            continue
        node = _get_child_node(tag)
        if node:
            missing_nodes.append(node)

    if not missing_nodes:
        return style_block

    insertion = "".join(missing_nodes)
    return _inject_missing_rpr_children(style_block, insertion)

def materialize_arch_style_block(style_block: str, style_id: str, arch_styles_xml_text: str) -> str:
    """
    Phase 2: import-time style hardening.

    Goal: ensure styles imported from the architect template remain visually stable
    when applied in a different document, without touching runs or numbering.xml.

    Strategy:
    - Inject pPr only for paragraph styles, and only if missing entirely.
    - Materialize a minimal typography FORCE set into rPr:
        w:rFonts, w:sz, w:szCs, w:lang
      Values are copied from the *effective* architect chain + docDefaults.
    """
    m = re.search(r'<w:style\b[^>]*w:type="([^"]+)"', style_block)
    stype = m.group(1) if m else None

    # Inject pPr only if missing entirely (paragraph styles only)
    if stype == "paragraph" and "<w:pPr" not in style_block:
        effp = _effective_ppr_inner_in_arch(arch_styles_xml_text, style_id)
        if effp.strip():
            style_block = style_block.replace(
                "</w:style>",
                f"\n  <w:pPr>{effp}</w:pPr>\n</w:style>"
            )

    # Typography materialization
    style_block = _materialize_minimal_typography(style_block, style_id, arch_styles_xml_text)

    return style_block

def _collect_style_deps_from_arch(arch_styles_text: str, style_id: str, seen: Set[str]) -> None:
    """
    Recursively collect styleId dependencies via basedOn, link, and next references.
    """
    if style_id in seen:
        return
    seen.add(style_id)

    blk = extract_style_block_raw(arch_styles_text, style_id)
    if not blk:
        return

    for tag in ('basedOn', 'link', 'next'):
        m = re.search(rf'<w:{tag}\b[^>]*w:val="([^"]+)"', blk)
        if m:
            ref = m.group(1)
            if ref and ref not in seen:
                _collect_style_deps_from_arch(arch_styles_text, ref, seen)


def collect_style_dependency_closure(
    arch_styles_text: str,
    style_ids: List[str],
) -> Set[str]:
    """Return requested styles plus all basedOn/link/next dependencies."""
    expanded: Set[str] = set()
    for style_id in style_ids:
        _collect_style_deps_from_arch(arch_styles_text, style_id, expanded)
    return expanded

def extract_style_block_raw(styles_xml_text: str, style_id: str) -> Optional[str]:
    """
    Extract the raw <w:style ...>...</w:style> block for a given styleId using regex.
    This avoids ET rewriting / reformatting.
    """
    # styleId can include characters that need escaping in regex
    sid = re.escape(style_id)
    m = re.search(rf'(<w:style\b[^>]*w:styleId="{sid}"[^>]*>[\s\S]*?</w:style>)', styles_xml_text)
    return m.group(1) + "\n" if m else None

def normalize_style_block_for_compare(style_block: str) -> str:
    return re.sub(r"\s+", " ", style_block).strip()

def style_blocks_equivalent(target_block: str, arch_block: str) -> bool:
    return normalize_style_block_for_compare(target_block) == normalize_style_block_for_compare(arch_block)

def replace_style_block(styles_xml_text: str, style_id: str, new_block: str) -> str:
    sid = re.escape(style_id)
    return re.sub(
        rf'(<w:style\b[^>]*w:styleId="{sid}"[^>]*>[\s\S]*?</w:style>\n?)',
        new_block,
        styles_xml_text,
        count=1,
    )

def import_arch_styles_into_target(
    target_extract_dir: Path,
    arch_styles_xml: str,
    needed_style_ids: List[str],
    log: List[str],
    style_numid_remap: Optional[Dict[str, Dict[str, int]]] = None
) -> None:
    """
    Copy specific style blocks from architect styles.xml into target styles.xml (idempotent),
    including basedOn dependencies.

    arch_styles_xml: synthetic or real styles.xml content as a string
    (built from arch_template_registry.json via build_arch_styles_xml_from_registry).
    """
    tgt_styles_path = target_extract_dir / "word" / "styles.xml"

    arch_styles_text = arch_styles_xml
    tgt_styles_text = read_xml_text(tgt_styles_path)
    original_tgt_styles_text = tgt_styles_text

    # Expand basedOn deps
    expanded = collect_style_dependency_closure(arch_styles_text, needed_style_ids)

    blocks: List[str] = []
    replaced_any = False
    missing: List[str] = []
    for sid in sorted(expanded):
        blk = extract_style_block_raw(arch_styles_text, sid)
        if not blk:
            if sid in WORD_BUILTIN_STYLE_IDS:
                if extract_style_block_raw(tgt_styles_text, sid):
                    continue
                log.append(f"Skipped built-in style dependency (implicit in Word): {sid}")
                continue
            missing.append(sid)
            continue

        # handle numPr: remap if we have mapping, otherwise strip
        if "<w:numPr" in blk:
            if style_numid_remap and sid in style_numid_remap:
                # remap numId to the imported numbering
                remap = style_numid_remap[sid]
                old_num_id = remap["old_numId"]
                new_num_id = remap["new_numId"]
                blk = re.sub(
                    r'(<w:numId\s+w:val=")' + str(old_num_id) + r'"',
                    rf'\g<1>{new_num_id}"',
                    blk
                )
                log.append(f"Remapped numId {old_num_id} -> {new_num_id} in style: {sid}")
            else:
                raise ValueError(
                    f"Style '{sid}' contains <w:numPr> but no numbering remap is "
                    f"available. Numbering import failed or was skipped."
                )

        # HARDEN: make style self-contained (pPr/rPr) to prevent font drift
        blk = materialize_arch_style_block(blk, sid, arch_styles_text)

        existing_blk = extract_style_block_raw(tgt_styles_text, sid)
        if existing_blk:
            if style_blocks_equivalent(existing_blk, blk):
                log.append(f"Style already matches architect: {sid}")
                continue
            tgt_styles_text = replace_style_block(tgt_styles_text, sid, blk)
            replaced_any = True
            log.append(f"Replaced conflicting target style with architect definition: {sid}")
            continue

        blocks.append(blk)
        log.append(f"Imported style from architect: {sid}")

    # Priority-1 hardening: if the architect template is missing any required style or dependency,
    # fail fast rather than emitting a partially formatted output.
    if missing:
        missing_sorted = ", ".join(sorted(set(missing)))
        raise ValueError(
            "Architect styles.xml is missing required styleIds needed for Phase 2 import: "
            f"{missing_sorted}"
        )

    if not blocks and not replaced_any:
        return

    tgt_new = insert_styles_into_styles_xml(tgt_styles_text, blocks)
    if tgt_new != original_tgt_styles_text:
        write_xml_text(tgt_styles_path, tgt_new)

def insert_styles_into_styles_xml(styles_xml_text: str, style_blocks: List[str]) -> str:
    if not style_blocks:
        return styles_xml_text

    # Idempotence: skip inserting styles that already exist in styles.xml
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
