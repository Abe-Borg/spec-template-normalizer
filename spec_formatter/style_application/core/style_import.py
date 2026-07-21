"""
Style extraction, materialization, and import for Phase 2.

Handles importing architect styles into target documents with
property materialization for cross-document portability.
"""

import hashlib
import re
import xml.etree.ElementTree as ET
from dataclasses import dataclass
from pathlib import Path
from typing import Dict, List, Set, Optional

from .ooxml_text import read_xml_text, write_xml_text
from .xml_helpers import edit_preserving_out_of_scope_subtrees

# Word built-in styles that exist implicitly in every DOCX.
# These never need to be imported — Word creates them internally
# even when styles.xml has no explicit <w:style> block for them.
WORD_BUILTIN_STYLE_IDS = frozenset({
    "Normal",
    "DefaultParagraphFont",
    "TableNormal",
    "NoList",
})

_W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
ET.register_namespace("w", _W_NS)


@dataclass(frozen=True)
class StyleImportResult:
    """Collision-safe style IDs selected for an imported architect graph."""

    # Shell/header/footer consumers use the complete dependency graph.
    style_id_map: Dict[str, str]
    # Format-only body roles use fully materialized, numbering-detached clones.
    body_style_id_map: Dict[str, str]
    imported_style_ids: Set[str]


def _namespaced_style_id(
    seed: str,
    style_id: str,
    style_block: str,
    *,
    variant: str,
) -> str:
    """Return a deterministic app-owned Word style ID."""

    safe = re.sub(r"[^A-Za-z0-9_-]+", "_", style_id).strip("_")[:28] or "Style"
    safe_variant = re.sub(r"[^A-Za-z0-9_-]+", "_", variant).strip("_")[:12] or "STYLE"
    graph_hash = hashlib.sha256(seed.encode("utf-8")).hexdigest()[:8]
    style_hash = hashlib.sha256(style_block.encode("utf-8")).hexdigest()[:8]
    return f"SF_{graph_hash}_{safe_variant}_{safe}_{style_hash}"


def _rewrite_style_id_and_references(
    style_block: str,
    source_style_id: str,
    destination_style_id: str,
    style_id_map: Dict[str, str],
) -> str:
    out = re.sub(
        rf'(w:styleId="){re.escape(source_style_id)}(")',
        rf"\g<1>{destination_style_id}\2",
        style_block,
        count=1,
    )
    for tag in ("basedOn", "link", "next"):
        out = re.sub(
            rf'(<w:{tag}\b[^>]*w:val=")([^"]+)(")',
            lambda match: (
                match.group(1)
                + style_id_map.get(match.group(2), match.group(2))
                + match.group(3)
            ),
            out,
        )
    return out


def _make_format_only_body_style_self_contained(style_block: str) -> str:
    """Remove every path by which an architect list can reach a body style."""

    out = re.sub(r"<w:numPr\b[^>]*/>", "", style_block)
    out = re.sub(r"<w:numPr\b[^>]*>[\s\S]*?</w:numPr>", "", out, flags=re.S)
    # Rendering properties are materialized before this helper is called.  The
    # dependency graph must then be detached so inherited architect numbering
    # cannot leak back into a Format-only paragraph.
    for tag in ("basedOn", "link", "next"):
        out = re.sub(rf"<w:{tag}\b[^>]*/>", "", out)
        out = re.sub(rf"<w:{tag}\b[^>]*>[\s\S]*?</w:{tag}>", "", out, flags=re.S)
    return out

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
    def _materialize(live_xml: str) -> str:
        # pPrChange is protected by edit_preserving_out_of_scope_subtrees, so
        # neither its historical pStyle nor its historical numPr can satisfy
        # these live-property checks.
        if "<w:numPr" in live_xml or "<w:sectPr" in live_xml:
            return live_xml
        pstyle_match = re.search(
            r'<w:pStyle\b[^>]*w:val="([^"]+)"',
            live_xml,
        )
        if not pstyle_match:
            return live_xml
        numpr = _find_style_numpr_in_chain(
            current_styles_xml,
            pstyle_match.group(1),
        )
        if not numpr:
            return live_xml
        if re.search(r"<w:pPr\b[^>]*>", live_xml):
            return re.sub(
                r"(<w:pPr\b[^>]*>)",
                rf"\1{numpr}",
                live_xml,
                count=1,
            )
        return re.sub(
            r"(<w:p\b[^>]*>)",
            rf"\1<w:pPr>{numpr}</w:pPr>",
            live_xml,
            count=1,
        )

    return edit_preserving_out_of_scope_subtrees(paragraph_xml, _materialize)

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

def _ppr_children_by_name(inner_xml: str) -> List[tuple[str, str]]:
    """Parse direct pPr children without flattening nested borders/tabs."""

    if not inner_xml.strip():
        return []
    try:
        root = ET.fromstring(f'<w:pPr xmlns:w="{_W_NS}">{inner_xml}</w:pPr>')
    except ET.ParseError as exc:
        raise ValueError(f"Architect style contains invalid pPr XML: {exc}") from exc
    children: List[tuple[str, str]] = []
    for child in list(root):
        local_name = child.tag.rsplit("}", 1)[-1]
        if local_name in {"pStyle", "numPr", "sectPr", "pPrChange"}:
            continue
        serialized = ET.tostring(child, encoding="unicode", short_empty_elements=True)
        serialized = serialized.replace(f' xmlns:w="{_W_NS}"', "")
        children.append((local_name, serialized))
    return children


def _effective_ppr_inner_in_arch(arch_styles_xml_text: str, style_id: str) -> str:
    """Resolve each formatting property independently through basedOn."""

    resolved: Dict[str, str] = {}
    order: List[str] = []
    seen: Set[str] = set()
    cur = style_id
    while cur and cur not in seen and len(seen) < 50:
        seen.add(cur)
        block = _extract_style_block(arch_styles_xml_text, cur)
        if not block:
            break
        inner = _extract_tag_inner(block, "w:pPr") or ""
        for name, node in _ppr_children_by_name(inner):
            if name not in resolved:
                resolved[name] = node
                order.append(name)
        cur = _extract_basedOn(block)

    for name, node in _ppr_children_by_name(_docdefaults_ppr_inner(arch_styles_xml_text)):
        if name not in resolved:
            resolved[name] = node
            order.append(name)
    return "".join(resolved[name] for name in order)


def _rpr_children_by_name(inner_xml: str) -> List[tuple[str, str]]:
    """Parse direct style rPr children for property-wise inheritance."""

    if not inner_xml.strip():
        return []
    try:
        root = ET.fromstring(f'<w:rPr xmlns:w="{_W_NS}">{inner_xml}</w:rPr>')
    except ET.ParseError as exc:
        raise ValueError(f"Architect style contains invalid rPr XML: {exc}") from exc
    children: List[tuple[str, str]] = []
    for child in list(root):
        local_name = child.tag.rsplit("}", 1)[-1]
        if local_name in {"rStyle", "rPrChange"}:
            continue
        serialized = ET.tostring(child, encoding="unicode", short_empty_elements=True)
        serialized = serialized.replace(f' xmlns:w="{_W_NS}"', "")
        children.append((local_name, serialized))
    return children


def _effective_full_rpr_inner_in_arch(
    arch_styles_xml_text: str,
    style_id: str,
) -> str:
    """Resolve every reusable run property independently through basedOn."""

    resolved: Dict[str, str] = {}
    order: List[str] = []
    seen: Set[str] = set()
    cur = style_id
    while cur and cur not in seen and len(seen) < 50:
        seen.add(cur)
        block = _extract_style_block(arch_styles_xml_text, cur)
        if not block:
            break
        inner = _extract_tag_inner(block, "w:rPr") or ""
        for name, node in _rpr_children_by_name(inner):
            if name not in resolved:
                resolved[name] = node
                order.append(name)
        cur = _extract_basedOn(block)

    for name, node in _rpr_children_by_name(_docdefaults_rpr_inner(arch_styles_xml_text)):
        if name not in resolved:
            resolved[name] = node
            order.append(name)
    return "".join(resolved[name] for name in order)


def _materialize_full_rpr_for_detached_body(
    style_block: str,
    style_id: str,
    arch_styles_xml_text: str,
) -> str:
    """Make a body clone independent of its architect run-style chain."""

    effective = _effective_full_rpr_inner_in_arch(
        arch_styles_xml_text,
        style_id,
    )
    if re.search(r"<w:rPr\b[^>]*/>", style_block):
        return re.sub(
            r"<w:rPr\b[^>]*/>",
            f"<w:rPr>{effective}</w:rPr>" if effective else "",
            style_block,
            count=1,
        )
    if re.search(r"<w:rPr\b[^>]*>[\s\S]*?</w:rPr>", style_block):
        return re.sub(
            r"<w:rPr\b[^>]*>[\s\S]*?</w:rPr>",
            f"<w:rPr>{effective}</w:rPr>" if effective else "",
            style_block,
            count=1,
        )
    if not effective:
        return style_block
    return style_block.replace(
        "</w:style>",
        f"\n  <w:rPr>{effective}</w:rPr>\n</w:style>",
    )

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

    # Replace pPr with the full effective formatting contract, resolving each
    # property through basedOn. Numbering is deliberately excluded and is
    # handled by the mode policy. Preserve a numPr declared directly on this
    # style so Canadian mode can remap it; Format-only removes it afterward.
    if stype == "paragraph":
        direct_numpr = _extract_numpr_block(style_block) or ""
        effp = _effective_ppr_inner_in_arch(arch_styles_xml_text, style_id)
        combined_ppr = direct_numpr + effp
        if combined_ppr.strip():
            if re.search(r"<w:pPr\b[^>]*/>", style_block):
                style_block = re.sub(
                    r"<w:pPr\b[^>]*/>",
                    f"<w:pPr>{combined_ppr}</w:pPr>",
                    style_block,
                    count=1,
                )
            elif re.search(r"<w:pPr\b[^>]*>[\s\S]*?</w:pPr>", style_block):
                style_block = re.sub(
                    r"<w:pPr\b[^>]*>[\s\S]*?</w:pPr>",
                    f"<w:pPr>{combined_ppr}</w:pPr>",
                    style_block,
                    count=1,
                )
            else:
                style_block = style_block.replace(
                    "</w:style>",
                    f"\n  <w:pPr>{combined_ppr}</w:pPr>\n</w:style>",
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
    style_numid_remap: Optional[Dict[str, Dict[str, int]]] = None,
    *,
    format_only_body_style_ids: Optional[Set[str]] = None,
    shell_style_ids: Optional[Set[str]] = None,
    namespace_seed: Optional[str] = None,
) -> StyleImportResult:
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

    body_roots = set(format_only_body_style_ids or set())
    requested = set(needed_style_ids)

    # Format-only role roots are fully materialized and detached from the
    # architect dependency graph. Header/footer and Canadian styles retain a
    # complete graph. A source style used by both gets two distinct clones so
    # body numbering detachment cannot alter the shell style.
    shell_roots = set(shell_style_ids or set()) | (requested - body_roots)
    shell_expanded = collect_style_dependency_closure(
        arch_styles_text,
        sorted(shell_roots),
    )
    expanded = set(shell_expanded) | body_roots

    missing: List[str] = []
    source_blocks: Dict[str, str] = {}
    for sid in sorted(expanded):
        blk = extract_style_block_raw(arch_styles_text, sid)
        if blk:
            source_blocks[sid] = blk
            continue
        if sid in WORD_BUILTIN_STYLE_IDS:
            log.append(f"Skipped built-in style dependency (implicit in Word): {sid}")
            continue
        missing.append(sid)

    # Priority-1 hardening: if the architect template is missing any required
    # style or dependency, fail before modifying styles.xml.
    if missing:
        missing_sorted = ", ".join(sorted(set(missing)))
        raise ValueError(
            "Architect styles.xml is missing required styleIds needed for Phase 2 import: "
            f"{missing_sorted}"
        )

    seed = namespace_seed or hashlib.sha256(arch_styles_text.encode("utf-8")).hexdigest()
    shell_style_id_map: Dict[str, str] = {
        sid: _namespaced_style_id(
            seed,
            sid,
            source_blocks[sid],
            variant="SHELL",
        )
        for sid in sorted(shell_expanded)
        if sid in source_blocks
    }
    detached_body_style_id_map: Dict[str, str] = {
        sid: _namespaced_style_id(
            seed,
            sid,
            source_blocks[sid],
            variant="BODY",
        )
        for sid in sorted(body_roots)
        if sid in source_blocks
    }

    # ``style_id_map`` remains the mapping used by imported shell parts. For
    # body-only imports it also exposes the sole detached clone for backwards
    # compatibility. Application code uses ``body_style_id_map`` explicitly.
    style_id_map = dict(shell_style_id_map)
    for sid, final_id in detached_body_style_id_map.items():
        style_id_map.setdefault(sid, final_id)
    body_style_id_map = {
        sid: detached_body_style_id_map.get(sid, shell_style_id_map.get(sid, sid))
        for sid in requested
    }

    blocks: List[str] = []
    imported_ids: Set[str] = set()
    prepared_blocks: List[tuple[str, str, str]] = []

    for sid in sorted(shell_style_id_map):
        blk = materialize_arch_style_block(source_blocks[sid], sid, arch_styles_text)
        if "<w:numPr" in blk:
            if style_numid_remap and sid in style_numid_remap:
                remap = style_numid_remap[sid]
                old_num_id = remap["old_numId"]
                new_num_id = remap["new_numId"]
                blk = re.sub(
                    r'(<w:numId\s+w:val=")' + str(old_num_id) + r'"',
                    rf'\g<1>{new_num_id}"',
                    blk,
                )
                log.append(f"Remapped numId {old_num_id} -> {new_num_id} in style: {sid}")
            else:
                raise ValueError(
                    f"Style '{sid}' contains <w:numPr> but no numbering remap is "
                    "available. Numbering import failed or was skipped."
                )
        final_id = shell_style_id_map[sid]
        blk = _rewrite_style_id_and_references(
            blk,
            sid,
            final_id,
            shell_style_id_map,
        )
        prepared_blocks.append((sid, final_id, blk))

    for sid in sorted(detached_body_style_id_map):
        blk = materialize_arch_style_block(source_blocks[sid], sid, arch_styles_text)
        blk = _materialize_full_rpr_for_detached_body(
            blk,
            sid,
            arch_styles_text,
        )
        blk = _make_format_only_body_style_self_contained(blk)
        final_id = detached_body_style_id_map[sid]
        blk = _rewrite_style_id_and_references(blk, sid, final_id, {})
        prepared_blocks.append((sid, final_id, blk))

    for sid, final_id, blk in prepared_blocks:
        existing_final = extract_style_block_raw(tgt_styles_text, final_id)
        if existing_final:
            if not style_blocks_equivalent(existing_final, blk):
                raise ValueError(
                    "Deterministic architect style namespace collision for "
                    f"{sid!r} -> {final_id!r}"
                )
            log.append(f"Namespaced architect style already matches: {final_id}")
            continue

        blocks.append(blk)
        imported_ids.add(final_id)
        if extract_style_block_raw(tgt_styles_text, sid):
            log.append(
                f"Imported architect style {sid} as {final_id}; "
                f"preserved target style {sid}"
            )
        else:
            log.append(f"Imported architect style {sid} as app-owned {final_id}")

    if not blocks:
        return StyleImportResult(
            style_id_map=style_id_map,
            body_style_id_map=body_style_id_map,
            imported_style_ids=imported_ids,
        )

    tgt_new = insert_styles_into_styles_xml(tgt_styles_text, blocks)
    if tgt_new != original_tgt_styles_text:
        write_xml_text(tgt_styles_path, tgt_new)
    return StyleImportResult(
        style_id_map=style_id_map,
        body_style_id_map=body_style_id_map,
        imported_style_ids=imported_ids,
    )

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
