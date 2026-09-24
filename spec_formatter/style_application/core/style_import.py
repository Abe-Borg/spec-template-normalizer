"""
Style extraction, materialization, and import for Phase 2.

Handles importing architect styles into target documents with
property materialization for cross-document portability.
"""

import hashlib
import functools
import re
from dataclasses import dataclass
from pathlib import Path
from typing import Callable, Dict, FrozenSet, Iterator, List, Set, Optional, Tuple

from .errors import EngineError
from .ooxml_text import read_xml_text, write_xml_text
from .untrusted_xml import UntrustedXmlError, parse_untrusted_xml
from .xml_helpers import (
    NamespaceReconciliationError,
    RootNamespaces,
    declare_fragment_namespaces,
    edit_preserving_out_of_scope_subtrees,
    iter_direct_child_xml_blocks,
    root_namespace_additions,
    root_opening_tag,
)

# Word built-in styles that exist implicitly in every DOCX.
# These never need to be imported — Word creates them internally
# even when styles.xml has no explicit <w:style> block for them.
WORD_BUILTIN_STYLE_IDS = frozenset({
    "Normal",
    "DefaultParagraphFont",
    "TableNormal",
    "NoList",
})

@dataclass(frozen=True)
class StyleImportResult:
    """Collision-safe style IDs selected for an imported architect graph."""

    # Shell/header/footer consumers use the complete dependency graph.
    style_id_map: Dict[str, str]
    # Format-only body roles use fully materialized, numbering-detached clones.
    body_style_id_map: Dict[str, str]
    imported_style_ids: Set[str]
    # What the import added to the target stylesheet's root so the imported
    # blocks keep their meaning: prefixes, and ``mc:Ignorable=<prefix>``.
    declared_namespace_prefixes: Tuple[str, ...] = ()


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


# Elements whose ``w:val`` names a style: content refers to one from a
# paragraph, run, or table, and a style definition refers to another style.
CONTENT_STYLE_REFERENCES = ("pStyle", "rStyle", "tblStyle")
STYLE_DEFINITION_REFERENCES = ("basedOn", "link", "next")

# One attribute in either quoting, with optional whitespace around "=". Each
# value is matched whole, so a quote, a ">" or a w:val lookalike inside some
# other attribute's value is never taken for the end of a tag or a reference.
_ATTRIBUTE = r"""\s+[^\s=/>"']+\s*=\s*(?:"[^"]*"|'[^']*')"""
_ATTRIBUTE_OTHER_THAN_VAL = r"""\s+(?!w:val\s*=)[^\s=/>"']+\s*=\s*(?:"[^"]*"|'[^']*')"""

# Comments, CDATA sections and processing instructions hold text, not markup,
# the same three regions ``iter_element_xml_blocks`` steps over. They are
# matched only so that a reference-shaped string inside one is skipped whole.
# An unterminated one runs to the end, so nothing after its start counts.
# Written without their opening "<", which the pattern below matches once.
_NON_MARKUP = r"!--.*?(?:-->|\Z)|!\[CDATA\[.*?(?:\]\]>|\Z)|\?.*?(?:\?>|\Z)"


@functools.lru_cache(maxsize=8)
def _style_reference_re(elements: Tuple[str, ...]) -> "re.Pattern[str]":
    """A non-markup region, or the opening tag of one of ``elements``
    with its ``w:val`` captured.

    Every alternative opens with "<", matched once ahead of them. A pattern
    that starts with a literal lets the regex engine jump from one "<" to
    the next instead of trying each alternative at every character, several
    times faster on a style block. The ``basedOn`` walks run it on every hop.
    """

    names = "|".join(re.escape(name) for name in elements)
    return re.compile(
        rf"<(?:(?P<skip>{_NON_MARKUP})"
        rf"|w:(?:{names})(?=[\s/>])(?:{_ATTRIBUTE_OTHER_THAN_VAL})*"
        r"""\s+(?P<val>w:val\s*=\s*(?:"(?P<double>[^"]*)"|'(?P<single>[^']*)'))"""
        rf"(?:{_ATTRIBUTE})*\s*/?>)",
        re.S,
    )


def _iter_style_references(
    xml_text: str,
    elements: Tuple[str, ...],
) -> Iterator[Tuple[int, int, str]]:
    """Yield ``(start, end, value)`` for each reference in ``elements``.

    ``start`` and ``end`` span the ``w:val`` attribute from its name to its
    closing quote; ``value`` is the raw attribute text between the quotes.
    Text inside a comment, CDATA section or processing instruction is never
    a reference, even when it is shaped like one.
    """

    for match in _style_reference_re(tuple(elements)).finditer(xml_text):
        if match.group("skip") is not None:
            continue
        value = match.group("double")
        if value is None:
            value = match.group("single")
        yield match.start("val"), match.end("val"), value


def referenced_style_ids(xml_text: str, elements: Tuple[str, ...]) -> List[str]:
    """Every style ID named by a reference in ``elements``, in document order.

    This is the reader paired with :func:`remap_style_references`. Whatever
    decides which architect styles are imported must see exactly the
    references that are later pointed at the clones. One it misses gets no
    clone and keeps its architect ID, which then resolves silently to the
    target's own style of that name, or fails publication if there is none.
    An empty ``w:val`` names no style and is skipped.
    """

    return [value for _start, _end, value in _iter_style_references(xml_text, elements) if value]


def remap_style_references(
    xml_text: str,
    elements: Tuple[str, ...],
    style_id_map: Dict[str, str],
) -> str:
    """Point each style reference in ``elements`` at its mapped style ID.

    A reference is found whatever its quoting and whatever whitespace
    surrounds its ``=``. Its value is looked up as raw attribute text, the
    domain in which the map's keys were collected and matched against the
    architect's ``w:styleId`` attributes. Nothing is decoded, so a
    double-quoted reference maps exactly as it always has.

    A remapped reference is written back as ``w:val="..."``, the form Word
    writes and the only one some of the engine's regex readers match, the
    paragraph ``pStyle`` readers among them. A reference that is empty, has
    no mapping, or maps to itself is left exactly as written. Edits are
    spliced by position, so a style ID is never read as a regex replacement
    template.
    """

    pieces: List[str] = []
    cursor = 0
    for start, end, value in _iter_style_references(xml_text, elements):
        if not value:
            continue
        destination = style_id_map.get(value, value)
        if destination == value:
            continue
        if (
            not isinstance(destination, str)
            or not destination
            or '"' in destination
            or "<" in destination
        ):
            raise ValueError(
                f"Cannot point style reference {value!r} at {destination!r}: "
                "a destination must be a non-empty style ID that can stand "
                "between double quotes"
            )
        pieces.append(xml_text[cursor:start])
        pieces.append(f'w:val="{destination}"')
        cursor = end
    if not pieces:
        return xml_text
    pieces.append(xml_text[cursor:])
    return "".join(pieces)


def _rewrite_style_id_and_references(
    style_block: str,
    source_style_id: str,
    destination_style_id: str,
    style_id_map: Dict[str, str],
) -> str:
    """Give a cloned style block its new ID and point its references at clones.

    ``basedOn``, ``link`` and ``next`` go through ``remap_style_references``,
    the rewriter paired with the reader that built the dependency closure, so
    a single-quoted reference is remapped like a double-quoted one. The
    block's own ``w:styleId`` is matched only in its double-quoted form, the
    one ``extract_style_block_raw`` found the block by.
    """

    own_id = re.search(rf'w:styleId="{re.escape(source_style_id)}"', style_block)
    if own_id is not None:
        style_block = (
            style_block[:own_id.start()]
            + f'w:styleId="{destination_style_id}"'
            + style_block[own_id.end():]
        )
    return remap_style_references(style_block, STYLE_DEFINITION_REFERENCES, style_id_map)


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

_STYLE_ID_ATTR_RE = re.compile(r'<w:style\b[^>]*?\sw:styleId="([^"]+)"')


@functools.lru_cache(maxsize=16)
def _style_block_index(styles_xml_text: str) -> Dict[str, str]:
    """Every ``w:style`` block keyed by styleId, built in one structural pass.

    Cached on the styles text itself: a formatting run reads the same
    ``styles.xml`` string thousands of times (once per paragraph per basedOn
    hop), and re-scanning it with a regex each time was the engine's
    second-largest cost. The scan is structure-aware, so a self-closing
    ``<w:style .../>`` cannot swallow the block that follows it. The first
    occurrence of a duplicated ID wins, matching the old ``re.search``.
    """

    from .xml_helpers import iter_element_xml_blocks

    index: Dict[str, str] = {}
    for _start, _end, block in iter_element_xml_blocks(styles_xml_text, "w:style"):
        match = _STYLE_ID_ATTR_RE.match(block)
        if match is not None:
            index.setdefault(match.group(1), block)
    return index


def _extract_style_block(styles_xml_text: str, style_id: str) -> Optional[str]:
    return _style_block_index(styles_xml_text).get(style_id)

def _extract_basedOn(style_block: str) -> Optional[str]:
    """The style ``style_block`` is based on, as raw attribute text.

    Every ``basedOn`` walk takes its next hop from here, so it reads with the
    grammar of ``referenced_style_ids``: the one the dependency closure
    follows and the clone rewrite remaps. A walk that stopped at a
    single-quoted reference lost the parent without failing anything. A
    detached Format-only clone dropped the parent's formatting, a shell clone
    pinned the document defaults over it, a target paragraph's inherited list
    was invisible to the style swap that preserves it and to the invariant
    that checks it, and the hierarchy converters left that paragraph behind
    when they renumbered its list. The first non-empty reference wins.
    """

    # Most blocks, and the root style that ends most walks, have no basedOn.
    # The grammar cannot match without this literal, so skip the scan.
    if "<w:basedOn" not in style_block:
        return None
    for _start, _end, value in _iter_style_references(style_block, ("basedOn",)):
        if value:
            return value
    return None

def _extract_numpr_block(style_block: str) -> Optional[str]:
    m = re.search(r'(<w:numPr\b[^>]*>[\s\S]*?</w:numPr>)', style_block, flags=re.S)
    return m.group(1) if m else None

@functools.lru_cache(maxsize=8192)
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


# Children a style does not pass on through ``basedOn``: its own identity
# (pStyle/rStyle), numbering (the mode policy owns it), section properties,
# and tracked-change history. Matched by qualified name, like every child.
_PPR_CHILDREN_NOT_INHERITED = frozenset({"w:pStyle", "w:numPr", "w:sectPr", "w:pPrChange"})
_RPR_CHILDREN_NOT_INHERITED = frozenset({"w:rStyle", "w:rPrChange"})


def _property_children(
    inner_xml: str,
    container: str,
    excluded: FrozenSet[str],
    owner: str,
) -> List[Tuple[str, str]]:
    """Direct children of a property container as ``(qualified name, source bytes)``.

    Read lexically, never parsed. A property block is a fragment whose
    prefixes are declared on the architect stylesheet's root, not on the
    block, so a parser given only the ``w`` namespace fails on the first
    extension child -- ``w14:ligatures`` in the document defaults of every
    template current Word saves -- and re-serializing would reformat the
    children besides. ``owner`` names the style in the error for a fragment
    that cannot be read; the fragment itself is never quoted.
    """

    if not inner_xml.strip():
        return []
    try:
        children = list(iter_direct_child_xml_blocks(f"<{container}>{inner_xml}</{container}>"))
    except ValueError as exc:
        raise ValueError(f"{owner} contains malformed {container} XML") from exc
    return [(name, block) for _start, _end, name, block in children if name not in excluded]


def _ppr_children_by_name(inner_xml: str, owner: str) -> List[Tuple[str, str]]:
    """Direct pPr children, nested borders and tabs left whole."""

    return _property_children(inner_xml, "w:pPr", _PPR_CHILDREN_NOT_INHERITED, owner)


def _rpr_children_by_name(inner_xml: str, owner: str) -> List[Tuple[str, str]]:
    """Direct rPr children for property-wise inheritance."""

    return _property_children(inner_xml, "w:rPr", _RPR_CHILDREN_NOT_INHERITED, owner)


def _core_children_first(order: List[str]) -> List[str]:
    """``w:`` children in first-seen order, then every extension child.

    Word writes extension-namespace children (``w14:``, ``w15:``) after the
    ``w:`` ones, and the strict schema does not know them at all.
    """

    return [name for name in order if name.startswith("w:")] + [
        name for name in order if not name.startswith("w:")
    ]


def _effective_properties_in_arch(
    arch_styles_xml_text: str,
    style_id: str,
    container: str,
    read_children: Callable[[str, str], List[Tuple[str, str]]],
    defaults_inner: str,
) -> str:
    """Resolve each property child independently through ``basedOn``, then the defaults.

    Keyed by qualified name: ``w:shadow`` and ``w14:shadow`` are different
    properties that share a local name, and keying by local name let the
    nearer one hide the other. Extension children are carried, not dropped:
    ``w14:textFill`` is visible, and dropping it would silently change what
    the architect specified.
    """

    resolved: Dict[str, str] = {}
    order: List[str] = []
    seen: Set[str] = set()
    cur = style_id
    while cur and cur not in seen and len(seen) < 50:
        seen.add(cur)
        block = _extract_style_block(arch_styles_xml_text, cur)
        if not block:
            break
        inner = _extract_tag_inner(block, container) or ""
        for name, node in read_children(inner, f"Architect style {cur!r}"):
            if name not in resolved:
                resolved[name] = node
                order.append(name)
        cur = _extract_basedOn(block)

    for name, node in read_children(defaults_inner, "Architect document defaults"):
        if name not in resolved:
            resolved[name] = node
            order.append(name)
    return "".join(resolved[name] for name in _core_children_first(order))


def _effective_ppr_inner_in_arch(arch_styles_xml_text: str, style_id: str) -> str:
    """Resolve each formatting property independently through basedOn."""

    return _effective_properties_in_arch(
        arch_styles_xml_text,
        style_id,
        "w:pPr",
        _ppr_children_by_name,
        _docdefaults_ppr_inner(arch_styles_xml_text),
    )


def _effective_full_rpr_inner_in_arch(
    arch_styles_xml_text: str,
    style_id: str,
) -> str:
    """Resolve every reusable run property independently through basedOn."""

    return _effective_properties_in_arch(
        arch_styles_xml_text,
        style_id,
        "w:rPr",
        _rpr_children_by_name,
        _docdefaults_rpr_inner(arch_styles_xml_text),
    )


def _replace_first(pattern: str, replacement: str, text: str) -> str:
    """Replace the first match of ``pattern`` with ``replacement``, literally.

    Materialized fragments are source bytes. ``re.sub`` would read a
    backslash in one as a group reference in the replacement template.
    """

    return re.sub(pattern, lambda _match: replacement, text, count=1)


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
        return _replace_first(
            r"<w:rPr\b[^>]*/>",
            f"<w:rPr>{effective}</w:rPr>" if effective else "",
            style_block,
        )
    if re.search(r"<w:rPr\b[^>]*>[\s\S]*?</w:rPr>", style_block):
        return _replace_first(
            r"<w:rPr\b[^>]*>[\s\S]*?</w:rPr>",
            f"<w:rPr>{effective}</w:rPr>" if effective else "",
            style_block,
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

def _rpr_opening_before(style_block: str, close: int) -> Optional[int]:
    """Where the ``w:rPr`` element that ``</w:rPr>`` at ``close`` ends begins."""

    position = close
    while True:
        position = style_block.rfind("<w:rPr", 0, position)
        if position < 0:
            return None
        # <w:rPrChange> shares the spelling; only <w:rPr> itself qualifies.
        if style_block[position + len("<w:rPr"):position + len("<w:rPr") + 1] in (
            " ", "\t", "\r", "\n", ">", "/",
        ):
            return position


def _inject_missing_rpr_children(style_block: str, missing_children_xml: str) -> str:
    """Insert missing ``w:`` rPr children (raw XML) into the first rPr.

    They go ahead of the rPr's first extension-namespace child, where Word
    keeps its own ``w:`` children, or just before ``</w:rPr>`` when it has
    none. Only the first closing tag is used, as before.
    """
    if not missing_children_xml.strip():
        return style_block
    close = style_block.find("</w:rPr>")
    if close < 0:
        return style_block
    insert_at = close
    opening = _rpr_opening_before(style_block, close)
    if opening is not None:
        rpr_block = style_block[opening:close + len("</w:rPr>")]
        try:
            for start, _end, name, _block in iter_direct_child_xml_blocks(rpr_block):
                if not name.startswith("w:"):
                    insert_at = opening + start
                    break
        except ValueError:
            # A malformed rPr keeps the old insertion point. The imported
            # stylesheet is parsed before it is written, and that names the
            # style.
            pass
    return style_block[:insert_at] + missing_children_xml + style_block[insert_at:]

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
                style_block = _replace_first(
                    r"<w:pPr\b[^>]*/>",
                    f"<w:pPr>{combined_ppr}</w:pPr>",
                    style_block,
                )
            elif re.search(r"<w:pPr\b[^>]*>[\s\S]*?</w:pPr>", style_block):
                style_block = _replace_first(
                    r"<w:pPr\b[^>]*>[\s\S]*?</w:pPr>",
                    f"<w:pPr>{combined_ppr}</w:pPr>",
                    style_block,
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

    References are read with ``referenced_style_ids``, the grammar the clones
    are later rewritten with, so a reference in either quoting is followed.
    """
    if style_id in seen:
        return
    seen.add(style_id)

    blk = extract_style_block_raw(arch_styles_text, style_id)
    if not blk:
        return

    for ref in referenced_style_ids(blk, STYLE_DEFINITION_REFERENCES):
        if ref not in seen:
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
    architect_namespaces: Optional[RootNamespaces] = None,
) -> StyleImportResult:
    """
    Copy specific style blocks from architect styles.xml into target styles.xml (idempotent),
    including basedOn dependencies.

    arch_styles_xml: the architect styles as a string -- the bundle's
    ``portable_styles.xml``. The architect-free modes never import styles.

    architect_namespaces: the declarations on ``arch_styles_xml``'s root. The
    shared application path builds it once per target and passes the same
    value to docDefaults application; when omitted it is read from
    ``arch_styles_xml`` here, which gives the same answer.

    Every prefix the imported blocks use is declared on the target
    stylesheet's root with the meaning it had in the architect's, and keeps
    its ``mc:Ignorable`` status. A prefix the target already binds to a
    different namespace, or one the architect root does not declare, fails
    with ``style_import_namespace_conflict`` before anything is written. The
    stylesheet is parsed before it is written; a block that does not parse
    fails naming its style ID.
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

    try:
        tgt_new = declare_fragment_namespaces(
            tgt_styles_text,
            blocks,
            architect_namespaces or RootNamespaces.of(arch_styles_text),
        )
    except NamespaceReconciliationError as exc:
        raise EngineError(
            "style_import_namespace_conflict",
            f"Architect styles cannot be written into the target styles.xml: {exc}",
        ) from exc
    declared = root_namespace_additions(tgt_styles_text, tgt_new)
    if declared:
        log.append(
            "Added XML namespace declarations to the target styles.xml root "
            "for imported architect styles: "
            + ", ".join(declared)
        )
    tgt_new = insert_styles_into_styles_xml(tgt_new, blocks)
    _require_well_formed_styles(
        tgt_new,
        [(sid, blk) for sid, final_id, blk in prepared_blocks if final_id in imported_ids],
    )
    if tgt_new != original_tgt_styles_text:
        write_xml_text(tgt_styles_path, tgt_new)
    return StyleImportResult(
        style_id_map=style_id_map,
        body_style_id_map=body_style_id_map,
        imported_style_ids=imported_ids,
        declared_namespace_prefixes=declared,
    )

def _require_well_formed_styles(
    styles_xml_text: str,
    imported_blocks: List[Tuple[str, str]],
) -> None:
    """Refuse to write a stylesheet that does not parse, naming the culprit.

    Materialization reads fragments lexically and never parses them, so this
    is where a malformed architect fragment is caught -- before the target is
    touched, not at packaging. When the whole part fails, each imported block
    is parsed on its own under the part's root declarations, and the first
    that fails is named by its source style ID. The XML is never quoted.
    """

    try:
        parse_untrusted_xml(styles_xml_text, "word/styles.xml")
    except UntrustedXmlError as exc:
        # The blocks were inserted before </w:styles>, so the root is w:styles.
        opening = root_opening_tag(styles_xml_text)
        for style_id, block in imported_blocks:
            try:
                parse_untrusted_xml(f"{opening}{block}</w:styles>", "word/styles.xml")
            except UntrustedXmlError:
                raise ValueError(
                    f"Architect style {style_id!r} is not well-formed XML; "
                    "it was not imported"
                ) from None
        raise ValueError(
            "The target styles.xml is not well-formed after architect styles were "
            "imported, and every imported style parses on its own"
        ) from exc


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
