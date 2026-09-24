"""
XML helper functions for paragraph-level DOCX manipulation.

All functions use regex-based XML processing (not DOM/ElementTree)
to preserve byte-level fidelity.
"""

import html
import re
from dataclasses import dataclass, field
from types import MappingProxyType
from typing import (
    AbstractSet,
    Any,
    Callable,
    Dict,
    FrozenSet,
    Generator,
    Iterable,
    Iterator,
    List,
    Mapping,
    Optional,
    Set,
    Tuple,
)
from xml.sax.saxutils import escape as xml_escape

_QUALIFIED_NAME_RE = re.compile(r"[A-Za-z_][\w.-]*:[A-Za-z_][\w.-]*")
_TAG_NAME_RE = re.compile(r"<\s*(?P<close>/)?\s*(?P<name>[A-Za-z_][\w.:-]*)(?=\s|/?>)")
_SELF_CLOSING_RE = re.compile(r"/\s*>$")
_NAME_TERMINATORS = frozenset(" \t\r\n/>")


def element_is_mentioned(xml_text: str, qualified_name: str) -> bool:
    """Cheap pre-check: does ``<qualified_name`` occur as a tag start at all?

    Tokenizing a paragraph is the engine's dominant cost, and most paragraphs
    contain no drawing, text box, or revision subtree, so every scanner asks
    this first and skips the tokenizer when the answer is no.
    """

    probe = f"<{qualified_name}"
    cursor = 0
    while True:
        start = xml_text.find(probe, cursor)
        if start < 0:
            return False
        after = start + len(probe)
        if after >= len(xml_text) or xml_text[after] in _NAME_TERMINATORS:
            return True
        cursor = after

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
OUT_OF_SCOPE_SUBTREE_NAMES = (
    "w:drawing",
    "w:pict",
    "w:object",
    "w:txbxContent",
    "v:textbox",
    "wps:txbx",
    # This subtree stores the *previous* paragraph properties for a tracked
    # formatting revision.  Its pStyle/numPr are historical metadata, not the
    # live properties of the paragraph, and must be ignored by readers and
    # preserved byte-for-byte by paragraph edits.
    "w:pPrChange",
    # Historical run properties are likewise metadata, not live direct
    # formatting. They must never be stripped while applying a paragraph style.
    "w:rPrChange",
)


def iter_element_xml_blocks(
    xml_text: str,
    qualified_name: str,
) -> Generator[Tuple[int, int, str], None, None]:
    """Yield non-overlapping outermost blocks for a qualified XML element.

    The scanner deliberately works on raw text so callers can splice modified
    blocks back into the original XML without serializing the document.  Unlike
    a non-greedy regex, it keeps track of nesting and therefore cannot terminate
    an outer paragraph at a text-box paragraph's closing tag.  Both paired and
    self-closing elements are supported.
    """
    if not _QUALIFIED_NAME_RE.fullmatch(qualified_name):
        raise ValueError(f"Invalid qualified XML element name: {qualified_name!r}")
    if not element_is_mentioned(xml_text, qualified_name):
        return

    depth = 0
    outer_start: Optional[int] = None

    cursor = 0
    while True:
        start = xml_text.find("<", cursor)
        if start < 0:
            break
        if xml_text.startswith("<!--", start):
            marker = xml_text.find("-->", start + 4)
            if marker < 0:
                raise ValueError("Malformed XML: unterminated comment")
            cursor = marker + 3
            continue
        if xml_text.startswith("<![CDATA[", start):
            marker = xml_text.find("]]>", start + 9)
            if marker < 0:
                raise ValueError("Malformed XML: unterminated CDATA")
            cursor = marker + 3
            continue
        if xml_text.startswith("<?", start):
            marker = xml_text.find("?>", start + 2)
            if marker < 0:
                raise ValueError("Malformed XML: unterminated processing instruction")
            cursor = marker + 2
            continue

        quote: Optional[str] = None
        tag_end = start + 1
        while tag_end < len(xml_text):
            char = xml_text[tag_end]
            if quote is not None:
                if char == quote:
                    quote = None
            elif char in {'"', "'"}:
                quote = char
            elif char == ">":
                break
            tag_end += 1
        if tag_end >= len(xml_text):
            raise ValueError(f"Malformed XML: unterminated tag at character {start}")

        end = tag_end + 1
        token = xml_text[start:end]
        cursor = end
        name_match = _TAG_NAME_RE.match(token)
        if name_match is None or name_match.group("name") != qualified_name:
            continue

        is_close = bool(name_match.group("close"))
        is_self_closing = not is_close and bool(_SELF_CLOSING_RE.search(token))

        if is_close:
            if depth == 0:
                # Ignore an unmatched close tag here; XML validation happens at
                # package boundaries and this scanner must remain non-mutating.
                continue
            depth -= 1
            if depth == 0 and outer_start is not None:
                yield outer_start, end, xml_text[outer_start:end]
                outer_start = None
            continue

        if depth == 0:
            if is_self_closing:
                yield start, end, token
                continue
            outer_start = start

        if not is_self_closing:
            depth += 1

    if depth != 0 or outer_start is not None:
        raise ValueError(f"Malformed XML: unclosed <{qualified_name}> element")


def _scan_markup(xml_text: str, start: int) -> Tuple[int, str, Optional[str], bool, bool]:
    """Scan the markup that begins with the ``<`` at ``start``.

    Returns ``(end, kind, name, is_close, is_self_closing)``. ``kind`` is
    ``"tag"`` for an element tag and ``"misc"`` for a comment, CDATA section,
    processing instruction or any other ``<`` construct, which carry no
    element name. Quoted attribute values are stepped over whole, so a ``>``
    inside one never ends the tag.
    """

    if xml_text.startswith("<!--", start):
        marker = xml_text.find("-->", start + 4)
        if marker < 0:
            raise ValueError("Malformed XML: unterminated comment")
        return marker + 3, "misc", None, False, False
    if xml_text.startswith("<![CDATA[", start):
        marker = xml_text.find("]]>", start + 9)
        if marker < 0:
            raise ValueError("Malformed XML: unterminated CDATA")
        return marker + 3, "misc", None, False, False
    if xml_text.startswith("<?", start):
        marker = xml_text.find("?>", start + 2)
        if marker < 0:
            raise ValueError("Malformed XML: unterminated processing instruction")
        return marker + 2, "misc", None, False, False

    quote: Optional[str] = None
    tag_end = start + 1
    while tag_end < len(xml_text):
        char = xml_text[tag_end]
        if quote is not None:
            if char == quote:
                quote = None
        elif char in {'"', "'"}:
            quote = char
        elif char == ">":
            break
        tag_end += 1
    if tag_end >= len(xml_text):
        raise ValueError(f"Malformed XML: unterminated tag at character {start}")

    end = tag_end + 1
    token = xml_text[start:end]
    name_match = _TAG_NAME_RE.match(token)
    if name_match is None:
        return end, "misc", None, False, False
    is_close = bool(name_match.group("close"))
    is_self_closing = not is_close and bool(_SELF_CLOSING_RE.search(token))
    return end, "tag", name_match.group("name"), is_close, is_self_closing


def iter_direct_child_xml_blocks(
    element_xml_text: str,
) -> Generator[Tuple[int, int, str, str], None, None]:
    """Yield direct child element blocks without serializing the XML.

    Each item is ``(start, end, qualified_name, raw_block)``.  Nested elements,
    comments, CDATA, and processing instructions are scanned structurally so a
    matching descendant cannot be mistaken for a direct property child.
    """

    cursor = 0
    root_name: Optional[str] = None
    while root_name is None:
        start = element_xml_text.find("<", cursor)
        if start < 0:
            raise ValueError("Malformed XML: element block has no opening tag")
        end, kind, name, is_close, is_self_closing = _scan_markup(element_xml_text, start)
        cursor = end
        if kind == "misc":
            continue
        if is_close or name is None:
            raise ValueError("Malformed XML: element block starts with a close tag")
        root_name = name
        if is_self_closing:
            return

    depth = 0
    child_start: Optional[int] = None
    child_name: Optional[str] = None
    while True:
        start = element_xml_text.find("<", cursor)
        if start < 0:
            break
        end, kind, name, is_close, is_self_closing = _scan_markup(element_xml_text, start)
        cursor = end
        if kind == "misc" or name is None:
            continue
        if is_close:
            if depth == 0:
                if name != root_name:
                    raise ValueError(
                        f"Malformed XML: expected </{root_name}>, found </{name}>"
                    )
                return
            depth -= 1
            if depth == 0 and child_start is not None and child_name is not None:
                yield (
                    child_start,
                    end,
                    child_name,
                    element_xml_text[child_start:end],
                )
                child_start = None
                child_name = None
            continue

        if depth == 0:
            if is_self_closing:
                yield start, end, name, element_xml_text[start:end]
                continue
            child_start = start
            child_name = name
        if not is_self_closing:
            depth += 1

    raise ValueError(f"Malformed XML: unclosed <{root_name}> element")


def iter_paragraph_xml_blocks(document_xml_text: str) -> Generator[Tuple[int, int, str], None, None]:
    """Yield stable, non-overlapping top-level ``w:p`` blocks."""
    yield from iter_element_xml_blocks(document_xml_text, "w:p")


def _remove_element_blocks(xml_text: str, qualified_names: Iterable[str]) -> str:
    ranges = []
    for name in qualified_names:
        if not element_is_mentioned(xml_text, name):
            continue
        ranges.extend((start, end) for start, end, _block in iter_element_xml_blocks(xml_text, name))
    if not ranges:
        return xml_text

    # Elements in this helper can be nested (for example a drawing containing a
    # text box).  Merge overlapping ranges before slicing.
    merged = []
    for start, end in sorted(ranges):
        if merged and start <= merged[-1][1]:
            merged[-1] = (merged[-1][0], max(merged[-1][1], end))
        else:
            merged.append((start, end))

    pieces = []
    last = 0
    for start, end in merged:
        pieces.append(xml_text[last:start])
        last = end
    pieces.append(xml_text[last:])
    return "".join(pieces)


def strip_out_of_scope_subtrees(xml_text: str) -> str:
    """Remove non-live/protected subtrees for host-paragraph analysis only."""
    return _remove_element_blocks(xml_text, OUT_OF_SCOPE_SUBTREE_NAMES)


def _protect_out_of_scope_subtrees(xml_text: str) -> Tuple[str, List[Tuple[str, str]]]:
    ranges = []
    for name in OUT_OF_SCOPE_SUBTREE_NAMES:
        if not element_is_mentioned(xml_text, name):
            continue
        ranges.extend(
            (start, end)
            for start, end, _block in iter_element_xml_blocks(xml_text, name)
        )
    merged: List[Tuple[int, int]] = []
    for start, end in sorted(ranges):
        if merged and start <= merged[-1][1]:
            merged[-1] = (merged[-1][0], max(merged[-1][1], end))
        else:
            merged.append((start, end))
    if not merged:
        return xml_text, []

    nonce = 0
    while f"__PHASE2_PRESERVED_{nonce}_" in xml_text:
        nonce += 1
    replacements: List[Tuple[str, str]] = []
    pieces: List[str] = []
    last = 0
    for index, (start, end) in enumerate(merged):
        token = f"__PHASE2_PRESERVED_{nonce}_{index}__"
        pieces.extend((xml_text[last:start], token))
        replacements.append((token, xml_text[start:end]))
        last = end
    pieces.append(xml_text[last:])
    return "".join(pieces), replacements


def _restore_out_of_scope_subtrees(
    xml_text: str,
    replacements: List[Tuple[str, str]],
) -> str:
    restored = xml_text
    for token, subtree in replacements:
        if restored.count(token) != 1:
            raise ValueError("Out-of-scope subtree placeholder was altered during paragraph edit")
        restored = restored.replace(token, subtree, 1)
    return restored


def edit_preserving_out_of_scope_subtrees(
    xml_text: str,
    edit: Callable[[str], str],
) -> str:
    protected, replacements = _protect_out_of_scope_subtrees(xml_text)
    return _restore_out_of_scope_subtrees(edit(protected), replacements)

def paragraph_text_from_block(p_xml: str) -> str:
    # Deleted/moved-from text and field instructions are not visible paragraph
    # content.  Drawing/text-box subtrees are intentionally outside Phase 2 and
    # must not be conflated with their host paragraph's text.
    visible = strip_out_of_scope_subtrees(p_xml)
    visible = _remove_element_blocks(
        visible,
        ("w:del", "w:moveFrom", "w:instrText"),
    )

    # Tabs and explicit breaks separate words in Word despite not being w:t
    # nodes.  Preserve the normalizer's visible-text semantics exactly.
    separator_token = "\ue000"
    nonbreaking_hyphen_token = "\ue001"
    visible = re.sub(
        r"<w:(tab|br|cr)\b[^>]*>\s*</w:\1\s*>",
        separator_token,
        visible,
        flags=re.S,
    )
    visible = re.sub(r"<w:(?:tab|br|cr)\b[^>]*/>", separator_token, visible)
    visible = re.sub(
        r"<w:noBreakHyphen\b[^>]*>\s*</w:noBreakHyphen\s*>",
        nonbreaking_hyphen_token,
        visible,
        flags=re.S,
    )
    visible = re.sub(r"<w:noBreakHyphen\b[^>]*/>", nonbreaking_hyphen_token, visible)
    visible = re.sub(
        r"<w:softHyphen\b[^>]*>\s*</w:softHyphen\s*>",
        "",
        visible,
        flags=re.S,
    )
    visible = re.sub(r"<w:softHyphen\b[^>]*/>", "", visible)
    pieces = re.findall(
        rf"<w:t\b[^>]*>([\s\S]*?)</w:t>|({separator_token})|({nonbreaking_hyphen_token})",
        visible,
    )
    if not pieces:
        return ""
    joined = html.unescape(
        "".join(
            text
            if text
            else (
                " "
                if separator
                else ("\u2011" if nonbreaking_hyphen else "")
            )
            for text, separator, nonbreaking_hyphen in pieces
        )
    )
    joined = re.sub(r"\s+", " ", joined).strip()
    return joined

def paragraph_contains_sectpr(p_xml: str) -> bool:
    live_xml = strip_out_of_scope_subtrees(p_xml)
    return next(iter_element_xml_blocks(live_xml, "w:sectPr"), None) is not None

def paragraph_pstyle_from_block(p_xml: str) -> Optional[str]:
    live_xml = strip_out_of_scope_subtrees(p_xml)
    m = re.search(r"<w:pStyle\b[^>]*w:val=\"([^\"]+)\"", live_xml)
    return m.group(1) if m else None

def paragraph_numpr_from_block(p_xml: str) -> Dict[str, Optional[str]]:
    numId = None
    ilvl = None
    live_xml = strip_out_of_scope_subtrees(p_xml)
    m1 = re.search(r"<w:numId\b[^>]*w:val=\"([^\"]+)\"", live_xml)
    m2 = re.search(r"<w:ilvl\b[^>]*w:val=\"([^\"]+)\"", live_xml)
    if m1: numId = m1.group(1)
    if m2: ilvl = m2.group(1)
    return {"numId": numId, "ilvl": ilvl}

def paragraph_ppr_hints_from_block(p_xml: str) -> Dict[str, Any]:
    # lightweight hints (alignment + ind + spacing)
    p_xml = strip_out_of_scope_subtrees(p_xml)
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
    for k in ["before", "after", "line"]:
        m3 = re.search(rf"<w:spacing\b[^>]*w:{k}=\"([^\"]+)\"", p_xml)
        if m3:
            spacing[k] = m3.group(1)
    if spacing:
        hints["spacing"] = spacing
    return hints

def apply_pstyle_to_paragraph_block(p_xml: str, styleId: str) -> str:
    """Return the paragraph block with ``styleId`` as its live style.

    ``styleId`` is the style ID itself, not attribute text: it is escaped
    here, so passing text that is already escaped would escape it twice.

    The element is always written the way Word writes it,
    ``<w:pStyle w:val="..."/>``, because every pStyle reader in this engine
    matches that form. An existing live ``w:pStyle`` is therefore replaced
    whole rather than having its value substituted. A substitution that
    expected ``w:val="..."`` matched nothing for ``w:val='Old'`` or
    ``w:val=""`` and silently returned the paragraph with its old style.

    Everything is spliced by position, so the style ID is never read as a
    regex replacement template. Out-of-scope subtrees (text boxes, drawings,
    tracked property changes) are protected, so their own ``w:pStyle`` is
    never taken for the live one and they come back byte-for-byte.
    """

    if not isinstance(styleId, str) or not styleId:
        raise ValueError(f"Paragraph style ID must be a non-empty string: {styleId!r}")
    # &, < and " cannot appear literally in a double-quoted attribute. > is
    # escaped as well because the engine's tag patterns stop at the first >.
    value = xml_escape(styleId, {'"': "&quot;"})
    pstyle = f'<w:pStyle w:val="{value}"/>'

    p_xml, preserved = _protect_out_of_scope_subtrees(p_xml)
    ppr = next(iter_element_xml_blocks(p_xml, "w:pPr"), None)
    if ppr is None:
        # No pPr at all: create one as the paragraph's first child.
        opening = re.match(r"<w:p\b[^>]*>", p_xml)
        if opening is None:
            raise ValueError("Paragraph block does not start with a <w:p> element")
        tag = opening.group(0)
        closing = ""
        if _SELF_CLOSING_RE.search(tag):
            # An empty <w:p/> must be opened up, or the pPr lands outside it.
            tag = _SELF_CLOSING_RE.sub(">", tag)
            closing = "</w:p>"
        edited = f"{tag}<w:pPr>{pstyle}</w:pPr>{closing}{p_xml[opening.end():]}"
        return _restore_out_of_scope_subtrees(edited, preserved)

    start, end, block = ppr
    if _SELF_CLOSING_RE.search(block):
        block = f"<w:pPr>{pstyle}</w:pPr>"
    else:
        live = next(iter_element_xml_blocks(block, "w:pStyle"), None)
        if live is not None:
            block = block[:live[0]] + pstyle + block[live[1]:]
        else:
            # pStyle is the first child in CT_PPr's sequence.
            opening_end = re.match(r"<w:pPr\b[^>]*>", block).end()
            block = block[:opening_end] + pstyle + block[opening_end:]
    return _restore_out_of_scope_subtrees(p_xml[:start] + block + p_xml[end:], preserved)

def strip_direct_run_properties(
    p_xml: str,
    properties: Iterable[str],
) -> str:
    """Remove selected direct ``w:rPr`` children from visible runs only.

    Paragraph-style application may remove a direct run property only when
    the effective replacement style supplies that same property.  Callers
    therefore pass the exact local OOXML names resolved from the architect
    style (for example ``rFonts``, ``sz``, ``szCs``, or ``lang``).
    """

    property_names = set(properties)
    for name in property_names:
        if not re.fullmatch(r"[A-Za-z_][\w.-]*", name):
            raise ValueError(f"Invalid direct run property name: {name!r}")
    if not property_names:
        return p_xml

    p_xml, preserved = _protect_out_of_scope_subtrees(p_xml)

    def strip_properties_from_rpr(rpr_text: str) -> str:
        ranges = [
            (start, end)
            for start, end, qualified_name, _block in iter_direct_child_xml_blocks(
                rpr_text
            )
            if qualified_name.startswith("w:")
            and qualified_name[2:] in property_names
        ]
        if ranges:
            pieces: List[str] = []
            last = 0
            for start, end in ranges:
                pieces.append(rpr_text[last:start])
                last = end
            pieces.append(rpr_text[last:])
            result = "".join(pieces)
        else:
            result = rpr_text

        inner = re.sub(
            r'<w:rPr\b[^>]*>([\s\S]*)</w:rPr>',
            r'\1',
            result,
            flags=re.S,
        )
        return '' if not inner.strip() else result

    def process_run(run_match: re.Match[str]) -> str:
        return re.sub(
            r'<w:rPr\b[^>]*>[\s\S]*?</w:rPr>',
            lambda match: strip_properties_from_rpr(match.group(0)),
            run_match.group(0),
            count=1,
            flags=re.S,
        )

    result = re.sub(
        r'<w:r\b[^>]*>[\s\S]*?</w:r>',
        process_run,
        p_xml,
        flags=re.S,
    )
    return _restore_out_of_scope_subtrees(result, preserved)


def strip_run_font_formatting(p_xml: str) -> str:
    """
    Strip font-related formatting from all runs in a paragraph.

    This allows the paragraph style's font definitions to take effect,
    overriding hardcoded run-level fonts (common in MasterSpec/ARCOM docs).

    Strips from <w:rPr> inside <w:r>:
    - <w:rFonts .../> (font family)
    - <w:sz .../> (font size)
    - <w:szCs .../> (complex script font size)

    Preserves:
    - Bold, italic, underline, strikethrough
    - Colors, highlighting
    - Character styles (<w:rStyle>)
    - Everything else
    """
    return strip_direct_run_properties(p_xml, {"rFonts", "sz", "szCs"})

_DIRECT_PPR_OVERRIDE_TAGS = ("jc", "ind", "spacing", "numPr")

def strip_conflicting_direct_ppr(
    p_xml: str,
    *,
    preserve_numpr: bool = False,
) -> str:
    """
    Remove direct paragraph-layout overrides that commonly win over paragraph styles.

    Strips these tags from paragraph-level <w:pPr> only:
    - <w:jc>
    - <w:ind>
    - <w:spacing>
    - <w:numPr>, unless ``preserve_numpr`` is true

    Preserves section properties and other unrelated pPr children.
    """
    p_xml, preserved = _protect_out_of_scope_subtrees(p_xml)

    def _strip_from_ppr(match):
        ppr = match.group(0)
        tags = (
            tuple(tag for tag in _DIRECT_PPR_OVERRIDE_TAGS if tag != "numPr")
            if preserve_numpr
            else _DIRECT_PPR_OVERRIDE_TAGS
        )
        for tag in tags:
            ppr = re.sub(rf'<w:{tag}\b[^>]*/>', '', ppr)
            ppr = re.sub(rf'<w:{tag}\b[^>]*>[\s\S]*?</w:{tag}>', '', ppr, flags=re.S)
        return ppr

    result = re.sub(r'<w:pPr\b[^>]*>[\s\S]*?</w:pPr>', _strip_from_ppr, p_xml, count=1, flags=re.S)
    return _restore_out_of_scope_subtrees(result, preserved)


# ─────────────────────────────────────────────────────────────────────────────
# Namespace declarations for fragments that move between parts
# ─────────────────────────────────────────────────────────────────────────────
#
# A fragment copied out of one part keeps its prefixes but not the
# declarations that gave them meaning: those sit on the source part's root.
# Written into another part it means whatever that part's root says, or
# nothing at all (``unbound prefix``). Every current Word template puts
# ``w14:ligatures`` in its document defaults, so this is the ordinary case,
# not an exotic one.
#
# These helpers stay lexical like the rest of this module: the destination is
# edited only in its root opening tag, and fragments are never re-serialized.
# A prefix counts wherever the markup gives it meaning: in an element or
# attribute name, and in the values of the markup-compatibility attributes that
# list prefixes or qualified names (``Requires`` on ``Choice``, ``Ignorable``,
# ``MustUnderstand``, ``ProcessContent``, ``PreserveElements``,
# ``PreserveAttributes``). A prefix named only in such a value is still one an
# MCE consumer must resolve; left undeclared, it silently takes the fallback or
# drops the feature.

#: Markup Compatibility and Extensibility (ECMA-376 Part 3).
MCE_NS = "http://schemas.openxmlformats.org/markup-compatibility/2006"

# ``xml`` is bound by definition and ``xmlns`` is not a prefix at all.
_RESERVED_PREFIXES = frozenset({"xml", "xmlns"})

# Markup-compatibility attributes (ECMA-376 Part 3) whose values name
# prefixes: lists of prefixes, and lists of qualified names whose prefixes
# count. ``Requires`` is the unqualified attribute of ``mc:Choice``.
_MCE_PREFIX_LIST_ATTRIBUTES = frozenset({"Ignorable", "MustUnderstand"})
_MCE_QNAME_LIST_ATTRIBUTES = frozenset({"ProcessContent", "PreserveElements", "PreserveAttributes"})
_NO_BINDINGS: Mapping[str, str] = MappingProxyType({})

# One attribute of a start tag: its name and its quoted value. Values are
# matched whole, so a quote or ">" inside another value cannot end one early.
_TAG_ATTRIBUTE_RE = re.compile(
    r"""\s(?P<name>[^\s=/>"']+)\s*=\s*(?P<quote>["'])(?P<value>.*?)(?P=quote)""",
    re.S,
)
_XML_REFERENCE_RE = re.compile(r"&(?:#x([0-9A-Fa-f]+)|#([0-9]+)|(lt|gt|amp|quot|apos));")
_PREDEFINED_ENTITIES = {"lt": "<", "gt": ">", "amp": "&", "quot": '"', "apos": "'"}
_TAG_CLOSE_RE = re.compile(r"\s*/?\s*>\Z")


class NamespaceReconciliationError(ValueError):
    """A fragment's prefix cannot keep its meaning in the part it moves into.

    ``reason`` is ``"conflict"`` when the destination root already binds the
    prefix to a different namespace, and ``"undeclared"`` when the source
    root does not say what the prefix means. ``prefix`` is ``""`` for the
    default namespace.
    """

    def __init__(self, prefix: str, reason: str, message: str) -> None:
        super().__init__(message)
        self.prefix = prefix
        self.reason = reason


def xml_unescape(text: str) -> str:
    """Expand XML's five predefined entities and numeric character references.

    Deliberately not :func:`html.unescape`, which also expands HTML entities
    such as ``&nbsp;`` that are not XML and cannot occur in well-formed
    WordprocessingML.
    """

    def expand(match: "re.Match[str]") -> str:
        hexadecimal, decimal, name = match.groups()
        if hexadecimal is not None:
            return chr(int(hexadecimal, 16))
        if decimal is not None:
            return chr(int(decimal))
        return _PREDEFINED_ENTITIES[name]

    return _XML_REFERENCE_RE.sub(expand, text)


def _attribute_value(raw: str) -> str:
    """An attribute's value as a parser reports it (XML 1.0, section 3.3.3)."""

    normalized = raw.replace("\r\n", "\n").replace("\r", "\n")
    normalized = normalized.replace("\t", " ").replace("\n", " ")
    return xml_unescape(normalized)


def _double_quoted(value: str) -> str:
    """``value`` escaped to stand between double quotes in an attribute."""

    return xml_escape(value, {'"': "&quot;"})


def _tag_attributes(tag: str) -> Iterator["re.Match[str]"]:
    name = _TAG_NAME_RE.match(tag)
    return _TAG_ATTRIBUTE_RE.finditer(tag, name.end() if name is not None else 0)


def _declared_prefix(attribute_name: str) -> Optional[str]:
    """The prefix an ``xmlns`` attribute declares (``""`` for the default)."""

    if attribute_name == "xmlns":
        return ""
    if attribute_name.startswith("xmlns:"):
        return attribute_name[len("xmlns:"):]
    return None


def _tag_declarations(tag: str) -> Dict[str, str]:
    declarations: Dict[str, str] = {}
    for attribute in _tag_attributes(tag):
        prefix = _declared_prefix(attribute.group("name"))
        if prefix is not None:
            declarations[prefix] = _attribute_value(attribute.group("value"))
    return declarations


def _mce_ignorable_attribute(
    tag: str,
    declarations: Mapping[str, str],
) -> Optional["re.Match[str]"]:
    """The tag's markup-compatibility ``Ignorable`` attribute, found by namespace.

    Word 2007 bound markup compatibility to ``ve`` rather than ``mc``, so the
    literal prefix proves nothing; an ``mc:Ignorable`` whose ``mc`` means
    something else is not the attribute at all.
    """

    for attribute in _tag_attributes(tag):
        prefix, _, local = attribute.group("name").partition(":")
        if local == "Ignorable" and declarations.get(prefix) == MCE_NS:
            return attribute
    return None


def _root_open_tag(xml_text: str) -> Tuple[int, int]:
    """Locate the document element's opening tag, skipping the prolog."""

    cursor = 0
    while True:
        start = xml_text.find("<", cursor)
        if start < 0:
            raise ValueError("XML part has no root element")
        end, kind, name, is_close, _is_self_closing = _scan_markup(xml_text, start)
        cursor = end
        if kind == "misc":
            if xml_text.startswith("<!", start) and not xml_text.startswith("<!--", start):
                # A document type declaration: never legitimate in a package
                # part, and rejected at every parse boundary.
                raise ValueError("XML part declares a document type before its root element")
            continue
        if is_close or name is None:
            raise ValueError("XML part starts with a closing tag")
        return start, end


def root_opening_tag(xml_text: str) -> str:
    """The document element's opening tag, exactly as written."""

    start, end = _root_open_tag(xml_text)
    return xml_text[start:end]


def root_namespace_declarations(xml_text: str) -> Dict[str, str]:
    """Prefix -> namespace URI for every declaration on the root opening tag.

    The default namespace is keyed by ``""``. Declarations on descendants are
    not included: only the root's are in scope for a child written directly
    beneath it.
    """

    start, end = _root_open_tag(xml_text)
    return _tag_declarations(xml_text[start:end])


def root_ignorable_prefixes(xml_text: str) -> FrozenSet[str]:
    """The prefixes the root lists in its markup-compatibility ``Ignorable``."""

    start, end = _root_open_tag(xml_text)
    tag = xml_text[start:end]
    attribute = _mce_ignorable_attribute(tag, _tag_declarations(tag))
    if attribute is None:
        return frozenset()
    return frozenset(_attribute_value(attribute.group("value")).split())


@dataclass(frozen=True)
class RootNamespaces:
    """The declarations and ignorable prefixes on one part's root element.

    Built once from the architect's stylesheet and handed to every step that
    writes an architect fragment into the target's, so each of them resolves
    a prefix to the same namespace.
    """

    declarations: Mapping[str, str]
    ignorable: FrozenSet[str] = field(default_factory=frozenset)

    def __post_init__(self) -> None:
        object.__setattr__(self, "declarations", MappingProxyType(dict(self.declarations)))
        object.__setattr__(self, "ignorable", frozenset(self.ignorable))

    @classmethod
    def of(cls, xml_text: str) -> "RootNamespaces":
        return cls(root_namespace_declarations(xml_text), root_ignorable_prefixes(xml_text))


def prefixes_used(fragment: str, context: Mapping[str, str] = _NO_BINDINGS) -> Set[str]:
    """The prefixes a fragment needs from whatever encloses it.

    Element and attribute names are read; ``xml`` and ``xmlns`` never need a
    declaration. A prefix the fragment declares itself is not reported
    within that declaration's scope, and is reported again if it is used
    outside it. An unprefixed element needs the default namespace, reported
    as ``""``; an unprefixed attribute is in no namespace and needs nothing.
    Comments, CDATA sections and processing instructions are text.

    Prefixes named in markup-compatibility values count as well (see the
    section comment above). Those attributes are recognised by namespace, not
    by the literal ``mc`` -- Word 2007 used ``ve`` -- so ``context``, the
    declarations in scope around the fragment, is how one is recognised when
    the fragment does not declare the markup-compatibility prefix itself.
    ``context`` only identifies them: whatever the fragment does not declare
    is reported, whether or not ``context`` declares it.
    """

    needed: Set[str] = set()
    scopes: List[Dict[str, str]] = []

    def in_scope(prefix: str) -> bool:
        return any(prefix in scope for scope in scopes)

    def resolve(prefix: str) -> Optional[str]:
        for scope in reversed(scopes):
            if prefix in scope:
                return scope[prefix]
        return context.get(prefix)

    def need(prefix: str) -> None:
        if prefix not in _RESERVED_PREFIXES and not in_scope(prefix):
            needed.add(prefix)

    cursor = 0
    while True:
        start = fragment.find("<", cursor)
        if start < 0:
            return needed
        end, kind, name, is_close, is_self_closing = _scan_markup(fragment, start)
        cursor = end
        if kind != "tag" or name is None:
            continue
        if is_close:
            if scopes:
                scopes.pop()
            continue
        attributes = [
            (match.group("name"), match.group("value"))
            for match in _tag_attributes(fragment[start:end])
        ]
        declared_here: Dict[str, str] = {}
        for attribute, value in attributes:
            prefix = _declared_prefix(attribute)
            if prefix is not None:
                declared_here[prefix] = _attribute_value(value)
        scopes.append(declared_here)

        element_prefix, _, element_local = name.rpartition(":")
        need(element_prefix)
        is_mce_choice = element_local == "Choice" and resolve(element_prefix) == MCE_NS
        for attribute, value in attributes:
            if _declared_prefix(attribute) is not None:
                continue
            prefix, _, local = attribute.rpartition(":")
            if prefix:
                need(prefix)
            named: List[str] = []
            if prefix and resolve(prefix) == MCE_NS:
                if local in _MCE_PREFIX_LIST_ATTRIBUTES:
                    named = _attribute_value(value).split()
                elif local in _MCE_QNAME_LIST_ATTRIBUTES:
                    named = [
                        token.partition(":")[0]
                        for token in _attribute_value(value).split()
                        if ":" in token
                    ]
            elif not prefix and local == "Requires" and is_mce_choice:
                named = _attribute_value(value).split()
            for named_prefix in named:
                if named_prefix:
                    need(named_prefix)
        if is_self_closing:
            scopes.pop()


def _conflict(prefix: str, existing: str, wanted: str) -> NamespaceReconciliationError:
    label = "the default namespace" if prefix == "" else f"prefix {prefix!r}"
    return NamespaceReconciliationError(
        prefix,
        "conflict",
        f"{label} is bound to {existing!r} in the destination root, "
        f"but the inserted markup needs {wanted!r}",
    )


def ensure_root_declarations(
    part_xml: str,
    needed: Mapping[str, str],
    ignorable: AbstractSet[str] = frozenset(),
) -> str:
    """Declare ``needed`` prefixes on the part's root and mark ``ignorable`` ones.

    ``needed`` maps each prefix inserted markup uses to the namespace it must
    mean. A prefix the root already binds to that namespace needs nothing; a
    missing one gains an ``xmlns:`` declaration; one the root binds to a
    different namespace raises :class:`NamespaceReconciliationError`, because
    writing the markup would silently change what it says. The default
    namespace (``""``) is compared but never added or changed: that would
    change the meaning of every unprefixed name already in the part.

    Each prefix in ``ignorable`` must be declared once ``needed`` is applied,
    and is added to the root's markup-compatibility ``Ignorable`` list. An
    existing list keeps its prefix, quoting and tokens; without one, the root
    gains ``mc:Ignorable`` and, when no prefix is bound to the
    markup-compatibility namespace yet, ``xmlns:mc``.

    Only the root opening tag changes, and a part that needs nothing is
    returned unchanged.
    """

    start, end = _root_open_tag(part_xml)
    tag = part_xml[start:end]
    declared = _tag_declarations(tag)
    additions: List[Tuple[str, str]] = []
    for prefix in sorted(needed):
        wanted = needed[prefix]
        if prefix in _RESERVED_PREFIXES:
            continue
        if prefix == "":
            existing = declared.get("", "")
            if existing != wanted:
                raise _conflict("", existing, wanted)
            continue
        existing = declared.get(prefix)
        if existing is None:
            additions.append((prefix, wanted))
            declared[prefix] = wanted
        elif existing != wanted:
            raise _conflict(prefix, existing, wanted)

    tokens = sorted(prefix for prefix in ignorable if prefix)
    for prefix in tokens:
        if prefix not in declared:
            raise ValueError(
                f"cannot list prefix {prefix!r} as ignorable: the root does not declare it"
            )

    edited = tag
    new_attributes = "".join(
        f' xmlns:{prefix}="{_double_quoted(uri)}"' for prefix, uri in additions
    )
    if tokens:
        attribute = _mce_ignorable_attribute(tag, declared)
        if attribute is not None:
            raw = attribute.group("value")
            listed = _attribute_value(raw).split()
            missing = [prefix for prefix in tokens if prefix not in listed]
            if missing:
                value = (raw.rstrip() + " " if raw.strip() else "") + " ".join(missing)
                edited = (
                    tag[: attribute.start("value")] + value + tag[attribute.end("value"):]
                )
        else:
            mce_prefixes = [prefix for prefix, uri in declared.items() if prefix and uri == MCE_NS]
            if "mc" in mce_prefixes or not mce_prefixes:
                mce_prefix = "mc"
            else:
                mce_prefix = mce_prefixes[0]
            if mce_prefix not in declared:
                new_attributes += f' xmlns:mc="{MCE_NS}"'
                declared["mc"] = MCE_NS
            elif declared[mce_prefix] != MCE_NS:
                raise _conflict(mce_prefix, declared[mce_prefix], MCE_NS)
            new_attributes += f' {mce_prefix}:Ignorable="{" ".join(tokens)}"'

    if new_attributes:
        close = _TAG_CLOSE_RE.search(edited)
        if close is None:  # pragma: no cover - _root_open_tag returns whole tags
            raise ValueError("XML part root tag is not terminated")
        edited = edited[: close.start()] + new_attributes + edited[close.start():]
    if edited == tag:
        return part_xml
    return part_xml[:start] + edited + part_xml[end:]


def root_namespace_additions(before_xml: str, after_xml: str) -> Tuple[str, ...]:
    """Prefixes the root of ``after_xml`` declares or lists as ignorable anew.

    For reporting what :func:`ensure_root_declarations` changed: each entry is
    a prefix, ``mc:Ignorable=<prefix>`` for an ignorable token. Prefixes are
    markup identifiers, never document text.
    """

    declared = sorted(
        set(root_namespace_declarations(after_xml)) - set(root_namespace_declarations(before_xml))
    )
    ignorable = sorted(root_ignorable_prefixes(after_xml) - root_ignorable_prefixes(before_xml))
    return tuple(declared) + tuple(f"mc:Ignorable={prefix}" for prefix in ignorable)


def declare_fragment_namespaces(
    part_xml: str,
    fragments: Iterable[str],
    source: RootNamespaces,
) -> str:
    """Declare in ``part_xml`` every prefix ``fragments`` use, meaning what it meant in ``source``.

    ``source`` is the root the fragments were taken from. A prefix it does
    not declare raises :class:`NamespaceReconciliationError` with reason
    ``"undeclared"``: its meaning at the fragment's original position came
    from a declaration somewhere below the source root, which a fragment
    moved out of that position no longer has. Prefixes ``source`` lists as
    ignorable stay ignorable in ``part_xml``.
    """

    used: Set[str] = set()
    for fragment in fragments:
        used |= prefixes_used(fragment, source.declarations)
    needed: Dict[str, str] = {}
    for prefix in sorted(used):
        if prefix == "":
            needed[""] = source.declarations.get("", "")
        elif prefix in source.declarations:
            needed[prefix] = source.declarations[prefix]
        else:
            raise NamespaceReconciliationError(
                prefix,
                "undeclared",
                f"prefix {prefix!r} is used by the inserted markup but its "
                "source root does not declare it",
            )
    return ensure_root_declarations(part_xml, needed, source.ignorable & used)

