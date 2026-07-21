"""
XML helper functions for paragraph-level DOCX manipulation.

All functions use regex-based XML processing (not DOM/ElementTree)
to preserve byte-level fidelity.
"""

import html
import re
from typing import Dict, Any, Callable, Generator, Iterable, List, Optional, Tuple

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
    if not re.fullmatch(r"[A-Za-z_][\w.-]*:[A-Za-z_][\w.-]*", qualified_name):
        raise ValueError(f"Invalid qualified XML element name: {qualified_name!r}")

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
        name_match = re.match(
            r"<\s*(?P<close>/)?\s*(?P<name>[A-Za-z_][\w.:-]*)(?=\s|/?>)",
            token,
        )
        if name_match is None or name_match.group("name") != qualified_name:
            continue

        is_close = bool(name_match.group("close"))
        is_self_closing = not is_close and bool(re.search(r"/\s*>$", token))

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


def iter_paragraph_xml_blocks(document_xml_text: str) -> Generator[Tuple[int, int, str], None, None]:
    """Yield stable, non-overlapping top-level ``w:p`` blocks."""
    yield from iter_element_xml_blocks(document_xml_text, "w:p")


def _remove_element_blocks(xml_text: str, qualified_names: Iterable[str]) -> str:
    ranges = []
    for name in qualified_names:
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
    p_xml, preserved = _protect_out_of_scope_subtrees(p_xml)
    # If pStyle already exists, replace its value
    if re.search(r"<w:pStyle\b", p_xml):
        p_xml = re.sub(
            r'(<w:pStyle\b[^>]*w:val=")([^"]+)(")',
            rf'\g<1>{styleId}\g<3>',
            p_xml,
            count=1
        )
        return _restore_out_of_scope_subtrees(p_xml, preserved)

    # Handle self-closing pPr: <w:pPr/> or <w:pPr />
    if re.search(r"<w:pPr\b[^>]*/>", p_xml):
        p_xml = re.sub(
            r"<w:pPr\b[^>]*/>",
            rf'<w:pPr><w:pStyle w:val="{styleId}"/></w:pPr>',
            p_xml,
            count=1
        )
        return _restore_out_of_scope_subtrees(p_xml, preserved)

    # If pPr exists as a normal open/close element, insert pStyle right after opening tag
    if "<w:pPr" in p_xml:
        p_xml = re.sub(
            r'(<w:pPr\b[^>]*>)',
            rf'\1<w:pStyle w:val="{styleId}"/>',
            p_xml,
            count=1
        )
        return _restore_out_of_scope_subtrees(p_xml, preserved)

    # No pPr at all: create one right after <w:p ...>
    p_xml = re.sub(
        r'(<w:p\b[^>]*>)',
        rf'\1<w:pPr><w:pStyle w:val="{styleId}"/></w:pPr>',
        p_xml,
        count=1
    )
    return _restore_out_of_scope_subtrees(p_xml, preserved)

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
        result = rpr_text
        for name in sorted(property_names):
            result = re.sub(rf'<w:{name}\b[^>]*/>', '', result)
            result = re.sub(
                rf'<w:{name}\b[^>]*>[\s\S]*?</w:{name}>',
                '',
                result,
                flags=re.S,
            )

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
