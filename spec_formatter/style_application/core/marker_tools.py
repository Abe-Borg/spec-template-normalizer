"""Leading-marker machinery shared by both hierarchy converters.

``csi_to_canadian`` and ``canadian_to_csi`` are inverses, and they run over the
same OOXML problem from opposite ends: find a paragraph's leading numbering
marker, decide what it counts as, and edit it without disturbing anything else
in the paragraph. That work -- mapping visible text back to individual
``w:t`` nodes, refusing to touch a marker tangled in a tracked change or a
field, removing a marker together with its delimiter tab, and proving
afterwards that only the marker changed -- is identical in both directions and
lives here so neither converter carries its own copy of it.

The role marker tables are deliberately shared too. A Canadian ``.1`` and a CSI
``1.`` are both matched by the same per-role alternatives, because the tables
describe *what a leading marker for this role can look like*, which does not
depend on which way a document is being converted.

Nothing here decides policy. Whether a marker should be removed, what should
replace it, and when a sequence is unprovable are the converters' business.
"""

from __future__ import annotations

import bisect
import html
import re
import xml.etree.ElementTree as ET
from dataclasses import dataclass
from typing import Any, Callable, Dict, Optional, Tuple

from spec_formatter.numbering_roles import (
    role_from_numbering_catalog,
    role_from_numbering_signature,
)
from spec_formatter.role_contract import (
    NUMBERED_BODY_ROLES,
    ROLE_FALLBACKS,
    ROLE_LEVEL,
    ROLE_PARENT,
)

from .errors import EngineError
from .section_numbers import section_number_display_form
from .xml_helpers import (
    OUT_OF_SCOPE_SUBTREE_NAMES,
    edit_preserving_out_of_scope_subtrees,
    iter_element_xml_blocks,
    paragraph_text_from_block,
)


NUMBERED_ROLES = NUMBERED_BODY_ROLES

_ROLE_LEVEL = ROLE_LEVEL

#: Roles the locator counts when numbering headings after a SECTION line.
_LOCATOR_HEADING_ROLES = frozenset(NUMBERED_ROLES) | {"PART"}


@dataclass(frozen=True)
class _LiteralMarker:
    marker: str
    family: str
    body_text: str


@dataclass(frozen=True)
class _SourceEvidence:
    paragraph_index: int
    role: str
    source_kind: str
    literal: Optional[_LiteralMarker]
    automatic_numpr: Optional[Dict[str, str]]
    automatic_pattern: Optional[Dict[str, str]]


_ROLE_MARKERS = {
    "PART": re.compile(
        r"^\s*(?P<marker>PART\s+(?:\d+|[IVXLCDM]+))(?=\s|[-\u2010-\u2015:]|$)",
        re.IGNORECASE,
    ),
    "ARTICLE": re.compile(r"^\s*(?P<marker>\d{1,2}\.\d{1,3})(?=\s|$)"),
    "PARAGRAPH": re.compile(r"^\s*(?P<marker>[A-Z]\.|\.\d+)(?=\s|$)"),
    "SUBPARAGRAPH": re.compile(r"^\s*(?P<marker>\d+\.|\.\d+)(?=\s|$)"),
    "SUBSUBPARAGRAPH": re.compile(r"^\s*(?P<marker>[a-z]\.|\.\d+)(?=\s|$)"),
    "SUBPARAGRAPH_LEVEL_5": re.compile(
        r"^\s*(?P<marker>\d+\)|\.\d+)(?=\s|$)"
    ),
    "SUBPARAGRAPH_LEVEL_6": re.compile(
        r"^\s*(?P<marker>[a-z]\)|\.\d+)(?=\s|$)"
    ),
    "SUBPARAGRAPH_LEVEL_7": re.compile(
        r"^\s*(?P<marker>\(\d+\)|\.\d+)(?=\s|$)"
    ),
    "SUBPARAGRAPH_LEVEL_8": re.compile(
        r"^\s*(?P<marker>\([a-z]\)|\.\d+)(?=\s|$)"
    ),
}
# These variants omit delimiter lookaheads because a visible delimiter can be
# a sibling ``w:tab`` element and therefore absent from the joined ``w:t``
# text used to map the marker back to individual runs.
_RAW_ROLE_MARKERS = {
    "PART": re.compile(
        r"^\s*(?P<marker>PART\s+(?:\d+|[IVXLCDM]+))",
        re.IGNORECASE,
    ),
    "ARTICLE": re.compile(r"^\s*(?P<marker>\d{1,2}\.\d{1,3})"),
    "PARAGRAPH": re.compile(r"^\s*(?P<marker>[A-Z]\.|\.\d+)"),
    "SUBPARAGRAPH": re.compile(r"^\s*(?P<marker>\d+\.|\.\d+)"),
    "SUBSUBPARAGRAPH": re.compile(r"^\s*(?P<marker>[a-z]\.|\.\d+)"),
    "SUBPARAGRAPH_LEVEL_5": re.compile(r"^\s*(?P<marker>\d+\)|\.\d+)"),
    "SUBPARAGRAPH_LEVEL_6": re.compile(r"^\s*(?P<marker>[a-z]\)|\.\d+)"),
    "SUBPARAGRAPH_LEVEL_7": re.compile(r"^\s*(?P<marker>\(\d+\)|\.\d+)"),
    "SUBPARAGRAPH_LEVEL_8": re.compile(r"^\s*(?P<marker>\([a-z]\)|\.\d+)"),
}
_ANY_MARKERS = (
    re.compile(r"^\s*\d{1,2}\.\d{1,3}(?=\s|$)"),
    re.compile(r"^\s*\.\d+(?=\s|$)"),
    re.compile(r"^\s*\(\d+\)(?=\s|$)"),
    re.compile(r"^\s*\([a-z]\)(?=\s|$)"),
    re.compile(r"^\s*[A-Z][.)](?=\s|$)"),
    re.compile(r"^\s*[a-z][.)](?=\s|$)"),
    re.compile(r"^\s*\d+[.)](?=\s|$)"),
)
_TEXT_NODE_RX = re.compile(
    r"(?P<open><w:t\b[^>]*>)(?P<text>[\s\S]*?)(?P<close></w:t>)"
)
_TAB_RX = re.compile(r"<w:tab\b[^>]*/>|<w:tab\b[^>]*>\s*</w:tab>", re.S)
_TRACKED_OR_FIELD_RX = re.compile(
    r"<w:(?:ins|del|moveFrom|moveTo|instrText|fldChar|fldSimple)\b"
)
_BREAK_RX = re.compile(r"<w:(?:br|cr)\b")
_W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
def _wq(local_name: str) -> str:
    return f"{{{_W_NS}}}{local_name}"


def _marker_family(role: str, marker: str) -> str:
    if role == "PART":
        return "csi_part"
    if role == "ARTICLE":
        second = marker.split(".", 1)[1]
        return "csi_article" if len(second) > 1 and second.startswith("0") else "csc_article"
    if marker.startswith("."):
        return "csc_dot_decimal"
    if re.fullmatch(r"[A-Z][.)]", marker):
        return "upper_alpha"
    if re.fullmatch(r"[a-z][.)]", marker):
        return "lower_alpha"
    return "decimal"


def _detect_literal_marker(
    text: str,
    role: str,
    *,
    raw_xml_text: bool = False,
) -> Optional[_LiteralMarker]:
    pattern = (_RAW_ROLE_MARKERS if raw_xml_text else _ROLE_MARKERS).get(role)
    if pattern is None:
        return None
    match = pattern.match(text)
    if match is None:
        return None
    marker = match.group("marker")
    separator_chars = (
        " \t\u00a0-\u2010\u2011\u2012\u2013\u2014\u2015:"
        if role == "PART"
        else " \t\u00a0"
    )
    body = text[match.end():].lstrip(separator_chars)
    return _LiteralMarker(marker, _marker_family(role, marker), body)


def _detect_any_literal_marker(text: str) -> Optional[str]:
    for pattern in _ANY_MARKERS:
        match = pattern.match(text)
        if match is not None:
            return match.group(0).strip()
    return None


def _has_heading_like_article_body(body_text: str) -> bool:
    """Require a heading-like word before accepting a spaced ``1.1`` marker."""

    first_letter = next((char for char in body_text if char.isalpha()), "")
    return bool(first_letter and first_letter.isupper())


def _decoded_text_segments(paragraph_xml: str):
    segments = []
    cursor = 0
    for match in _TEXT_NODE_RX.finditer(paragraph_xml):
        raw_inner = match.group("text")
        if "<" in raw_inner:
            raise EngineError("canadian_target_markup", "A Word text node contains unsupported nested markup")
        decoded = html.unescape(raw_inner)
        start = cursor
        cursor += len(decoded)
        segments.append((match, decoded, start, cursor))
    return segments, "".join(item[1] for item in segments)


def _marker_markup_delimiter(paragraph_xml: str, role: str) -> Optional[str]:
    """Return a structural delimiter immediately following a typed marker."""

    segments, joined = _decoded_text_segments(paragraph_xml)
    marker_match = _RAW_ROLE_MARKERS[role].match(joined)
    if marker_match is None:
        return None
    marker_end = marker_match.end()
    for index, (match, _text, start, end) in enumerate(segments):
        if not (start < marker_end <= end or (marker_end == 0 and index == 0)):
            continue
        if marker_end < end:
            return None
        next_start = (
            segments[index + 1][0].start()
            if index + 1 < len(segments)
            else len(paragraph_xml)
        )
        between = paragraph_xml[match.end():next_start]
        tab = _TAB_RX.search(between)
        line_break = _BREAK_RX.search(between)
        if line_break is not None and (
            tab is None or line_break.start() < tab.start()
        ):
            return "line_break"
        if tab is not None:
            return "tab"
        return None
    return None


def _roman_to_int(value: str) -> int:
    values = {"I": 1, "V": 5, "X": 10, "L": 50, "C": 100, "D": 500, "M": 1000}
    total = 0
    previous = 0
    for char in reversed(value.upper()):
        current = values[char]
        total += -current if current < previous else current
        previous = max(previous, current)
    return total


def _alpha_to_int(value: str) -> int:
    total = 0
    for char in value.upper():
        total = total * 26 + (ord(char) - ord("A") + 1)
    return total


def _literal_counter(role: str, literal: _LiteralMarker) -> Tuple[Optional[int], int]:
    marker = literal.marker.strip()
    if role == "PART":
        value = marker.split(None, 1)[1]
        return None, int(value) if value.isdigit() else _roman_to_int(value)
    if role == "ARTICLE":
        part, article = marker.split(".", 1)
        return int(part), int(article)
    if marker.startswith("."):
        return None, int(marker[1:])
    value = marker[1:-1] if marker.startswith("(") else marker[:-1]
    if role in {
        "PARAGRAPH",
        "SUBSUBPARAGRAPH",
        "SUBPARAGRAPH_LEVEL_6",
        "SUBPARAGRAPH_LEVEL_8",
    }:
        return None, _alpha_to_int(value)
    return None, int(value)


def _no_locator(index: int) -> str:
    return ""


def _paragraph_locator(
    blocks: list,
    role_by_index: Dict[int, str],
) -> Callable[[int], str]:
    """Return a describer that places a paragraph by SECTION number and heading.

    Engine messages name paragraphs by their ``word/document.xml`` index,
    which nobody can find in Word. The describer appends
    `` (Section 21 13 13, heading 5)``: the number on the nearest preceding
    SectionID paragraph and the 1-based ordinal of the paragraph among the
    PART and numbered-role headings after that SECTION line. A paragraph
    that is not itself a heading reports ``after heading 5`` (or ``before
    its first heading``), and one ahead of every SECTION line says so.
    Only the section number and counts are reported, never body text.
    """

    section_indices: list[int] = []
    section_labels: list[str] = []
    heading_indices: list[int] = []
    for index in sorted(role_by_index):
        role = role_by_index[index]
        if role == "SectionID" and index < len(blocks):
            number = section_number_display_form(
                paragraph_text_from_block(blocks[index][2])
            )
            section_indices.append(index)
            section_labels.append(
                f"Section {number}" if number else "an unnumbered SECTION line"
            )
        elif role in _LOCATOR_HEADING_ROLES:
            heading_indices.append(index)

    def describe(index: int) -> str:
        position = bisect.bisect_right(section_indices, index) - 1
        if position >= 0:
            section_index = section_indices[position]
            label = section_labels[position]
        else:
            section_index = -1
            label = "before any SECTION line"
        first = bisect.bisect_right(heading_indices, section_index)
        last = bisect.bisect_right(heading_indices, index)
        ordinal = last - first
        if ordinal and role_by_index.get(index) in _LOCATOR_HEADING_ROLES:
            place = f"heading {ordinal}"
        elif ordinal:
            place = f"after heading {ordinal}"
        else:
            place = "before its first heading"
        return f" ({label}, {place})"

    return describe


def _find_numbering_instance(root: ET.Element, num_id: str) -> Optional[ET.Element]:
    return next(
        (
            node
            for node in root.findall(_wq("num"))
            if node.attrib.get(_wq("numId")) == num_id
        ),
        None,
    )


def _find_numbering_level(
    root: ET.Element,
    num_id: str,
    ilvl: str,
) -> Tuple[ET.Element, Optional[ET.Element]]:
    num = _find_numbering_instance(root, num_id)
    if num is None:
        raise EngineError("canadian_numbering_unprovable", f"Numbering instance numId={num_id!r} is missing")
    override = next(
        (
            node
            for node in num.findall(_wq("lvlOverride"))
            if node.attrib.get(_wq("ilvl")) == ilvl
        ),
        None,
    )
    abstract_ref = num.find(_wq("abstractNumId"))
    abstract_id = (
        abstract_ref.attrib.get(_wq("val")) if abstract_ref is not None else None
    )
    abstract = next(
        (
            node
            for node in root.findall(_wq("abstractNum"))
            if node.attrib.get(_wq("abstractNumId")) == abstract_id
        ),
        None,
    )
    if abstract is None:
        raise EngineError("canadian_numbering_unprovable", 
            f"Numbering instance numId={num_id!r} references a missing abstract list"
        )
    level = next(
        (
            node
            for node in abstract.findall(_wq("lvl"))
            if node.attrib.get(_wq("ilvl")) == ilvl
        ),
        None,
    )
    override_level = override.find(_wq("lvl")) if override is not None else None
    effective_level = override_level if override_level is not None else level
    if effective_level is None:
        raise EngineError("canadian_numbering_unprovable", 
            f"Numbering instance numId={num_id!r} has no level ilvl={ilvl!r}"
        )
    return effective_level, override


def _validate_numbering_start(
    level: ET.Element,
    override: Optional[ET.Element],
    *,
    context: str,
    reject_override: bool,
) -> None:
    if override is not None and reject_override:
        raise EngineError("canadian_numbering_unprovable", 
            f"{context} uses a list-level override; its counter state cannot be "
            "proven without a Word numbering walker."
        )
    start_override = override.find(_wq("startOverride")) if override is not None else None
    start = level.find(_wq("start"))
    start_value = None
    if start_override is not None:
        start_value = start_override.attrib.get(_wq("val"))
    elif start is not None:
        start_value = start.attrib.get(_wq("val"))
    if start_value not in {None, "1"}:
        raise EngineError("canadian_numbering_unprovable", f"{context} starts at {start_value!r}, not 1")
    if level.find(_wq("lvlRestart")) is not None:
        raise EngineError("canadian_numbering_unprovable", 
            f"{context} uses an explicit restart rule that Canadian conversion "
            "cannot yet prove safe."
        )


def _with_preserve_space(open_tag: str, text: str) -> str:
    if text and (text[0].isspace() or text[-1].isspace()):
        if not re.search(r"\bxml:space\s*=", open_tag):
            return open_tag[:-1] + ' xml:space="preserve">'
    return open_tag


def _remove_marker_from_unprotected_xml(
    paragraph_xml: str,
    role: str,
) -> Tuple[str, bool]:
    if _TRACKED_OR_FIELD_RX.search(paragraph_xml):
        raise EngineError("canadian_target_markup", 
            "Leading numbering crosses or shares a tracked-change/field paragraph; "
            "accept the changes or unlink the field before Canadian conversion."
        )

    segments, joined = _decoded_text_segments(paragraph_xml)
    marker = _detect_literal_marker(joined, role, raw_xml_text=True)
    if marker is None:
        raise EngineError("canadian_target_markup", "Could not map the visible leading marker to Word text nodes")
    marker_match = _RAW_ROLE_MARKERS[role].match(joined)
    assert marker_match is not None
    remove_end = marker_match.end()
    separator_chars = (
        " \t\u00a0-\u2010\u2011\u2012\u2013\u2014\u2015:"
        if role == "PART"
        else " \t\u00a0"
    )
    while remove_end < len(joined) and joined[remove_end] in separator_chars:
        remove_end += 1

    # A manually typed marker is frequently followed by a real Word tab rather
    # than a space inside w:t. Remove that delimiter as part of the marker so
    # the architect's automatic numbering cannot create doubled spacing.
    tab_removed = False
    containing_index = next(
        (
            index
            for index, (_match, _text, start, end) in enumerate(segments)
            if start < remove_end <= end or (remove_end == 0 and index == 0)
        ),
        None,
    )
    if containing_index is not None:
        match, _text, _start, end = segments[containing_index]
        if remove_end == end:
            next_start = (
                segments[containing_index + 1][0].start()
                if containing_index + 1 < len(segments)
                else len(paragraph_xml)
            )
            between = paragraph_xml[match.end():next_start]
            tab = _TAB_RX.search(between)
            if tab is not None:
                absolute_start = match.end() + tab.start()
                absolute_end = match.end() + tab.end()
                paragraph_xml = (
                    paragraph_xml[:absolute_start] + paragraph_xml[absolute_end:]
                )
                tab_removed = True
                segments, joined = _decoded_text_segments(paragraph_xml)

    pieces = []
    last = 0
    for match, decoded, start, end in segments:
        pieces.append(paragraph_xml[last:match.start()])
        if start >= remove_end:
            pieces.append(match.group(0))
        else:
            kept = decoded[max(0, remove_end - start):] if end > remove_end else ""
            opening = _with_preserve_space(match.group("open"), kept)
            pieces.append(opening + html.escape(kept, quote=False) + match.group("close"))
        last = match.end()
    pieces.append(paragraph_xml[last:])
    return "".join(pieces), tab_removed


def _remove_literal_marker(paragraph_xml: str, role: str) -> Tuple[str, bool]:
    return_result: Tuple[str, bool] = (paragraph_xml, False)

    def _edit(unprotected: str) -> str:
        nonlocal return_result
        return_result = _remove_marker_from_unprotected_xml(unprotected, role)
        return return_result[0]

    edited = edit_preserving_out_of_scope_subtrees(paragraph_xml, _edit)
    return edited, return_result[1]


def _text_skeleton(paragraph_xml: str) -> str:
    return _TEXT_NODE_RX.sub(
        lambda match: match.group("open") + match.group("close"),
        paragraph_xml,
    )


def _remove_first_tab(xml_text: str) -> str:
    return _TAB_RX.sub("", xml_text, count=1)


def _verify_changed_paragraph(
    before: str,
    after: str,
    expected_body: str,
    *,
    tab_removed: bool,
) -> None:
    actual_body = paragraph_text_from_block(after)
    if actual_body != expected_body:
        raise RuntimeError(
            "Canadian conversion invariant failed: substantive paragraph text changed "
            f"({expected_body!r} != {actual_body!r})"
        )
    before_skeleton = _text_skeleton(before)
    if tab_removed:
        before_skeleton = _remove_first_tab(before_skeleton)
    if before_skeleton != _text_skeleton(after):
        raise RuntimeError(
            "Canadian conversion invariant failed: content outside the leading marker changed"
        )
    for name in OUT_OF_SCOPE_SUBTREE_NAMES:
        before_blocks = [item[2] for item in iter_element_xml_blocks(before, name)]
        after_blocks = [item[2] for item in iter_element_xml_blocks(after, name)]
        if before_blocks != after_blocks:
            raise RuntimeError(
                "Canadian conversion invariant failed: an out-of-scope drawing, object, "
                "or text box changed"
            )


def _validate_source_sequence(
    evidence: list[_SourceEvidence],
    *,
    describe: Callable[[int], str] = _no_locator,
    error_code: str = "canadian_target_hierarchy",
) -> None:
    """Prove that regenerated counters preserve a canonical source sequence.

    ``describe`` renders the SECTION/heading locator appended to a paragraph
    index in every message (see :func:`_paragraph_locator`).

    Both converters need this proof and both need it to fail under their own
    error code, so ``error_code`` selects which one the caller reports. The
    rules themselves are direction-independent: a hierarchy with a missing
    parent or a counter gap is unconvertible whichever way it is going.
    """

    source_kinds: Dict[str, set[str]] = {}
    for item in evidence:
        source_kinds.setdefault(item.role, set()).add(item.source_kind)
    for role, kinds in source_kinds.items():
        if len(kinds) > 1:
            raise EngineError(error_code, 
                f"Cannot safely mix typed and automatic {role} numbering in one target."
            )

    roles_present = {item.role for item in evidence}
    active: Dict[str, Optional[int]] = {}
    for item in evidence:
        level = _ROLE_LEVEL[item.role]
        for deeper_role, deeper_level in _ROLE_LEVEL.items():
            if deeper_level > level:
                active.pop(deeper_role, None)

        parent = ROLE_PARENT.get(item.role)
        if item.role == "PARAGRAPH":
            parent = None
        if item.role == "PARAGRAPH" and "ARTICLE" in roles_present:
            parent = "ARTICLE"
        if parent is not None and parent not in active:
            raise EngineError(error_code, 
                f"Paragraph {item.paragraph_index}{describe(item.paragraph_index)} "
                f"is {item.role} without a preceding "
                f"{parent}; the hierarchy cannot be proven."
            )

        if item.source_kind == "automatic":
            active[item.role] = None
            continue

        assert item.literal is not None
        parent_number, counter = _literal_counter(item.role, item.literal)
        if item.role == "ARTICLE":
            part_number = active.get("PART")
            if part_number is not None and parent_number != part_number:
                raise EngineError(error_code, 
                    f"Paragraph {item.paragraph_index}{describe(item.paragraph_index)} "
                    f"article {item.literal.marker!r} "
                    f"does not belong to the active PART {part_number}."
                )

        previous = active.get(item.role)
        expected = 1 if previous is None else previous + 1
        if counter != expected:
            raise EngineError(error_code, 
                f"Paragraph {item.paragraph_index}{describe(item.paragraph_index)} "
                f"has non-contiguous {item.role} marker "
                f"{item.literal.marker!r}; expected counter {expected}. Canadian conversion "
                "does not silently repair gaps or restarts."
            )
        active[item.role] = counter


def _validate_automatic_source(
    item: _SourceEvidence,
    numbering_root: Optional[ET.Element],
    numbering_catalog: Dict[str, Any],
    available_roles: set[str],
    *,
    describe: Callable[[int], str] = _no_locator,
    error_code: str = "canadian_numbering_unprovable",
) -> None:
    """Prove a paragraph's automatic numbering really is the role it was given.

    Both converters need this and for the same reason: the classified role and
    the level Word is actually rendering must be the same thing. The forward
    converter would otherwise retarget a paragraph to the wrong architect
    level; the reverse converter would write the wrong number into the text as
    literal characters, which is worse because nothing downstream can correct
    it. ``error_code`` selects which converter's code the caller reports.
    """

    where = describe(item.paragraph_index)
    if numbering_root is None or item.automatic_numpr is None:
        raise EngineError(error_code, 
            f"Paragraph {item.paragraph_index}{where} uses automatic numbering, but the "
            "target numbering.xml is unavailable."
        )
    pattern = item.automatic_pattern
    if not isinstance(pattern, dict):
        raise EngineError(error_code, 
            f"Paragraph {item.paragraph_index}{where} automatic numbering cannot be "
            "resolved."
        )
    num_id = str(item.automatic_numpr["numId"])
    ilvl = str(item.automatic_numpr.get("ilvl", "0"))
    inferred = role_from_numbering_catalog(
        numbering_catalog,
        num_id,
        ilvl,
    )
    if inferred is None:
        inferred = role_from_numbering_signature(
            pattern.get("numFmt"), pattern.get("lvlText"), pattern.get("ilvl")
        )
    resolved = next(
        (
            candidate
            for candidate in ROLE_FALLBACKS.get(inferred, (inferred,))
            if candidate in available_roles
        ),
        None,
    ) if inferred is not None else None
    if resolved != item.role:
        raise EngineError(error_code, 
            f"Paragraph {item.paragraph_index}{where} is classified as {item.role}, but its "
            f"automatic numbering signature resolves to {inferred or 'no safe role'}"
            + (
                f" (available-role fallback: {resolved})."
                if resolved is not None and resolved != inferred
                else "."
            )
        )
    level, override = _find_numbering_level(numbering_root, num_id, ilvl)
    _validate_numbering_start(
        level,
        override,
        context=f"Paragraph {item.paragraph_index}{where} source numbering",
        reject_override=True,
    )


__all__ = [
    "NUMBERED_ROLES",
    "_ANY_MARKERS",
    "_BREAK_RX",
    "_LiteralMarker",
    "_RAW_ROLE_MARKERS",
    "_ROLE_MARKERS",
    "_SourceEvidence",
    "_TAB_RX",
    "_TEXT_NODE_RX",
    "_TRACKED_OR_FIELD_RX",
    "_W_NS",
    "_alpha_to_int",
    "_decoded_text_segments",
    "_detect_any_literal_marker",
    "_detect_literal_marker",
    "_find_numbering_instance",
    "_find_numbering_level",
    "_has_heading_like_article_body",
    "_literal_counter",
    "_marker_family",
    "_marker_markup_delimiter",
    "_no_locator",
    "_paragraph_locator",
    "_remove_first_tab",
    "_remove_literal_marker",
    "_remove_marker_from_unprotected_xml",
    "_roman_to_int",
    "_text_skeleton",
    "_validate_automatic_source",
    "_validate_numbering_start",
    "_validate_source_sequence",
    "_verify_changed_paragraph",
    "_with_preserve_space",
    "_wq",
]
