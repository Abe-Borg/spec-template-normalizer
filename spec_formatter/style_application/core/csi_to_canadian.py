"""Fail-closed CSI-to-Canadian PageFormat hierarchy conversion.

The converter deliberately changes only leading, typed CSI/CSC numbering
markers.  The selected architect template remains the sole source of the
rendered Canadian numbering, styles, and layout.  Consequently Canadian mode
requires the architect's numbered roles to use true Word automatic numbering
with Canadian numeric signatures.
"""

from __future__ import annotations

import html
import re
import xml.etree.ElementTree as ET
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Dict, Optional, Tuple

from spec_formatter.numbering_roles import (
    role_from_numbering_catalog,
    role_from_numbering_signature,
)
from spec_formatter.role_contract import (
    BODY_HIERARCHY_ROLES,
    NUMBERED_BODY_ROLES,
    ROLE_LEVEL,
    ROLE_PARENT,
)

from .classification import (
    _build_numbering_catalog,
    _effective_numpr,
    _resolve_numbering_pattern,
)
from .ooxml_text import prepare_xml_text_for_utf8, read_xml_text, write_xml_text
from .sectpr_tools import extract_all_sectpr_blocks
from .xml_helpers import (
    OUT_OF_SCOPE_SUBTREE_NAMES,
    edit_preserving_out_of_scope_subtrees,
    iter_element_xml_blocks,
    iter_paragraph_xml_blocks,
    paragraph_text_from_block,
    strip_out_of_scope_subtrees,
)


FORMAT_ONLY = "format_only"
CSI_TO_CANADIAN = "csi_to_canadian"
VALID_CONVERSION_MODES = frozenset({FORMAT_ONLY, CSI_TO_CANADIAN})
NUMBERED_ROLES = NUMBERED_BODY_ROLES


def validate_conversion_mode(value: object) -> str:
    """Return a validated mode string or raise before any document work."""

    if not isinstance(value, str) or value not in VALID_CONVERSION_MODES:
        choices = ", ".join(sorted(VALID_CONVERSION_MODES))
        raise ValueError(f"conversion_mode must be one of: {choices}")
    return value


@dataclass(frozen=True)
class ConversionIssue:
    paragraph_index: int
    code: str
    message: str
    text_preview: str

    def as_dict(self) -> Dict[str, Any]:
        return {
            "paragraph_index": self.paragraph_index,
            "code": self.code,
            "message": self.message,
            "text_preview": self.text_preview,
        }


@dataclass(frozen=True)
class MarkerEdit:
    paragraph_index: int
    role: str
    source_kind: str
    target_kind: str
    source_marker: Optional[str]
    target_marker: Optional[str]

    def as_dict(self) -> Dict[str, Any]:
        return {
            "paragraph_index": self.paragraph_index,
            "role": self.role,
            "source_kind": self.source_kind,
            "target_kind": self.target_kind,
            "source_marker": self.source_marker,
            "target_marker": self.target_marker,
        }


@dataclass(frozen=True)
class CanadianConversionReport:
    paragraphs_examined: int
    paragraphs_converted: int
    literal_markers_removed: int
    automatic_numbering_retargeted: int
    unnumbered_paragraphs_numbered: int
    edits: Tuple[MarkerEdit, ...]
    warnings: Tuple[ConversionIssue, ...]

    def as_dict(self) -> Dict[str, Any]:
        return {
            "paragraphs_examined": self.paragraphs_examined,
            "paragraphs_converted": self.paragraphs_converted,
            "literal_markers_removed": self.literal_markers_removed,
            "automatic_numbering_retargeted": self.automatic_numbering_retargeted,
            "unnumbered_paragraphs_numbered": self.unnumbered_paragraphs_numbered,
            "edits": [item.as_dict() for item in self.edits],
            "warnings": [item.as_dict() for item in self.warnings],
        }


@dataclass(frozen=True)
class ConversionPlan:
    document_xml: str
    report: CanadianConversionReport


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
    re.compile(r"^\s*\d{1,2}\.\d{1,3}"),
    re.compile(r"^\s*\.\d+"),
    re.compile(r"^\s*\(\d+\)"),
    re.compile(r"^\s*\([a-z]\)"),
    re.compile(r"^\s*[A-Z][.)]"),
    re.compile(r"^\s*[a-z][.)]"),
    re.compile(r"^\s*\d+[.)]"),
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


def _validate_canadian_role_contract(
    role: str,
    spec: object,
    *,
    allow_numeric_part: bool = False,
) -> None:
    if not isinstance(spec, dict):
        raise ValueError(
            f"Canadian conversion requires a complete architect role contract for {role}."
        )
    provenance = spec.get("numbering_provenance")
    if provenance not in {"style_numpr", "direct_numpr"}:
        raise ValueError(
            "Canadian conversion requires true Word automatic numbering in the "
            f"architect template for {role}; found {provenance!r}."
        )
    pattern = spec.get("numbering_pattern")
    if not isinstance(pattern, dict):
        raise ValueError(
            f"Canadian conversion is missing the architect numbering pattern for {role}."
        )
    num_fmt = str(pattern.get("numFmt") or "")
    lvl_text = str(pattern.get("lvlText") or "")
    for key in ("start", "startOverride"):
        value = pattern.get(key)
        if value is not None and str(value) != "1":
            raise ValueError(
                f"Architect role {role} starts at {value!r}; Canadian conversion "
                "requires numbering that starts at 1."
            )
    if pattern.get("lvlRestart") is not None:
        raise ValueError(
            f"Architect role {role} uses an explicit numbering restart rule that "
            "Canadian conversion cannot yet prove safe."
        )
    if num_fmt != "decimal":
        raise ValueError(
            f"Architect role {role} is not Canadian numeric numbering "
            f"(numFmt={num_fmt!r})."
        )
    if role == "PART":
        labeled_part = re.fullmatch(
            r"\s*PART\s+%\d+\s*", lvl_text, re.IGNORECASE
        )
        numeric_part = re.fullmatch(r"\s*%\d+\.?\s*", lvl_text)
        valid = labeled_part or (numeric_part if allow_numeric_part else None)
        expected = "PART %1 (or %1 / %1. in a proven CSC hierarchy)"
    elif role == "ARTICLE":
        valid = re.fullmatch(r"\s*%\d+\s*\.\s*%\d+\s*", lvl_text)
        expected = "%1.%2"
    else:
        valid = re.fullmatch(r"\s*\.\s*%\d+\s*", lvl_text)
        expected = ".%n"
    if valid is None:
        raise ValueError(
            f"Architect role {role} does not demonstrate Canadian PageFormat "
            f"numbering (lvlText={lvl_text!r}; expected a pattern like {expected!r})."
        )


def _validate_complete_article_hierarchy(
    role_specs: Dict[str, Dict[str, Any]],
    roles_in_target: set[str],
) -> None:
    """Require one coherent multilevel list when articles are converted."""

    if "PART" not in roles_in_target:
        raise ValueError(
            "Canadian article conversion requires a classified PART heading in the target "
            "so Word can establish the article's part number."
        )
    _validate_canadian_role_contract(
        "PART",
        role_specs.get("PART"),
        allow_numeric_part=True,
    )
    expected_levels = {
        "PART": (
            "0",
            re.compile(r"\s*(?:PART\s+%1|%1\.?)\s*", re.IGNORECASE),
        ),
        "ARTICLE": ("1", re.compile(r"\s*%1\s*\.\s*%2\s*")),
        **{
            role: (
                str(ROLE_LEVEL[role]),
                re.compile(rf"\s*\.\s*%{ROLE_LEVEL[role] + 1}\s*"),
            )
            for role in BODY_HIERARCHY_ROLES[2:]
        },
    }
    relevant = [role for role in expected_levels if role in roles_in_target]
    reference_num_id: Optional[str] = None
    for role in relevant:
        spec = role_specs.get(role)
        _validate_canadian_role_contract(
            role,
            spec,
            allow_numeric_part=(role == "PART"),
        )
        assert isinstance(spec, dict)
        pattern = spec["numbering_pattern"]
        num_id = str(pattern.get("numId") or "")
        ilvl = str(pattern.get("ilvl") or "0")
        lvl_text = str(pattern.get("lvlText") or "")
        expected_ilvl, expected_text = expected_levels[role]
        if ilvl != expected_ilvl or expected_text.fullmatch(lvl_text) is None:
            raise ValueError(
                "Canadian article conversion requires a coherent PART/article/list "
                f"hierarchy; architect role {role} has ilvl={ilvl!r}, "
                f"lvlText={lvl_text!r}."
            )
        if not num_id:
            raise ValueError(
                f"Architect role {role} is missing its Word numbering list identifier."
            )
        if reference_num_id is None:
            reference_num_id = num_id
        elif num_id != reference_num_id:
            raise ValueError(
                "Canadian PART, article, and subordinate roles must share one Word "
                "multilevel numbering list."
            )


def _decoded_text_segments(paragraph_xml: str):
    segments = []
    cursor = 0
    for match in _TEXT_NODE_RX.finditer(paragraph_xml):
        raw_inner = match.group("text")
        if "<" in raw_inner:
            raise ValueError("A Word text node contains unsupported nested markup")
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


_ROLE_LEVEL = ROLE_LEVEL


def _validate_source_sequence(evidence: list[_SourceEvidence]) -> None:
    """Prove that regenerated counters preserve a canonical source sequence."""

    source_kinds: Dict[str, set[str]] = {}
    for item in evidence:
        source_kinds.setdefault(item.role, set()).add(item.source_kind)
    for role, kinds in source_kinds.items():
        if len(kinds) > 1:
            raise ValueError(
                f"Canadian conversion cannot safely mix typed and automatic {role} "
                "numbering in one target."
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
            raise ValueError(
                f"Paragraph {item.paragraph_index} is {item.role} without a preceding "
                f"{parent}; Canadian conversion cannot prove the hierarchy."
            )

        if item.source_kind == "automatic":
            active[item.role] = None
            continue

        assert item.literal is not None
        parent_number, counter = _literal_counter(item.role, item.literal)
        if item.role == "ARTICLE":
            part_number = active.get("PART")
            if part_number is not None and parent_number != part_number:
                raise ValueError(
                    f"Paragraph {item.paragraph_index} article {item.literal.marker!r} "
                    f"does not belong to the active PART {part_number}."
                )

        previous = active.get(item.role)
        expected = 1 if previous is None else previous + 1
        if counter != expected:
            raise ValueError(
                f"Paragraph {item.paragraph_index} has non-contiguous {item.role} marker "
                f"{item.literal.marker!r}; expected counter {expected}. Canadian conversion "
                "does not silently repair gaps or restarts."
            )
        active[item.role] = counter


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
        raise ValueError(f"Numbering instance numId={num_id!r} is missing")
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
        raise ValueError(
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
        raise ValueError(
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
        raise ValueError(
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
        raise ValueError(f"{context} starts at {start_value!r}, not 1")
    if level.find(_wq("lvlRestart")) is not None:
        raise ValueError(
            f"{context} uses an explicit restart rule that Canadian conversion "
            "cannot yet prove safe."
        )


def _validate_automatic_source(
    item: _SourceEvidence,
    numbering_root: Optional[ET.Element],
    numbering_catalog: Dict[str, Any],
) -> None:
    if numbering_root is None or item.automatic_numpr is None:
        raise ValueError(
            f"Paragraph {item.paragraph_index} uses automatic numbering, but the "
            "target numbering.xml is unavailable."
        )
    pattern = item.automatic_pattern
    if not isinstance(pattern, dict):
        raise ValueError(
            f"Paragraph {item.paragraph_index} automatic numbering cannot be resolved."
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
    if inferred != item.role:
        raise ValueError(
            f"Paragraph {item.paragraph_index} is classified as {item.role}, but its "
            f"automatic numbering signature resolves to {inferred or 'no safe role'}."
        )
    level, override = _find_numbering_level(numbering_root, num_id, ilvl)
    _validate_numbering_start(
        level,
        override,
        context=f"Paragraph {item.paragraph_index} source numbering",
        reject_override=True,
    )


def _validate_architect_numbering(
    numbering_xml: str,
    role_specs: Dict[str, Dict[str, Any]],
    roles: set[str],
) -> None:
    if not numbering_xml.strip():
        raise ValueError(
            "Canadian conversion requires the architect template's numbering.xml."
        )
    root = ET.fromstring(prepare_xml_text_for_utf8(numbering_xml).encode("utf-8"))
    for role in sorted(roles):
        spec = role_specs[role]
        pattern = spec["numbering_pattern"]
        num_id = str(pattern.get("numId") or "")
        ilvl = str(pattern.get("ilvl") or "0")
        level, override = _find_numbering_level(root, num_id, ilvl)
        _validate_numbering_start(
            level,
            override,
            context=f"Architect role {role}",
            reject_override=False,
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
        raise ValueError(
            "Leading numbering crosses or shares a tracked-change/field paragraph; "
            "accept the changes or unlink the field before Canadian conversion."
        )

    segments, joined = _decoded_text_segments(paragraph_xml)
    marker = _detect_literal_marker(joined, role, raw_xml_text=True)
    if marker is None:
        raise ValueError("Could not map the visible leading marker to Word text nodes")
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


def plan_csi_to_canadian(
    document_xml: str,
    styles_xml: str,
    classifications: Dict[str, Any],
    role_specs: Optional[Dict[str, Dict[str, Any]]],
    *,
    numbering_xml: str = "",
    architect_numbering_xml: Optional[str] = None,
) -> ConversionPlan:
    """Validate and build the complete document edit before writing anything."""

    if not isinstance(role_specs, dict):
        raise ValueError(
            "Canadian conversion requires a strict current architect template profile."
        )
    items = classifications.get("classifications")
    if not isinstance(items, list):
        raise ValueError("Canadian conversion requires final paragraph classifications")

    role_by_index: Dict[int, str] = {}
    for item in items:
        if not isinstance(item, dict):
            raise ValueError("Canadian conversion classification entries must be objects")
        index = item.get("paragraph_index")
        role = item.get("csi_role")
        if not isinstance(index, int) or index < 0 or not isinstance(role, str):
            raise ValueError("Canadian conversion received an invalid classification entry")
        if index in role_by_index:
            raise ValueError(f"Duplicate Canadian conversion classification for paragraph {index}")
        role_by_index[index] = role

    roles_in_target = set(role_by_index.values())
    used_numbered_roles = sorted(roles_in_target & NUMBERED_ROLES)
    for role in used_numbered_roles:
        _validate_canadian_role_contract(role, role_specs.get(role))
    if "ARTICLE" in roles_in_target:
        _validate_complete_article_hierarchy(role_specs, roles_in_target)

    roles_to_convert = set(used_numbered_roles)
    part_spec = role_specs.get("PART")
    if "PART" in roles_in_target and isinstance(part_spec, dict) and part_spec.get(
        "numbering_provenance"
    ) in {"style_numpr", "direct_numpr"}:
        _validate_canadian_role_contract(
            "PART",
            part_spec,
            allow_numeric_part=("ARTICLE" in roles_in_target),
        )
        roles_to_convert.add("PART")
    if architect_numbering_xml is not None:
        _validate_architect_numbering(
            architect_numbering_xml,
            role_specs,
            roles_to_convert,
        )

    numbering_catalog = _build_numbering_catalog(numbering_xml)
    numbering_root = (
        ET.fromstring(prepare_xml_text_for_utf8(numbering_xml).encode("utf-8"))
        if numbering_xml.strip()
        else None
    )

    blocks = list(iter_paragraph_xml_blocks(document_xml))
    replacements: Dict[int, str] = {}
    evidence: list[_SourceEvidence] = []
    edits: list[MarkerEdit] = []
    warnings: list[ConversionIssue] = []
    literal_removed = 0
    automatic_retargeted = 0

    for index, role in sorted(role_by_index.items()):
        if role not in roles_to_convert:
            continue
        if index >= len(blocks):
            raise ValueError(f"Canadian conversion paragraph index is out of range: {index}")
        paragraph = blocks[index][2]
        text = paragraph_text_from_block(paragraph)
        literal = _detect_literal_marker(text, role)
        any_literal = _detect_any_literal_marker(text)
        analysis = strip_out_of_scope_subtrees(paragraph)
        automatic_numpr = _effective_numpr(analysis, styles_xml)
        automatic_pattern = _resolve_numbering_pattern(
            automatic_numpr,
            numbering_catalog,
        )
        automatic = automatic_numpr is not None

        if literal is None and any_literal is not None:
            raise ValueError(
                f"Paragraph {index} is classified as {role} but starts with incompatible "
                f"marker {any_literal!r}."
            )
        if literal is not None and automatic:
            raise ValueError(
                f"Paragraph {index} has both automatic numbering and typed marker "
                f"{literal.marker!r}; remove the doubled numbering before conversion."
            )

        if literal is not None:
            delimiter = _marker_markup_delimiter(analysis, role)
            if delimiter == "line_break":
                raise ValueError(
                    f"Paragraph {index} uses a line break after typed marker "
                    f"{literal.marker!r}; only a normal space or Word list tab is safe."
                )
            if (
                literal.family in {"csc_article", "csc_dot_decimal"}
                and delimiter != "tab"
            ):
                raise ValueError(
                    f"Paragraph {index} begins with ambiguous decimal text "
                    f"{literal.marker!r}. A typed Canadian marker is converted only "
                    "when followed by a structural Word tab."
                )
            evidence.append(
                _SourceEvidence(index, role, "literal", literal, None, None)
            )
            converted, tab_removed = _remove_literal_marker(paragraph, role)
            _verify_changed_paragraph(
                paragraph,
                converted,
                literal.body_text,
                tab_removed=tab_removed,
            )
            replacements[index] = converted
            literal_removed += 1
            source_kind = (
                "already_canadian"
                if literal.family in {"csc_article", "csc_dot_decimal"}
                else "literal"
            )
            edits.append(
                MarkerEdit(index, role, source_kind, "automatic", literal.marker, None)
            )
        elif automatic:
            evidence.append(
                _SourceEvidence(
                    index,
                    role,
                    "automatic",
                    None,
                    automatic_numpr,
                    automatic_pattern,
                )
            )
            automatic_retargeted += 1
            edits.append(MarkerEdit(index, role, "automatic", "automatic", None, None))
        else:
            raise ValueError(
                f"Paragraph {index} is classified as numbered role {role}, but it has "
                "neither a recognized typed marker nor Word automatic numbering. "
                "Canadian conversion will not insert an unproven list item."
            )

    automatic_ids: Dict[str, set[str]] = {}
    automatic_by_index: Dict[int, _SourceEvidence] = {}
    for item in evidence:
        if item.source_kind != "automatic":
            continue
        _validate_automatic_source(item, numbering_root, numbering_catalog)
        assert item.automatic_numpr is not None
        automatic_ids.setdefault(item.role, set()).add(item.automatic_numpr["numId"])
        automatic_by_index[item.paragraph_index] = item
    for role, num_ids in automatic_ids.items():
        if len(num_ids) > 1:
            raise ValueError(
                f"Automatic {role} numbering changes list instances ({sorted(num_ids)}); "
                "Canadian conversion cannot prove that the restart will be preserved."
            )
    converted_num_ids = set().union(*automatic_ids.values()) if automatic_ids else set()
    if len(converted_num_ids) > 1:
        raise ValueError(
            "Dependent automatic source roles use different Word list instances "
            f"({sorted(converted_num_ids)}); their counter relationship cannot be proven."
        )
    for index, (_start, _end, block) in enumerate(blocks):
        effective_numpr = _effective_numpr(
            strip_out_of_scope_subtrees(block),
            styles_xml,
        )
        if (
            effective_numpr is None
            or effective_numpr.get("numId") not in converted_num_ids
        ):
            continue
        converted_item = automatic_by_index.get(index)
        if converted_item is None:
            raise ValueError(
                f"Unconverted paragraph {index} shares automatic source list "
                f"numId={effective_numpr.get('numId')!r}; removing it would change "
                "following counters."
            )
    _validate_source_sequence(evidence)

    pieces = []
    last = 0
    for index, (start, end, block) in enumerate(blocks):
        pieces.append(document_xml[last:start])
        pieces.append(replacements.get(index, block))
        last = end
    pieces.append(document_xml[last:])
    converted_document = "".join(pieces)

    after_blocks = list(iter_paragraph_xml_blocks(converted_document))
    if len(after_blocks) != len(blocks):
        raise RuntimeError("Canadian conversion invariant failed: paragraph count changed")
    for index, (_start, _end, before) in enumerate(blocks):
        if index not in replacements and after_blocks[index][2] != before:
            raise RuntimeError(
                f"Canadian conversion invariant failed: untouched paragraph {index} changed"
            )
    if extract_all_sectpr_blocks(document_xml) != extract_all_sectpr_blocks(converted_document):
        raise RuntimeError("Canadian conversion invariant failed: section properties changed")
    ET.fromstring(prepare_xml_text_for_utf8(converted_document).encode("utf-8"))

    report = CanadianConversionReport(
        paragraphs_examined=sum(1 for role in role_by_index.values() if role in roles_to_convert),
        paragraphs_converted=len(edits),
        literal_markers_removed=literal_removed,
        automatic_numbering_retargeted=automatic_retargeted,
        unnumbered_paragraphs_numbered=0,
        edits=tuple(edits),
        warnings=tuple(warnings),
    )
    return ConversionPlan(converted_document, report)


def apply_csi_to_canadian(
    extract_dir: Path,
    classifications: Dict[str, Any],
    role_specs: Optional[Dict[str, Dict[str, Any]]],
    log: list[str],
    *,
    architect_numbering_xml: Optional[str] = None,
) -> CanadianConversionReport:
    """Apply a validated Canadian conversion to one extracted target DOCX."""

    document_path = Path(extract_dir) / "word" / "document.xml"
    styles_path = Path(extract_dir) / "word" / "styles.xml"
    numbering_path = Path(extract_dir) / "word" / "numbering.xml"
    plan = plan_csi_to_canadian(
        read_xml_text(document_path),
        read_xml_text(styles_path),
        classifications,
        role_specs,
        numbering_xml=(
            read_xml_text(numbering_path) if numbering_path.is_file() else ""
        ),
        architect_numbering_xml=architect_numbering_xml,
    )
    write_xml_text(document_path, plan.document_xml)
    report = plan.report
    log.append(
        "Canadian conversion: "
        f"{report.paragraphs_converted} numbered paragraphs; "
        f"removed {report.literal_markers_removed} typed markers; "
        f"retargeted {report.automatic_numbering_retargeted} automatic paragraphs"
    )
    for issue in report.warnings:
        log.append(
            f"Canadian conversion warning p[{issue.paragraph_index}]: {issue.message}"
        )
    return report


__all__ = [
    "CSI_TO_CANADIAN",
    "FORMAT_ONLY",
    "VALID_CONVERSION_MODES",
    "CanadianConversionReport",
    "ConversionIssue",
    "MarkerEdit",
    "apply_csi_to_canadian",
    "plan_csi_to_canadian",
    "validate_conversion_mode",
]
