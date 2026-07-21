"""
Phase 2 classification: applying LLM classifications to paragraphs,
building slim bundles for LLM input, and boilerplate filtering.
"""

import re
import xml.etree.ElementTree as ET
from dataclasses import dataclass, field
from pathlib import Path
from typing import TYPE_CHECKING, Dict, Any, List, Optional, Set, Tuple

if TYPE_CHECKING:
    from .application_policy import ApplicationPolicy

from spec_formatter.numbering_roles import (
    role_from_numbering_catalog,
    role_from_numbering_signature,
)
from spec_formatter.role_contract import BODY_HIERARCHY_ROLES, ROLE_FALLBACKS

from .xml_helpers import (
    iter_paragraph_xml_blocks,
    iter_element_xml_blocks,
    paragraph_text_from_block,
    paragraph_contains_sectpr,
    paragraph_numpr_from_block,
    paragraph_pstyle_from_block,
    paragraph_ppr_hints_from_block,
    apply_pstyle_to_paragraph_block,
    strip_direct_run_properties,
    strip_conflicting_direct_ppr,
    edit_preserving_out_of_scope_subtrees,
    strip_out_of_scope_subtrees,
)
from .style_import import (
    _find_style_numpr_in_chain,
    ensure_explicit_numpr_from_current_style,
)
from .ooxml_text import read_xml_text, write_xml_text


def _load_prompt_text(filename: str) -> str:
    prompt_path = Path(__file__).parent / "prompts" / filename
    return prompt_path.read_text(encoding="utf-8")


PHASE2_MASTER_PROMPT = _load_prompt_text("phase2_master_prompt.txt")
PHASE2_RUN_INSTRUCTION = _load_prompt_text("phase2_run_instruction.txt")


# -------------------------------
# Phase 2: Boilerplate filtering (LLM input only)
# -------------------------------

BOILERPLATE_PATTERNS = [
    # Specifier notes - bracketed formats
    (r'\[Note to [Ss]pecifier[:\s][^\]]*\]', 'specifier_note'),
    (r'\[Specifier[:\s][^\]]*\]', 'specifier_note'),
    (r'\[SPECIFIER[:\s][^\]]*\]', 'specifier_note'),
    (r'(?i)\*\*\s*note to specifier\s*\*\*[^\n]*(?:\n(?!\n)[^\n]*)*', 'specifier_note'),
    (r'(?i)<<\s*note to specifier[^>]*>>', 'specifier_note'),
    (r'(?i)^\s*note to specifier:.*$', 'specifier_note'),

    # MasterSpec / AIA / ARCOM editorial instructions
    (r'(?i)^Retain or delete this article.*$', 'masterspec_instruction'),
    (r'(?i)^Retain [^\n]*paragraph[^\n]*below.*$', 'masterspec_instruction'),
    (r'(?i)^Retain [^\n]*subparagraph[^\n]*below.*$', 'masterspec_instruction'),
    (r'(?i)^Retain [^\n]*article[^\n]*below.*$', 'masterspec_instruction'),
    (r'(?i)^Retain [^\n]*section[^\n]*below.*$', 'masterspec_instruction'),
    (r'(?i)^Retain [^\n]*if .*$', 'masterspec_instruction'),
    (r'(?i)^Retain one of.*$', 'masterspec_instruction'),
    (r'(?i)^Retain one or more of.*$', 'masterspec_instruction'),
    (r'(?i)^Revise this Section by deleting.*$', 'masterspec_instruction'),
    (r'(?i)^Revise [^\n]*to suit [Pp]roject.*$', 'masterspec_instruction'),
    (r'(?i)^This Section uses the term.*$', 'masterspec_instruction'),
    (r'(?i)^Verify that Section titles.*$', 'masterspec_instruction'),
    (r'(?i)^Coordinate [^\n]*paragraph[^\n]* with.*$', 'masterspec_instruction'),
    (r'(?i)^Coordinate [^\n]*revision[^\n]* with.*$', 'masterspec_instruction'),
    (r'(?i)^The list below matches.*$', 'masterspec_instruction'),
    (r'(?i)^See [^\n]*Evaluations?[^\n]* for .*$', 'masterspec_instruction'),
    (r'(?i)^See [^\n]*Article[^\n]* in the Evaluations.*$', 'masterspec_instruction'),
    (r'(?i)^If retaining [^\n]*paragraph.*$', 'masterspec_instruction'),
    (r'(?i)^If retaining [^\n]*subparagraph.*$', 'masterspec_instruction'),
    (r'(?i)^If retaining [^\n]*article.*$', 'masterspec_instruction'),
    (r'(?i)^When [^\n]*characteristics are important.*$', 'masterspec_instruction'),
    (r'(?i)^Inspections in this article are.*$', 'masterspec_instruction'),
    (r'(?i)^Materials and thicknesses in schedules below.*$', 'masterspec_instruction'),
    (r'(?i)^Insulation materials and thicknesses are identified below.*$', 'masterspec_instruction'),
    (r'(?i)^Do not duplicate requirements.*$', 'masterspec_instruction'),
    (r'(?i)^Not all materials and thicknesses may be suitable.*$', 'masterspec_instruction'),
    (r'(?i)^Consider the exposure of installed insulation.*$', 'masterspec_instruction'),
    (r'(?i)^Flexible elastomeric and polyolefin thicknesses are limited.*$', 'masterspec_instruction'),
    (r'(?i)^To comply with ASHRAE.*insulation should have.*$', 'masterspec_instruction'),
    (r'(?i)^Architect should be prepared to reject.*$', 'masterspec_instruction'),
    (r'(?i)^Retain [^\n]*remaining after.*$', 'masterspec_instruction'),
    (r'(?i)^Paragraph below is defined in Section.*$', 'masterspec_instruction'),
    (r'(?i)^.*requires calculating and detailing at each use.*$', 'masterspec_instruction'),
    (r'(?i)^Verify suitability of.*$', 'masterspec_instruction'),
    (r'(?i)^Specify parts in [^\n]*subparagraph.*$', 'masterspec_instruction'),
    (r'(?i)^Specify parts in first.*$', 'masterspec_instruction'),
    (r'(?i)^Option:\s+[^\n]*may be used.*$', 'masterspec_instruction'),
    (r"(?i)^.*catalogs indicate.*$", 'masterspec_instruction'),
    (r'(?i)^High-compressive-strength inserts may permit.*$', 'masterspec_instruction'),

    # Copyright notices
    (r'(?i)^Copyright\s*©?\s*\d{4}.*$', 'copyright'),
    (r'(?i)^©\s*\d{4}.*$', 'copyright'),
    (r'(?i)^Exclusively published and distributed by.*$', 'copyright'),
    (r'(?i)all rights reserved.*$', 'copyright'),
    (r'(?i)proprietary\s+information.*$', 'copyright'),

    # Separator lines
    (r'^[\*]{4,}\s*$', 'separator'),
    (r'^[-]{4,}\s*$', 'separator'),
    (r'^[=]{4,}\s*$', 'separator'),

    # Page artifacts
    (r'(?i)^page\s+\d+\s*(?:of\s*\d+)?\s*$', 'page_number'),

    # Revision marks
    (r'(?i)\{revision[^\}]*\}', 'revision_mark'),

    # Hidden text markers
    (r'(?i)<<[^>]*hidden[^>]*>>', 'hidden_text'),
]

# Pre-compile for speed and to avoid repeated regex compilation
_BOILERPLATE_RX = [(re.compile(pat, flags=re.MULTILINE), tag) for pat, tag in BOILERPLATE_PATTERNS]

_PART_RX = re.compile(r"^\s*PART\s+[123]\b", re.IGNORECASE)
_ARTICLE_RX = re.compile(r"^\s*\d{1,2}\.\d{1,3}\b")
_SECTION_ID_RX = re.compile(
    r"^\s*SECTION\s+(?:\d{6,}|\d{2}(?:[ \t\u00a0]+\d{2}){2,})\b",
    re.IGNORECASE,
)
_END_OF_SECTION_RX = re.compile(r"^\s*END\s+OF\s+SECTION\s*", re.IGNORECASE)
_ALL_CAPS_RX = re.compile(r"^[^a-z]*[A-Z][^a-z]*$")
_EDITORIAL_COMMENT_STYLE_IDS = frozenset({"CMT"})
LLM_IGNORED_REASON = "non_csi_content"
_FORMAT_ONLY_PROTECTED_PPR_PROPERTIES = frozenset({
    "pStyle",
    "numPr",
    "sectPr",
    "pPrChange",
    "cnfStyle",
})
_MARKER_RX = [
    (re.compile(r"^\s*[A-Z]\.\s+"), "upper_alpha"),
    (re.compile(r"^\s*\d+\.\s+"), "number"),
    (re.compile(r"^\s*[a-z]\.\s+"), "lower_alpha"),
    (re.compile(r"^\s*\d+\)\s+"), "deep_level_5"),
    (re.compile(r"^\s*[a-z]\)\s+"), "deep_level_6"),
    (re.compile(r"^\s*\(\d+\)\s+"), "deep_level_7"),
    (re.compile(r"^\s*\([a-z]\)\s+"), "deep_level_8"),
]


def _match_section_header(text: str) -> Optional[re.Match[str]]:
    """Match a section header without consuming sentence-form cross-references."""

    match = _SECTION_ID_RX.match(text)
    if match is None:
        return None
    remainder = text[match.end():].strip(
        " \t\u00a0-\u2010\u2011\u2012\u2013\u2014\u2015:"
    )
    if remainder and not _ALL_CAPS_RX.fullmatch(remainder):
        return None
    return match


def _table_ranges(document_xml_text: str) -> List[Tuple[int, int]]:
    return [
        (start, end)
        for start, end, _block in iter_element_xml_blocks(document_xml_text, "w:tbl")
    ]


def _in_any_range(pos: int, ranges: List[Tuple[int, int]]) -> bool:
    for start, end in ranges:
        if start <= pos < end:
            return True
    return False


def _extract_rpr_hints(p_xml: str) -> Dict[str, Any]:
    def _ooxml_on_off_state(tag: str) -> Optional[bool]:
        m = re.search(rf"<w:{tag}\b([^>]*)/?>", p_xml)
        if not m:
            return None
        attrs = m.group(1) or ""
        vm = re.search(r'w:val\s*=\s*"([^"]+)"', attrs)
        if not vm:
            return True
        val = vm.group(1).strip().lower()
        if val in {"false", "0", "off", "none"}:
            return False
        return True

    hints: Dict[str, Any] = {}
    bold = _ooxml_on_off_state("b")
    italic = _ooxml_on_off_state("i")
    underline = _ooxml_on_off_state("u")
    if bold is not None:
        hints["bold"] = bold
    if italic is not None:
        hints["italic"] = italic
    if underline is not None:
        hints["underline"] = underline
    return hints


def _detect_marker_type(text: str, numpr: Dict[str, Optional[str]]) -> Optional[str]:
    for rx, marker_type in _MARKER_RX:
        if rx.match(text):
            return marker_type
    if numpr.get("numId") and numpr.get("numId") != "0":
        return "automatic_numbering"
    return None


def _contains_drawing_or_textbox(paragraph_xml: str) -> bool:
    return any(
        next(iter_element_xml_blocks(paragraph_xml, name), None) is not None
        for name in (
            "w:drawing",
            "w:pict",
            "w:object",
            "w:txbxContent",
            "v:textbox",
            "wps:txbx",
        )
    )


def _effective_numpr(
    paragraph_xml: str,
    styles_xml_text: str,
) -> Optional[Dict[str, str]]:
    """Resolve direct and style-inherited numbering for one paragraph."""
    inherited: Dict[str, Optional[str]] = {"numId": None, "ilvl": None}
    style_id = paragraph_pstyle_from_block(paragraph_xml)
    if style_id:
        inherited_block = _find_style_numpr_in_chain(styles_xml_text, style_id)
        if inherited_block:
            inherited = paragraph_numpr_from_block(inherited_block)

    direct = paragraph_numpr_from_block(paragraph_xml)
    effective: Dict[str, str] = {
        key: value
        for key, value in inherited.items()
        if isinstance(value, str) and value
    }
    effective.update(
        {
            key: value
            for key, value in direct.items()
            if isinstance(value, str) and value
        }
    )
    if not effective.get("numId") or effective.get("numId") == "0":
        return None
    effective.setdefault("ilvl", "0")
    return effective


def _wq(local_name: str) -> str:
    return f"{{http://schemas.openxmlformats.org/wordprocessingml/2006/main}}{local_name}"


def _wval(node: Optional[ET.Element], attribute: str = "val") -> Optional[str]:
    if node is None:
        return None
    return node.attrib.get(_wq(attribute))


def _w_on_off_value(node: Optional[ET.Element]) -> Optional[str]:
    """Return canonical Word on/off semantics, including an empty true tag."""

    if node is None:
        return None
    value = _wval(node)
    if value is None:
        return "1"
    normalized = value.strip().lower()
    if normalized in {"0", "false", "off", "none"}:
        return "0"
    if normalized in {"1", "true", "on"}:
        return "1"
    return value


def _build_numbering_catalog(numbering_xml_text: str) -> Dict[str, Any]:
    """Parse just enough numbering.xml to resolve a rendered list signature."""
    if not numbering_xml_text.strip():
        return {"nums": {}, "abstracts": {}}
    root = ET.fromstring(numbering_xml_text)
    abstracts: Dict[str, Dict[str, Any]] = {}
    nums: Dict[str, Dict[str, Any]] = {}

    for abstract in root.findall(f".//{_wq('abstractNum')}"):
        abstract_id = abstract.attrib.get(_wq("abstractNumId"))
        if not abstract_id:
            continue
        levels: Dict[str, Dict[str, str]] = {}
        for level in abstract.findall(_wq("lvl")):
            ilvl = level.attrib.get(_wq("ilvl"))
            if ilvl is None:
                continue
            item: Dict[str, str] = {"ilvl": ilvl}
            for element_name, key in (
                ("start", "start"),
                ("numFmt", "numFmt"),
                ("lvlRestart", "lvlRestart"),
                ("pStyle", "pStyle"),
                ("lvlText", "lvlText"),
                ("suff", "suff"),
                ("isLgl", "isLgl"),
            ):
                element = level.find(_wq(element_name))
                value = (
                    _w_on_off_value(element)
                    if element_name == "isLgl"
                    else _wval(element)
                )
                if value is not None:
                    item[key] = value
            levels[ilvl] = item
        abstracts[abstract_id] = {"levels": levels}

    for num in root.findall(f".//{_wq('num')}"):
        num_id = num.attrib.get(_wq("numId"))
        if not num_id:
            continue
        item: Dict[str, Any] = {
            "abstractNumId": _wval(num.find(_wq("abstractNumId"))),
            "overrides": {},
        }
        for override in num.findall(_wq("lvlOverride")):
            ilvl = override.attrib.get(_wq("ilvl"))
            if ilvl is None:
                continue
            override_item: Dict[str, str] = {}
            start = _wval(override.find(_wq("startOverride")))
            if start is not None:
                override_item["startOverride"] = start
            override_level = override.find(_wq("lvl"))
            if override_level is not None:
                for element_name, key in (
                    ("start", "start"),
                    ("numFmt", "numFmt"),
                    ("lvlRestart", "lvlRestart"),
                    ("pStyle", "pStyle"),
                    ("lvlText", "lvlText"),
                    ("suff", "suff"),
                    ("isLgl", "isLgl"),
                ):
                    element = override_level.find(_wq(element_name))
                    value = (
                        _w_on_off_value(element)
                        if element_name == "isLgl"
                        else _wval(element)
                    )
                    if value is not None:
                        override_item[key] = value
            item["overrides"][ilvl] = override_item
        nums[num_id] = item

    return {"nums": nums, "abstracts": abstracts}


def _resolve_numbering_pattern(
    effective_numpr: Optional[Dict[str, str]],
    numbering_catalog: Dict[str, Any],
) -> Optional[Dict[str, str]]:
    if not effective_numpr:
        return None
    num_id = effective_numpr.get("numId")
    ilvl = effective_numpr.get("ilvl", "0")
    if not num_id or num_id == "0":
        return None

    pattern: Dict[str, str] = {"numId": num_id, "ilvl": ilvl}
    num = numbering_catalog.get("nums", {}).get(num_id, {})
    abstract_id = num.get("abstractNumId") if isinstance(num, dict) else None
    if isinstance(abstract_id, str) and abstract_id:
        pattern["abstractNumId"] = abstract_id
        level = (
            numbering_catalog.get("abstracts", {})
            .get(abstract_id, {})
            .get("levels", {})
            .get(ilvl, {})
        )
        if isinstance(level, dict):
            for key in (
                "start", "numFmt", "lvlRestart", "pStyle",
                "lvlText", "suff", "isLgl",
            ):
                if isinstance(level.get(key), str):
                    pattern[key] = level[key]

    override = num.get("overrides", {}).get(ilvl, {}) if isinstance(num, dict) else {}
    if isinstance(override, dict):
        for key in (
            "start", "numFmt", "lvlRestart", "pStyle",
            "lvlText", "suff", "isLgl", "startOverride",
        ):
            if isinstance(override.get(key), str):
                pattern[key] = override[key]
    return pattern


def _numbering_role_candidates(
    target_pattern: Optional[Dict[str, str]],
    role_specs: Optional[Dict[str, Dict[str, Any]]],
    available_roles: List[str],
) -> List[str]:
    """Return roles whose portable rendered numbering signature is exact.

    ``numId``, ``abstractNumId``, and the level's ``pStyle`` reference are
    package-local identifiers, so matching those across the architect and
    target packages would be incorrect.  A deterministic match requires both
    rendered signals (format and level text), plus every portable counter,
    restart, and marker constraint parsed from numbering.xml.
    """
    if not target_pattern or not role_specs:
        return []
    if not target_pattern.get("numFmt") or target_pattern.get("lvlText") is None:
        return []

    matches: List[str] = []
    for role in available_roles:
        spec = role_specs.get(role)
        if not isinstance(spec, dict):
            continue
        if spec.get("numbering_provenance") not in {"style_numpr", "direct_numpr"}:
            continue
        expected = spec.get("numbering_pattern")
        if not isinstance(expected, dict):
            continue
        if not expected.get("numFmt") or expected.get("lvlText") is None:
            continue

        # This is an exact portable signature.  Presence is part of the
        # comparison: an omitted restart/start/suffix property is not the same
        # contract as an explicitly supplied one, even if the common format
        # and level-text fields match.
        keys = (
            "ilvl",
            "start",
            "numFmt",
            "lvlRestart",
            "lvlText",
            "suff",
            "isLgl",
            "startOverride",
        )
        if all(
            _numbering_pattern_entry(target_pattern, key)
            == _numbering_pattern_entry(expected, key)
            for key in keys
        ):
            matches.append(role)
    return matches


def _numbering_pattern_entry(
    pattern: Dict[str, Any],
    key: str,
) -> Tuple[bool, Optional[str]]:
    """Return a value together with explicit presence information."""

    if key not in pattern:
        return False, None
    value = pattern[key]
    return True, None if value is None else str(value)


def _numbering_role_counter_conflicts(
    target_pattern: Optional[Dict[str, str]],
    role_specs: Optional[Dict[str, Dict[str, Any]]],
) -> Set[str]:
    """Return roles with the same marker but contradictory list semantics.

    Contextual target-only inference is useful when the architect deliberately
    uses another hierarchy (for example Canadian numeric markers versus CSI
    letters).  It must not override an architect role whose marker otherwise
    matches but whose start/restart/suffix/legal-numbering contract differs.
    Such a paragraph is intentionally left unresolved instead of being mapped
    from an unsafe near-match.
    """

    if not target_pattern or not role_specs:
        return set()
    marker_keys = ("ilvl", "numFmt", "lvlText")
    counter_keys = ("start", "lvlRestart", "suff", "isLgl", "startOverride")
    conflicts: Set[str] = set()
    for role, spec in role_specs.items():
        if not isinstance(spec, dict):
            continue
        if spec.get("numbering_provenance") not in {"style_numpr", "direct_numpr"}:
            continue
        expected = spec.get("numbering_pattern")
        if not isinstance(expected, dict):
            continue
        if not all(
            _numbering_pattern_entry(target_pattern, key)
            == _numbering_pattern_entry(expected, key)
            for key in marker_keys
        ):
            continue
        if any(
            _numbering_pattern_entry(target_pattern, key)
            != _numbering_pattern_entry(expected, key)
            for key in counter_keys
        ):
            conflicts.add(role)
    return conflicts


def _resolve_role(preferred: str, available_roles: List[str]) -> Optional[str]:
    special_fallbacks = {
        "SectionID": ["SectionID", "SectionTitle"],
        "END_OF_SECTION": ["END_OF_SECTION"],
        "SectionTitle": ["SectionTitle"],
    }
    fallback_chain = special_fallbacks.get(
        preferred,
        ROLE_FALLBACKS.get(preferred, (preferred,)),
    )
    for candidate in fallback_chain:
        if candidate in available_roles:
            return candidate
    return None


def _deterministic_role_for_paragraph(paragraph: Dict[str, Any], prev_text: str = "") -> Optional[str]:
    text = paragraph.get("text", "")
    if not text or paragraph.get("in_table"):
        return None

    # Word numbering is stronger evidence than visible text.  In automatically
    # numbered documents the rendered marker is not present in ``w:t`` at all:
    # a PART can therefore arrive here as just ``GENERAL``, and an automatically
    # numbered list item can begin with a sentence-form ``Section ...`` cross-
    # reference.  Resolve (or deliberately leave unresolved) that numbering
    # before applying text-only heading heuristics.
    if paragraph.get("effective_numPr"):
        numbering_role = paragraph.get("numbering_role")
        if isinstance(numbering_role, str):
            return numbering_role
        if paragraph.get("numbering_semantic_conflict") is True:
            return None
        pattern = paragraph.get("numbering_pattern")
        if isinstance(pattern, dict):
            return role_from_numbering_signature(
                pattern.get("numFmt"),
                pattern.get("lvlText"),
                pattern.get("ilvl"),
            )
        return None

    if _match_section_header(text):
        return "SectionID"
    if _END_OF_SECTION_RX.match(text):
        return "END_OF_SECTION"
    if _PART_RX.match(text):
        return "PART"
    if _ARTICLE_RX.match(text):
        return "ARTICLE"
    if prev_text and _match_section_header(prev_text) and _ALL_CAPS_RX.match(text):
        return "SectionTitle"

    marker_type = paragraph.get("marker_type")
    if marker_type == "upper_alpha":
        return "PARAGRAPH"
    if marker_type == "number":
        return "SUBPARAGRAPH"
    if marker_type == "lower_alpha":
        return "SUBSUBPARAGRAPH"
    if marker_type == "deep_level_5":
        return "SUBPARAGRAPH_LEVEL_5"
    if marker_type == "deep_level_6":
        return "SUBPARAGRAPH_LEVEL_6"
    if marker_type == "deep_level_7":
        return "SUBPARAGRAPH_LEVEL_7"
    if marker_type == "deep_level_8":
        return "SUBPARAGRAPH_LEVEL_8"
    return None


def preclassify_paragraphs(paragraphs: List[Dict[str, Any]], available_roles: List[str]) -> Dict[int, str]:
    out: Dict[int, str] = {}
    prev_text = ""
    for paragraph in paragraphs:
        preferred = _deterministic_role_for_paragraph(paragraph, prev_text=prev_text)
        resolved = _resolve_role(preferred, available_roles) if preferred else None
        if resolved:
            out[paragraph["paragraph_index"]] = resolved
        prev_text = paragraph.get("text", "")
    return out


def strip_boilerplate_with_report(content: str) -> tuple:
    """
    Strip boilerplate from a paragraph string and return (cleaned_text, matched_tags).
    Placeholders are NOT stripped here (your patterns do not remove generic [ ... ] placeholders).
    """
    cleaned = content
    hits: list = []

    for rx, tag in _BOILERPLATE_RX:
        if rx.search(cleaned):
            hits.append(tag)
            cleaned = rx.sub('', cleaned)

    # Clean up whitespace
    cleaned = re.sub(r'\n{3,}', '\n\n', cleaned)
    cleaned = re.sub(r'[ \t]+\n', '\n', cleaned)
    cleaned = cleaned.strip()

    # Deduplicate tags (stable order)
    if hits:
        seen = set()
        hits = [t for t in hits if not (t in seen or seen.add(t))]

    return cleaned, hits


def build_phase2_slim_bundle(
    extract_dir: Path,
    available_roles: Optional[List[str]] = None,
    role_specs: Optional[Dict[str, Dict[str, Any]]] = None,
) -> Dict[str, Any]:
    """
    Build the slim bundle for Phase 2 LLM classification.

    Args:
        extract_dir: Path to extracted DOCX folder
        available_roles: List of role names available in the architect template.
                        If None, all standard roles are allowed.
        role_specs: Strict Phase 1 role contracts.  When supplied, their
                    portable numbering patterns enable deterministic matching.

    Returns:
        Dict containing available_roles, filter_report, and paragraphs
    """
    doc_path = extract_dir / "word" / "document.xml"
    doc_text = read_xml_text(doc_path)

    paragraphs = []
    deterministic_ignored: List[Dict[str, Any]] = []
    filter_report = {
        "paragraphs_removed_entirely": [],
        "paragraphs_stripped": [],
        "paragraphs_out_of_scope": [],
    }

    table_ranges = _table_ranges(doc_text)
    raw_paragraphs = list(iter_paragraph_xml_blocks(doc_text))
    styles_path = extract_dir / "word" / "styles.xml"
    styles_text = read_xml_text(styles_path) if styles_path.is_file() else ""
    numbering_path = extract_dir / "word" / "numbering.xml"
    numbering_text = read_xml_text(numbering_path) if numbering_path.is_file() else ""
    numbering_catalog = _build_numbering_catalog(numbering_text)

    # Default roles if none specified.
    if available_roles is None:
        available_roles = [
            "SectionID",
            "SectionTitle",
            *BODY_HIERARCHY_ROLES,
        ]

    skip_next_section_title = False
    for idx, (start, _e, p_xml) in enumerate(raw_paragraphs):
        raw_text = paragraph_text_from_block(p_xml)
        in_table = _in_any_range(start, table_ranges)
        if in_table:
            filter_report["paragraphs_out_of_scope"].append({
                "paragraph_index": idx,
                "reason": "table",
                "original_text_preview": raw_text[:120],
            })
            continue
        contains_drawing_or_textbox = _contains_drawing_or_textbox(p_xml)
        analysis_xml = strip_out_of_scope_subtrees(p_xml)
        pstyle = paragraph_pstyle_from_block(analysis_xml)
        if contains_drawing_or_textbox:
            filter_report["paragraphs_out_of_scope"].append({
                "paragraph_index": idx,
                "reason": "drawing_or_textbox_subtree",
                "original_text_preview": raw_text[:120],
            })
            # The host paragraph is the ownership boundary.  Styling its
            # visible runs while merely protecting the nested drawing/textbox
            # would make one index both classified and out-of-scope and would
            # violate the promised byte stability for these structures.
            continue

        if not raw_text:
            if paragraph_contains_sectpr(analysis_xml):
                filter_report["paragraphs_removed_entirely"].append({
                    "paragraph_index": idx,
                    "tags": ["empty_structural_section_break"],
                    "original_text_preview": "",
                })
            continue

        cleaned_text, tags = strip_boilerplate_with_report(raw_text)

        if not cleaned_text:
            if tags:
                filter_report["paragraphs_removed_entirely"].append({
                    "paragraph_index": idx,
                    "tags": tags,
                    "original_text_preview": raw_text[:120]
                })
                deterministic_ignored.append({
                    "paragraph_index": idx,
                    "reason": "boilerplate",
                })
            continue

        if tags:
            filter_report["paragraphs_stripped"].append({
                "paragraph_index": idx,
                "tags": tags
            })

        # MasterSpec uses the dedicated CMT paragraph style for editorial
        # instructions that are intentionally preserved in the document but
        # must not receive an architect content style.  Text-pattern filtering
        # catches many of these notes, but not open-ended guidance such as
        # "Always retain..." or explanatory paragraphs containing URLs.
        if pstyle in _EDITORIAL_COMMENT_STYLE_IDS:
            filter_report["paragraphs_removed_entirely"].append({
                "paragraph_index": idx,
                "tags": ["editorial_comment_style"],
                "original_text_preview": raw_text[:120],
            })
            deterministic_ignored.append({
                "paragraph_index": idx,
                "reason": "editorial_comment_style",
            })
            continue

        # Resolve automatic numbering before any text-only section-header
        # filtering.  A rendered PART marker is absent from ``w:t`` (leaving
        # text such as GENERAL), and a numbered requirement can itself begin
        # with an uppercase SECTION cross-reference.
        numpr = paragraph_numpr_from_block(analysis_xml)
        effective_numpr = _effective_numpr(analysis_xml, styles_text)
        numbering_pattern = _resolve_numbering_pattern(
            effective_numpr,
            numbering_catalog,
        )
        numbering_candidates = _numbering_role_candidates(
            numbering_pattern,
            role_specs,
            available_roles,
        )
        numbering_counter_conflicts = _numbering_role_counter_conflicts(
            numbering_pattern,
            role_specs,
        )
        contextual_numbering_role = None
        contextual_numbering_conflict = False
        if isinstance(effective_numpr, dict):
            contextual_numbering_role = role_from_numbering_catalog(
                numbering_catalog,
                effective_numpr.get("numId"),
                effective_numpr.get("ilvl", "0"),
            )
            contextual_numbering_conflict = (
                contextual_numbering_role in numbering_counter_conflicts
            )
            if (
                contextual_numbering_role not in available_roles
                or contextual_numbering_conflict
            ):
                contextual_numbering_role = None

        # A source can contain a section-number/title header even when the
        # architect template has no semantic style for it.  Such a header has
        # no valid parent-role fallback: classifying it as PART would cause a
        # numbered Canadian PART style to insert a spurious list item.  Keep
        # the original paragraph untouched and outside classification instead.
        # A real combined section header is conventionally uppercase.
        # Sentence-form cross-references such as
        # `Section 012100 "Allowances" for ...` are numbered content and must
        # remain in the classification bundle.
        section_match = (
            None
            if effective_numpr is not None
            else _match_section_header(cleaned_text)
        )
        if skip_next_section_title:
            skip_next_section_title = False
            has_csi_marker = bool(
                effective_numpr is not None
                or _PART_RX.match(cleaned_text)
                or _ARTICLE_RX.match(cleaned_text)
                or _END_OF_SECTION_RX.match(cleaned_text)
                or any(pattern.match(cleaned_text) for pattern, _kind in _MARKER_RX)
            )
            if _ALL_CAPS_RX.fullmatch(cleaned_text) and not has_csi_marker:
                filter_report["paragraphs_removed_entirely"].append({
                    "paragraph_index": idx,
                    "tags": ["section_title_no_role"],
                    "original_text_preview": raw_text[:120],
                })
                deterministic_ignored.append({
                    "paragraph_index": idx,
                    "reason": "section_title_no_role",
                })
                continue
        if section_match is not None:
            remainder = cleaned_text[section_match.end():].strip(
                " \t\u00a0-\u2010\u2011\u2012\u2013\u2014\u2015:"
            )
            skip_next_section_title = not remainder and "SectionTitle" not in available_roles
            if "SectionID" not in available_roles and "SectionTitle" not in available_roles:
                filter_report["paragraphs_removed_entirely"].append({
                    "paragraph_index": idx,
                    "tags": ["section_header_no_role"],
                    "original_text_preview": raw_text[:120],
                })
                deterministic_ignored.append({
                    "paragraph_index": idx,
                    "reason": "section_header_no_role",
                })
                continue

        # Skip END OF SECTION lines when no available role can receive them.
        # These are deterministically identifiable but unstyled — sending them
        # to the LLM (whose prompt says to skip them) causes coverage failures.
        if _END_OF_SECTION_RX.match(cleaned_text):
            if available_roles is None or "END_OF_SECTION" not in available_roles:
                filter_report["paragraphs_removed_entirely"].append({
                    "paragraph_index": idx,
                    "tags": ["end_of_section_no_role"],
                    "original_text_preview": raw_text[:120]
                })
                deterministic_ignored.append({
                    "paragraph_index": idx,
                    "reason": "end_of_section_no_role",
                })
                continue

        ppr_hints = paragraph_ppr_hints_from_block(analysis_xml)
        rpr_hints = _extract_rpr_hints(analysis_xml)
        marker_type = _detect_marker_type(cleaned_text, effective_numpr or numpr)

        paragraphs.append({
            "paragraph_index": idx,
            "text": cleaned_text[:200],
            "prev_text": paragraphs[-1]["text"][:80] if paragraphs else "",
            "next_text": "",
            "pStyle": pstyle,
            "pPr_hints": ppr_hints,
            "rPr_hints": rpr_hints,
            "in_table": in_table,
            "marker_type": marker_type,
            "numPr": numpr if (numpr.get("numId") or numpr.get("ilvl")) else None,
            "effective_numPr": effective_numpr,
            "numbering_pattern": numbering_pattern,
            "numbering_match_candidates": numbering_candidates,
            "numbering_semantic_conflict": (
                contextual_numbering_conflict
                and len(numbering_candidates) != 1
            ),
            "numbering_role": (
                numbering_candidates[0]
                if len(numbering_candidates) == 1
                else contextual_numbering_role
            ),
            "contains_sectPr": paragraph_contains_sectpr(analysis_xml),
        })

    for i in range(len(paragraphs) - 1):
        paragraphs[i]["next_text"] = paragraphs[i + 1]["text"][:80]

    deterministic = preclassify_paragraphs(paragraphs, available_roles)
    unresolved_paragraphs = [p for p in paragraphs if not p.get("in_table") and p["paragraph_index"] not in deterministic]

    return {
        "available_roles": available_roles,
        "filter_report": filter_report,
        "paragraphs": unresolved_paragraphs,
        "deterministic_classifications": [
            {"paragraph_index": idx, "csi_role": role}
            for idx, role in sorted(deterministic.items())
        ],
        "deterministic_ignored_paragraphs": sorted(
            deterministic_ignored,
            key=lambda item: item["paragraph_index"],
        ),
    }


def _validate_payload_shape(
    classifications: Dict[str, Any],
    allowed_roles: List[str],
    allowed_indices: Set[int],
) -> Dict[int, str]:
    if not isinstance(classifications, dict):
        raise ValueError("classifications payload must be an object")
    items = classifications.get("classifications")
    if not isinstance(items, list):
        raise ValueError("classifications payload missing classifications list")

    allowed = set(allowed_roles)
    seen: Dict[int, str] = {}
    for item in items:
        if not isinstance(item, dict):
            raise ValueError("all classification entries must be objects")
        idx = item.get("paragraph_index")
        role = item.get("csi_role")
        if not isinstance(idx, int):
            raise ValueError(f"invalid paragraph_index: {idx!r}")
        if idx in seen:
            raise ValueError(f"duplicate classification for paragraph_index={idx}")
        if idx not in allowed_indices:
            raise ValueError(f"classification index not classifiable: {idx}")
        if role not in allowed:
            raise ValueError(f"invalid csi_role for paragraph_index={idx}: {role!r}")
        seen[idx] = role

    return seen


def _validate_ignored_shape(
    classifications: Dict[str, Any],
    allowed_indices: Set[int],
    classified_indices: Set[int],
) -> Dict[int, str]:
    """Validate explicit ignore dispositions and prove they are disjoint."""

    raw_items = classifications.get("ignored_paragraphs", [])
    if not isinstance(raw_items, list):
        raise ValueError("classifications payload ignored_paragraphs must be a list")

    seen: Dict[int, str] = {}
    for item in raw_items:
        if not isinstance(item, dict):
            raise ValueError("all ignored paragraph entries must be objects")
        idx = item.get("paragraph_index")
        reason = item.get("reason")
        if not isinstance(idx, int):
            raise ValueError(f"invalid ignored paragraph_index: {idx!r}")
        if idx in seen:
            raise ValueError(f"duplicate ignored disposition for paragraph_index={idx}")
        if idx in classified_indices:
            raise ValueError(
                f"paragraph_index={idx} cannot be both classified and ignored"
            )
        if idx not in allowed_indices:
            raise ValueError(f"ignored index not classifiable: {idx}")
        if not isinstance(reason, str) or not reason.strip():
            raise ValueError(
                f"invalid ignored reason for paragraph_index={idx}: {reason!r}"
            )
        seen[idx] = reason.strip()
    return seen


def _validate_disposition_payload(
    classifications: Dict[str, Any],
    allowed_roles: List[str],
    allowed_indices: Set[int],
) -> Tuple[Dict[int, str], Dict[int, str]]:
    classified = _validate_payload_shape(
        classifications,
        allowed_roles,
        allowed_indices,
    )
    ignored = _validate_ignored_shape(
        classifications,
        allowed_indices,
        set(classified),
    )
    return classified, ignored


def _bundle_unresolved_indices(bundle: Dict[str, Any]) -> Set[int]:
    raw = bundle.get("paragraphs", [])
    if not isinstance(raw, list):
        raise ValueError("bundle paragraphs must be a list")
    indices: Set[int] = set()
    for item in raw:
        if not isinstance(item, dict) or not isinstance(
            item.get("paragraph_index"),
            int,
        ):
            raise ValueError("bundle contains an invalid unresolved paragraph")
        idx = item["paragraph_index"]
        if idx in indices:
            raise ValueError(f"duplicate unresolved paragraph_index={idx}")
        indices.add(idx)
    return indices


def _bundle_deterministic_dispositions(
    bundle: Dict[str, Any],
    allowed_roles: List[str],
) -> Tuple[Dict[int, str], Dict[int, str]]:
    deterministic: Dict[int, str] = {}
    allowed = set(allowed_roles)
    for item in bundle.get("deterministic_classifications", []):
        if not isinstance(item, dict):
            raise ValueError("bundle contains an invalid deterministic classification")
        idx = item.get("paragraph_index")
        role = item.get("csi_role")
        if not isinstance(idx, int) or role not in allowed:
            raise ValueError("bundle contains an invalid deterministic classification")
        if idx in deterministic:
            raise ValueError(f"duplicate deterministic classification for paragraph_index={idx}")
        deterministic[idx] = role

    deterministic_ignored: Dict[int, str] = {}
    for item in bundle.get("deterministic_ignored_paragraphs", []):
        if not isinstance(item, dict):
            raise ValueError("bundle contains an invalid deterministic ignored disposition")
        idx = item.get("paragraph_index")
        reason = item.get("reason")
        if not isinstance(idx, int) or not isinstance(reason, str) or not reason.strip():
            raise ValueError("bundle contains an invalid deterministic ignored disposition")
        if idx in deterministic_ignored:
            raise ValueError(f"duplicate deterministic ignored disposition for paragraph_index={idx}")
        deterministic_ignored[idx] = reason.strip()

    overlap = set(deterministic) & set(deterministic_ignored)
    if overlap:
        raise ValueError(
            "bundle has conflicting deterministic dispositions for paragraph "
            f"indices: {sorted(overlap)[:20]}"
        )
    return deterministic, deterministic_ignored


def validate_phase2_llm_payload(bundle: Dict[str, Any], classifications: Dict[str, Any], allowed_roles: List[str]) -> None:
    unresolved = _bundle_unresolved_indices(bundle)
    classified, ignored = _validate_disposition_payload(
        classifications,
        allowed_roles,
        unresolved,
    )
    missing = sorted(unresolved - set(classified) - set(ignored))
    if missing:
        raise ValueError(f"missing coverage for paragraph indices: {missing[:20]}")


def coerce_to_final_classifications(
    bundle: Dict[str, Any],
    classifications: Dict[str, Any],
    allowed_roles: List[str],
) -> Dict[str, Any]:
    unresolved = _bundle_unresolved_indices(bundle)
    deterministic, deterministic_ignored = _bundle_deterministic_dispositions(
        bundle,
        allowed_roles,
    )
    deterministic_indices = set(deterministic) | set(deterministic_ignored)
    bundle_overlap = deterministic_indices & unresolved
    if bundle_overlap:
        raise ValueError(
            "bundle paragraph appears in both unresolved and deterministic "
            f"dispositions: {sorted(bundle_overlap)[:20]}"
        )
    total = set(deterministic) | set(deterministic_ignored) | unresolved

    incoming, incoming_ignored = _validate_disposition_payload(
        classifications,
        allowed_roles,
        total,
    )
    incoming_indices = set(incoming) | set(incoming_ignored)

    if incoming_indices <= unresolved:
        missing_unresolved = sorted(unresolved - incoming_indices)
        if missing_unresolved:
            raise ValueError(f"missing coverage for paragraph indices: {missing_unresolved[:20]}")
        merged = dict(deterministic)
        merged.update(incoming)
        merged_ignored = dict(deterministic_ignored)
        merged_ignored.update(incoming_ignored)
    elif incoming_indices == total:
        for idx, expected in deterministic.items():
            actual = incoming.get(idx)
            if actual != expected:
                raise ValueError(
                    f"deterministic override attempted at paragraph_index={idx}: "
                    f"expected {expected!r}, got {actual!r}"
                )
        for idx, expected in deterministic_ignored.items():
            actual = incoming_ignored.get(idx)
            if actual != expected:
                raise ValueError(
                    f"deterministic ignored override attempted at paragraph_index={idx}: "
                    f"expected {expected!r}, got {actual!r}"
                )
        merged = incoming
        merged_ignored = incoming_ignored
    else:
        raise ValueError("payload is neither unresolved-only nor valid final coverage")

    notes = classifications.get("notes", []) if isinstance(classifications, dict) else []
    return {
        "classifications": [
            {"paragraph_index": idx, "csi_role": role}
            for idx, role in sorted(merged.items())
        ],
        "ignored_paragraphs": [
            {"paragraph_index": idx, "reason": reason}
            for idx, reason in sorted(merged_ignored.items())
        ],
        "notes": notes if isinstance(notes, list) else [],
    }


def validate_phase2_final_payload(bundle: Dict[str, Any], classifications: Dict[str, Any], allowed_roles: List[str]) -> None:
    coerce_to_final_classifications(bundle, classifications, allowed_roles)


def validate_phase2_classification_contract(bundle: Dict[str, Any], classifications: Dict[str, Any], allowed_roles: List[str]) -> None:
    # Backward-compatible alias: strict unresolved-only LLM payload validation.
    validate_phase2_llm_payload(bundle, classifications, allowed_roles)


def _protect_named_subtrees(
    xml_text: str,
    qualified_names: Tuple[str, ...],
) -> Tuple[str, List[Tuple[str, str]]]:
    ranges: List[Tuple[int, int]] = []
    for name in qualified_names:
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
    while f"__PHASE2_CONTRACT_{nonce}_" in xml_text:
        nonce += 1
    pieces: List[str] = []
    replacements: List[Tuple[str, str]] = []
    last = 0
    for index, (start, end) in enumerate(merged):
        token = f"__PHASE2_CONTRACT_{nonce}_{index}__"
        pieces.extend((xml_text[last:start], token))
        replacements.append((token, xml_text[start:end]))
        last = end
    pieces.append(xml_text[last:])
    return "".join(pieces), replacements


def _restore_named_subtrees(
    xml_text: str,
    replacements: List[Tuple[str, str]],
) -> str:
    restored = xml_text
    for token, subtree in replacements:
        if restored.count(token) != 1:
            raise ValueError("Protected paragraph subtree changed during edit")
        restored = restored.replace(token, subtree, 1)
    return restored


def _normalize_paragraph_for_contract_unprotected(
    p_xml: str,
    allowed_ppr_properties: Optional[Set[str]] = None,
    allowed_rpr_properties: Optional[Set[str]] = None,
) -> str:
    """
    Normalize paragraph for contract comparison.
    Strips elements we're allowed to change: pStyle, numPr, and
    selected run-level formatting.  Also removes
    empty pPr / rPr shells so paragraphs that originally lacked
    these blocks compare equal after stripping.
    """
    out, protected_subtrees = _protect_named_subtrees(
        p_xml,
        ("w:pPrChange", "w:sectPr"),
    )
    # Strip pStyle (we change this)
    out = re.sub(r"<w:pStyle\b[^>]*/>", "", out)
    properties = (
        {"numPr", "jc", "ind", "spacing"}
        if allowed_ppr_properties is None
        else set(allowed_ppr_properties)
    )
    for name in sorted(properties):
        if not re.fullmatch(r"[A-Za-z_][\w.-]*", name):
            raise ValueError(f"Invalid contract pPr property name: {name!r}")
        out = re.sub(rf"<w:{name}\b[^>]*/>", "", out)
        out = re.sub(
            rf"<w:{name}\b[^>]*>[\s\S]*?</w:{name}>",
            "",
            out,
            flags=re.S,
        )
    run_properties = (
        {"rFonts", "sz", "szCs"}
        if allowed_rpr_properties is None
        else set(allowed_rpr_properties)
    )
    out = strip_direct_run_properties(out, run_properties)
    # Clean up empty rPr blocks that might result
    out = re.sub(r"<w:rPr>\s*</w:rPr>", "", out)
    out = re.sub(r"<w:rPr\s*/>", "", out)
    # Clean up empty pPr blocks that might result
    out = re.sub(r"<w:pPr>\s*</w:pPr>", "", out)
    out = re.sub(r"<w:pPr\s*/>", "", out)
    return _restore_named_subtrees(out, protected_subtrees)


def _normalize_paragraph_for_contract(
    p_xml: str,
    allowed_ppr_properties: Optional[Set[str]] = None,
    allowed_rpr_properties: Optional[Set[str]] = None,
) -> str:
    return edit_preserving_out_of_scope_subtrees(
        p_xml,
        lambda protected: _normalize_paragraph_for_contract_unprotected(
            protected,
            allowed_ppr_properties,
            allowed_rpr_properties,
        ),
    )


def _strip_direct_numpr_only_unprotected(paragraph_xml: str) -> str:
    protected, preserved = _protect_named_subtrees(
        paragraph_xml,
        ("w:pPrChange", "w:sectPr"),
    )
    out = re.sub(r"<w:numPr\b[^>]*/>", "", protected)
    out = re.sub(r"<w:numPr\b[^>]*>[\s\S]*?</w:numPr>", "", out, flags=re.S)
    out = re.sub(r"<w:pPr>\s*</w:pPr>", "", out)
    out = re.sub(r"<w:pPr\s*/>", "", out)
    return _restore_named_subtrees(out, preserved)


def _strip_direct_numpr_only(paragraph_xml: str) -> str:
    return edit_preserving_out_of_scope_subtrees(
        paragraph_xml,
        _strip_direct_numpr_only_unprotected,
    )


def _inject_direct_numpr(paragraph_xml: str, num_id: int, ilvl: int) -> str:
    def _edit(protected: str) -> str:
        out = _strip_direct_numpr_only_unprotected(protected)
        numpr = (
            f'<w:numPr><w:ilvl w:val="{ilvl}"/>'
            f'<w:numId w:val="{num_id}"/></w:numPr>'
        )
        pstyle = re.search(r"<w:pStyle\b[^>]*/>", out)
        if pstyle:
            return out[:pstyle.end()] + numpr + out[pstyle.end():]
        if re.search(r"<w:pPr\b[^>]*>", out):
            return re.sub(r"(<w:pPr\b[^>]*>)", rf"\1{numpr}", out, count=1)
        return re.sub(
            r"(<w:p\b[^>]*>)",
            rf"\1<w:pPr>{numpr}</w:pPr>",
            out,
            count=1,
        )

    return edit_preserving_out_of_scope_subtrees(paragraph_xml, _edit)


def _paragraph_uses_automatic_numbering(paragraph_xml: str, styles_xml: str) -> bool:
    analysis_xml = strip_out_of_scope_subtrees(paragraph_xml)
    return _effective_numpr(analysis_xml, styles_xml) is not None


def _materialize_effective_numpr(
    paragraph_xml: str,
    source_styles_xml: str,
) -> str:
    """Write the complete effective source numPr before changing pStyle."""

    effective = _effective_numpr(
        strip_out_of_scope_subtrees(paragraph_xml),
        source_styles_xml,
    )
    if effective is None:
        return paragraph_xml
    num_id = str(effective.get("numId") or "")
    ilvl = str(effective.get("ilvl") or "0")
    if not re.fullmatch(r"\d+", num_id) or not re.fullmatch(r"\d+", ilvl):
        raise ValueError(
            "Target paragraph has invalid effective Word numbering "
            f"(numId={num_id!r}, ilvl={ilvl!r})"
        )
    return _inject_direct_numpr(paragraph_xml, int(num_id), int(ilvl))


def _ensure_explicit_numpr_preserving_sectpr(
    paragraph_xml: str,
    styles_xml: str,
) -> str:
    """Materialize inherited numbering without treating sectPr as a skip flag."""
    def _edit(protected: str) -> str:
        out = ensure_explicit_numpr_from_current_style(protected, styles_xml)
        if out != protected or "<w:numPr" in protected:
            return out
        if "<w:sectPr" not in protected:
            return out

        style_id = paragraph_pstyle_from_block(protected)
        numpr = _find_style_numpr_in_chain(styles_xml, style_id) if style_id else None
        if not numpr:
            return protected
        if re.search(r"<w:pPr\b[^>]*>", protected):
            return re.sub(r"(<w:pPr\b[^>]*>)", rf"\1{numpr}", protected, count=1)
        return re.sub(
            r"(<w:p\b[^>]*>)",
            rf"\1<w:pPr>{numpr}</w:pPr>",
            protected,
            count=1,
        )

    return edit_preserving_out_of_scope_subtrees(paragraph_xml, _edit)


def _has_literal_numbering_marker(paragraph_xml: str) -> bool:
    text = paragraph_text_from_block(paragraph_xml).strip()
    return bool(re.match(
        r"(?:SECTION\s+\d|PART\s+\d|\d{1,2}\.\d{1,3}\b|\.\d+\s|"
        r"[A-Z]\.\s|\d+\.\s|[a-z]\.\s|\d+\)\s|[a-z]\)\s|"
        r"\(\d+\)\s|\([a-z]+\)\s|[ivxlcdm]+[.)]\s)",
        text,
        flags=re.IGNORECASE,
    ))


def _resolve_application_policy(
    policy: Optional["ApplicationPolicy"],
    conversion_mode: Optional[str],
) -> "ApplicationPolicy":
    """Resolve policy lazily to avoid classification/application-policy cycles."""

    from .application_policy import application_policy_for_mode

    if policy is None:
        return application_policy_for_mode(conversion_mode or "format_only")
    if conversion_mode is not None and policy.conversion_mode != conversion_mode:
        raise ValueError(
            "application policy conversion_mode does not match explicit "
            f"conversion_mode ({policy.conversion_mode!r} != {conversion_mode!r})"
        )
    return policy


def _style_replacement_ppr_properties(
    styles_xml_text: str,
    style_id: str,
) -> Set[str]:
    """Return paragraph properties supplied anywhere in a style's basedOn chain."""

    style_map = _build_style_xml_map(styles_xml_text)
    properties: Set[str] = set()
    visited: Set[str] = set()
    current = style_id
    while current and current not in visited:
        visited.add(current)
        block = style_map.get(current, "")
        if not block:
            break
        try:
            wrapped = ET.fromstring(
                '<root xmlns:w="http://schemas.openxmlformats.org/'
                f'wordprocessingml/2006/main">{block}</root>'
            )
        except ET.ParseError as exc:
            raise ValueError(
                f"Could not parse imported style {current!r} while resolving pPr"
            ) from exc
        style_element = next(iter(wrapped), None)
        ppr = style_element.find(_wq("pPr")) if style_element is not None else None
        if ppr is not None:
            properties.update(
                child.tag.rsplit("}", 1)[-1]
                for child in ppr
                if isinstance(child.tag, str)
            )
        based_on = re.search(
            r'<w:basedOn\b[^>]*w:val="([^"]+)"',
            block,
        )
        current = based_on.group(1) if based_on else ""
    return properties


_PROTECTED_DIRECT_RPR_PROPERTIES = frozenset({"rStyle", "rPrChange"})


def _style_replacement_rpr_properties(
    styles_xml_text: str,
    style_id: str,
) -> Set[str]:
    """Return run properties supplied by the effective replacement style."""

    style_map = _build_style_xml_map(styles_xml_text)
    properties: Set[str] = set()
    visited: Set[str] = set()
    current = style_id
    while current and current not in visited:
        visited.add(current)
        block = style_map.get(current, "")
        if not block:
            break
        try:
            wrapped = ET.fromstring(
                '<root xmlns:w="http://schemas.openxmlformats.org/'
                f'wordprocessingml/2006/main">{block}</root>'
            )
        except ET.ParseError as exc:
            raise ValueError(
                f"Could not parse imported style {current!r} while resolving rPr"
            ) from exc
        style_element = next(iter(wrapped), None)
        rpr = style_element.find(_wq("rPr")) if style_element is not None else None
        if rpr is not None:
            properties.update(
                child.tag.rsplit("}", 1)[-1]
                for child in rpr
                if isinstance(child.tag, str)
            )
        based_on = re.search(
            r'<w:basedOn\b[^>]*w:val="([^"]+)"',
            block,
        )
        current = based_on.group(1) if based_on else ""
    return properties - _PROTECTED_DIRECT_RPR_PROPERTIES


def _strip_direct_ppr_properties(
    paragraph_xml: str,
    properties: Set[str],
) -> str:
    """Strip only direct properties that the effective architect style replaces."""

    if not properties:
        return paragraph_xml

    def _edit(protected: str) -> str:
        protected, tracked_subtrees = _protect_named_subtrees(
            protected,
            ("w:pPrChange", "w:sectPr"),
        )

        def _strip_from_ppr(match: re.Match[str]) -> str:
            ppr = match.group(0)
            for name in sorted(properties):
                ppr = re.sub(rf"<w:{name}\b[^>]*/>", "", ppr)
                ppr = re.sub(
                    rf"<w:{name}\b[^>]*>[\s\S]*?</w:{name}>",
                    "",
                    ppr,
                    flags=re.S,
                )
            return ppr

        edited = re.sub(
            r"<w:pPr\b[^>]*>[\s\S]*?</w:pPr>",
            _strip_from_ppr,
            protected,
            count=1,
            flags=re.S,
        )
        return _restore_named_subtrees(edited, tracked_subtrees)

    return edit_preserving_out_of_scope_subtrees(paragraph_xml, _edit)


def _effective_numbering_semantics(
    paragraph_xml: str,
    styles_xml_text: str,
    numbering_catalog: Dict[str, Any],
) -> Optional[Dict[str, Any]]:
    effective = _effective_numpr(
        strip_out_of_scope_subtrees(paragraph_xml),
        styles_xml_text,
    )
    if effective is None:
        return None
    return {
        "numPr": {
            "numId": str(effective.get("numId", "")),
            "ilvl": str(effective.get("ilvl", "0")),
        },
        "pattern": _resolve_numbering_pattern(effective, numbering_catalog),
    }


def apply_phase2_classifications(
    extract_dir: Path,
    classifications: Dict[str, Any],
    arch_style_registry: Dict[str, str],
    log: List[str],
    role_specs: Optional[Dict[str, Dict[str, Any]]] = None,
    role_numpr_remap: Optional[Dict[str, Dict[str, Any]]] = None,
    source_styles_xml: Optional[str] = None,
    source_numbering_xml: Optional[str] = None,
    policy: Optional["ApplicationPolicy"] = None,
    conversion_mode: Optional[str] = None,
) -> "ApplyReport":
    """
    Apply CSI role classifications to paragraphs by setting pStyle.

    Removes direct run properties only where the effective replacement style
    supplies the same property.
    """
    application_policy = _resolve_application_policy(policy, conversion_mode)
    doc_path = extract_dir / "word" / "document.xml"
    doc_text = read_xml_text(doc_path)

    # Load styles once so we can verify architect style IDs exist in target styles.xml
    styles_xml_text = read_xml_text(extract_dir / "word" / "styles.xml")
    numbering_source_styles_xml = (
        source_styles_xml if source_styles_xml is not None else styles_xml_text
    )
    numbering_path = extract_dir / "word" / "numbering.xml"
    current_numbering_xml = (
        read_xml_text(numbering_path) if numbering_path.is_file() else ""
    )
    numbering_source_xml = (
        source_numbering_xml
        if source_numbering_xml is not None
        else current_numbering_xml
    )
    source_numbering_catalog = _build_numbering_catalog(numbering_source_xml)
    current_numbering_catalog = _build_numbering_catalog(current_numbering_xml)
    style_ids_in_styles = set(re.findall(r'w:styleId="([^"]+)"', styles_xml_text))

    blocks = list(iter_paragraph_xml_blocks(doc_text))
    para_blocks = [b[2] for b in blocks]

    report = ApplyReport(requested=0)

    text_before = [paragraph_text_from_block(p) for p in para_blocks]
    numbering_before = [
        _effective_numbering_semantics(
            p,
            numbering_source_styles_xml,
            source_numbering_catalog,
        )
        for p in para_blocks
    ]

    items = classifications.get("classifications", [])
    if not isinstance(items, list):
        raise ValueError("phase2 classifications: 'classifications' must be a list")
    report.requested = len(items)

    ignored_items = classifications.get("ignored_paragraphs", [])
    if not isinstance(ignored_items, list):
        raise ValueError("phase2 classifications: 'ignored_paragraphs' must be a list")
    classified_indices: Set[int] = set()
    for item in items:
        if not isinstance(item, dict):
            continue
        idx = item.get("paragraph_index")
        if not isinstance(idx, int):
            continue
        if idx in classified_indices:
            raise ValueError(f"Duplicate classified paragraph index: {idx}")
        classified_indices.add(idx)
    ignored_indices: Set[int] = set()
    for item in ignored_items:
        if not isinstance(item, dict):
            raise ValueError(f"Invalid ignored entry (not object): {item!r}")
        idx = item.get("paragraph_index")
        reason = item.get("reason")
        if not isinstance(idx, int) or idx < 0 or idx >= len(para_blocks):
            raise ValueError(f"Invalid ignored paragraph index: {idx!r}")
        if idx in ignored_indices:
            raise ValueError(f"Duplicate ignored paragraph index: {idx}")
        if idx in classified_indices:
            raise ValueError(f"Paragraph {idx} cannot be both classified and ignored")
        if not isinstance(reason, str) or not reason.strip():
            raise ValueError(f"Invalid ignored reason at paragraph {idx}: {reason!r}")
        ignored_indices.add(idx)
    report.ignored = len(ignored_indices)

    style_xml_by_id = _build_style_xml_map(styles_xml_text)
    replacement_style_ids = {
        arch_style_registry.get(item.get("csi_role"))
        for item in items
        if isinstance(item, dict) and isinstance(item.get("csi_role"), str)
    }
    replacement_style_ids.discard(None)
    ppr_properties_by_style = {
        style_id: _style_replacement_ppr_properties(styles_xml_text, style_id)
        for style_id in replacement_style_ids
    }
    rpr_properties_by_style = {
        style_id: _style_replacement_rpr_properties(styles_xml_text, style_id)
        for style_id in replacement_style_ids
    }
    allowed_ppr_by_index: List[Set[str]] = [set() for _ in para_blocks]
    allowed_rpr_by_index: List[Set[str]] = [set() for _ in para_blocks]
    for item in items:
        if not isinstance(item, dict):
            continue
        idx = item.get("paragraph_index")
        role = item.get("csi_role")
        if not isinstance(idx, int) or not (0 <= idx < len(para_blocks)):
            continue
        if not isinstance(role, str):
            continue
        style_id = arch_style_registry.get(role)
        if style_id:
            allowed_rpr_by_index[idx] = set(
                rpr_properties_by_style.get(style_id, set())
            )
        if application_policy.preserve_target_numbering and style_id:
            properties = ppr_properties_by_style.get(style_id, set())
            allowed_ppr_by_index[idx] = (
                properties - _FORMAT_ONLY_PROTECTED_PPR_PROPERTIES
            ) | {"numPr"}
        else:
            allowed_ppr_by_index[idx] = {"numPr", "jc", "ind", "spacing"}

    # Contract normalization is per paragraph: Format-only may remove only
    # properties actually supplied by that paragraph's effective architect
    # style, while Canadian mode retains its legacy broader mutation contract.
    contract_before = [
        _normalize_paragraph_for_contract(
            p,
            allowed_ppr_by_index[idx],
            allowed_rpr_by_index[idx],
        )
        for idx, p in enumerate(para_blocks)
    ]

    for item in items:
        if not isinstance(item, dict):
            raise ValueError(f"Invalid classification entry (not object): {item!r}")

        idx = item.get("paragraph_index")
        role = item.get("csi_role")

        if not isinstance(idx, int) or idx < 0 or idx >= len(para_blocks):
            report.invalid_indices.append(idx)
            continue

        if not isinstance(role, str):
            raise ValueError(f"Invalid csi_role type at paragraph {idx}: {role!r}")

        style_id = arch_style_registry.get(role)
        if not style_id:
            report.unmapped_roles.append((idx, role))
            continue

        if style_id not in style_ids_in_styles:
            report.missing_style_ids.add(style_id)
            continue

        pb = para_blocks[idx]
        style_xml = style_xml_by_id.get(style_id, "")
        role_spec = role_specs.get(role) if role_specs else None
        provenance = role_spec.get("numbering_provenance") if isinstance(role_spec, dict) else None
        if role_specs is not None and not isinstance(role_spec, dict):
            raise ValueError(f"Missing full role contract for classified role: {role}")

        if application_policy.preserve_target_numbering:
            source_effective = numbering_before[idx]
            replacement_properties = ppr_properties_by_style.get(style_id, set())
            replaceable_properties = (
                replacement_properties - _FORMAT_ONLY_PROTECTED_PPR_PROPERTIES
            )
            if replaceable_properties:
                pb = _strip_direct_ppr_properties(pb, replaceable_properties)
                report.stripped_direct_ppr += 1
            else:
                report.preserved_direct_ppr += 1

            if source_effective is not None:
                pb = _materialize_effective_numpr(
                    pb,
                    numbering_source_styles_xml,
                )
                report.preserved_automatic_numbering += 1
            elif "numPr" in replacement_properties:
                # An unnumbered target paragraph must remain unnumbered even
                # when its replacement style inherits architect numbering.
                direct = paragraph_numpr_from_block(
                    strip_out_of_scope_subtrees(pb)
                )
                if direct.get("numId") != "0":
                    pb = _inject_direct_numpr(pb, 0, 0)
                report.suppressed_architect_numbering += 1
        elif provenance in {"style_numpr", "direct_numpr"}:
            pb = strip_conflicting_direct_ppr(pb)
            report.stripped_direct_ppr += 1
        elif provenance == "text_literal":
            target_automatic = _paragraph_uses_automatic_numbering(
                pb,
                numbering_source_styles_xml,
            )
            if target_automatic and not _has_literal_numbering_marker(pb):
                # A typed marker belongs to the architect exemplar's text; it
                # is not part of the reusable paragraph style.  Format-only
                # application must therefore retain the target's proven Word
                # numbering instead of deleting it or inventing literal text.
                pb = _materialize_effective_numpr(
                    pb,
                    numbering_source_styles_xml,
                )
                if _style_has_replacement_ppr(style_xml):
                    pb = strip_conflicting_direct_ppr(pb, preserve_numpr=True)
                    report.stripped_direct_ppr += 1
                else:
                    report.preserved_direct_ppr += 1
                report.preserved_automatic_numbering += 1
            elif _style_has_replacement_ppr(style_xml):
                pb = strip_conflicting_direct_ppr(pb)
                report.stripped_direct_ppr += 1
            else:
                pb = _strip_direct_numpr_only(pb)
                report.preserved_direct_ppr += 1
        elif provenance == "none":
            if _style_has_replacement_ppr(style_xml):
                pb = strip_conflicting_direct_ppr(pb)
                report.stripped_direct_ppr += 1
            else:
                pb = _strip_direct_numpr_only(pb)
                report.preserved_direct_ppr += 1
        elif _style_has_replacement_ppr(style_xml):
            pb = strip_conflicting_direct_ppr(pb)
            report.stripped_direct_ppr += 1
        else:
            pb = _ensure_explicit_numpr_preserving_sectpr(pb, styles_xml_text)
            report.preserved_direct_ppr += 1

        replacement_run_properties = rpr_properties_by_style.get(style_id, set())
        updated_pb = strip_direct_run_properties(pb, replacement_run_properties)
        if updated_pb != pb:
            report.stripped_run_fonts += 1
        pb = updated_pb

        # Now safely swap pStyle
        pb = apply_pstyle_to_paragraph_block(pb, style_id)
        if provenance == "direct_numpr" and application_policy.import_body_numbering:
            remap = (role_numpr_remap or {}).get(role)
            if not isinstance(remap, dict):
                raise ValueError(f"Missing imported direct numbering mapping for role: {role}")
            pb = _inject_direct_numpr(pb, int(remap["new_numId"]), int(remap.get("ilvl", 0)))
        para_blocks[idx] = pb
        report.modified += 1

    if report.invalid_indices:
        raise ValueError(f"Invalid paragraph indices: {report.invalid_indices[:20]}")
    if report.missing_style_ids:
        raise ValueError(f"Missing style IDs in target styles.xml: {sorted(report.missing_style_ids)}")
    if report.unmapped_roles:
        preview = report.unmapped_roles[:10]
        raise ValueError(f"Unmapped roles encountered: {preview}")
    expected_targetable = report.requested
    if expected_targetable > 0 and report.modified == 0:
        raise ValueError("Applied styles to 0 paragraphs despite non-empty targetable classifications")
    if report.modified != expected_targetable:
        raise ValueError(
            f"Applied styles to {report.modified}/{expected_targetable} targetable paragraphs"
        )

    log.append(f"Applied styles to {report.modified} paragraphs")
    log.append(
        "Direct paragraph overrides stripped/preserved: "
        f"{report.stripped_direct_ppr}/{report.preserved_direct_ppr}"
    )
    log.append(
        "Stripped style-replaced direct run properties from "
        f"{report.stripped_run_fonts} paragraphs"
    )
    if report.preserved_automatic_numbering:
        log.append(
            "Preserved source Word numbering for "
            f"{report.preserved_automatic_numbering} paragraphs"
        )
    if report.suppressed_architect_numbering:
        log.append(
            "Suppressed architect numbering on "
            f"{report.suppressed_architect_numbering} unnumbered target paragraphs"
        )
    if report.ignored:
        log.append(f"Left {report.ignored} explicitly ignored paragraphs untouched")

    # Enforce the diff contract.
    contract_after = [
        _normalize_paragraph_for_contract(
            p,
            allowed_ppr_by_index[idx],
            allowed_rpr_by_index[idx],
        )
        for idx, p in enumerate(para_blocks)
    ]
    if len(contract_before) != len(contract_after):
        raise RuntimeError("Internal error: paragraph count changed during Phase 2 application")

    for i, (b, a) in enumerate(zip(contract_before, contract_after)):
        if b != a:
            raise ValueError(
                "Phase 2 invariant violation: paragraph content changed outside allowed edits "
                f"(pStyle/numPr/style-replaced formatting) at paragraph index {i}."
            )

    for idx in sorted(ignored_indices):
        if para_blocks[idx] != blocks[idx][2]:
            raise ValueError(
                "Phase 2 invariant violation: explicitly ignored paragraph "
                f"{idx} was modified"
            )

    if application_policy.preserve_target_numbering:
        text_after = [paragraph_text_from_block(p) for p in para_blocks]
        if text_before != text_after:
            changed = next(
                idx
                for idx, (before, after) in enumerate(zip(text_before, text_after))
                if before != after
            )
            raise ValueError(
                "FORMAT_ONLY invariant violation: target body text changed at "
                f"paragraph index {changed}"
            )
        numbering_after = [
            _effective_numbering_semantics(
                p,
                styles_xml_text,
                current_numbering_catalog,
            )
            for p in para_blocks
        ]
        if numbering_before != numbering_after:
            changed = next(
                idx
                for idx, (before, after) in enumerate(
                    zip(numbering_before, numbering_after)
                )
                if before != after
            )
            raise ValueError(
                "FORMAT_ONLY invariant violation: effective target numbering "
                f"changed at paragraph index {changed}"
            )
        report.numbering_checks = {
            "policy": application_policy.conversion_mode,
            "paragraphs_checked": len(para_blocks),
            "automatic_numbered_before": sum(
                semantics is not None for semantics in numbering_before
            ),
            "effective_numbering_preserved": True,
            "body_text_preserved": True,
        }

    # Rebuild document.xml
    out = []
    last = 0
    for (s, e, _), pb in zip(blocks, para_blocks):
        out.append(doc_text[last:s])
        out.append(pb)
        last = e
    out.append(doc_text[last:])
    write_xml_text(doc_path, "".join(out))
    return report


@dataclass
class ApplyReport:
    requested: int
    modified: int = 0
    invalid_indices: List[Any] = field(default_factory=list)
    skipped_sectpr: List[int] = field(default_factory=list)
    unmapped_roles: List[Tuple[int, str]] = field(default_factory=list)
    missing_style_ids: Set[str] = field(default_factory=set)
    stripped_direct_ppr: int = 0
    preserved_direct_ppr: int = 0
    preserved_automatic_numbering: int = 0
    suppressed_architect_numbering: int = 0
    stripped_run_fonts: int = 0
    ignored: int = 0
    numbering_checks: Dict[str, Any] = field(default_factory=dict)


def _build_style_xml_map(styles_xml_text: str) -> Dict[str, str]:
    out: Dict[str, str] = {}
    for match in re.finditer(r'(<w:style\b[^>]*w:styleId="([^"]+)"[^>]*>[\s\S]*?</w:style>)', styles_xml_text):
        out[match.group(2)] = match.group(1)
    return out


def _style_has_replacement_ppr(style_block_xml: str) -> bool:
    if not style_block_xml:
        return False
    return any(
        tag in style_block_xml
        for tag in ("<w:spacing", "<w:ind", "<w:jc", "<w:numPr")
    )
