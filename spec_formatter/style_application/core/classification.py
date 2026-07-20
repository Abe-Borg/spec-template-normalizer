"""
Phase 2 classification: applying LLM classifications to paragraphs,
building slim bundles for LLM input, and boilerplate filtering.
"""

import re
import difflib
import xml.etree.ElementTree as ET
from dataclasses import dataclass, field
from pathlib import Path
from typing import Dict, Any, List, Optional, Set, Tuple

from spec_formatter.numbering_roles import role_from_numbering_signature

from .xml_helpers import (
    iter_paragraph_xml_blocks,
    iter_element_xml_blocks,
    paragraph_text_from_block,
    paragraph_contains_sectpr,
    paragraph_numpr_from_block,
    paragraph_pstyle_from_block,
    paragraph_ppr_hints_from_block,
    apply_pstyle_to_paragraph_block,
    strip_run_font_formatting,
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
_ARTICLE_RX = re.compile(r"^\s*\d+\.\d{2}\b")
_SECTION_ID_RX = re.compile(r"^\s*SECTION\s+\d{2}(?:\s+\d{2}){2,}\b", re.IGNORECASE)
_END_OF_SECTION_RX = re.compile(r"^\s*END\s+OF\s+SECTION\s*", re.IGNORECASE)
_ALL_CAPS_RX = re.compile(r"^[^a-z]*[A-Z][^a-z]*$")
_MARKER_RX = [
    (re.compile(r"^\s*[A-Z]\.\s+"), "upper_alpha"),
    (re.compile(r"^\s*\d+\.\s+"), "number"),
    (re.compile(r"^\s*[a-z]\.\s+"), "lower_alpha"),
]


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
            num_fmt = _wval(level.find(_wq("numFmt")))
            lvl_text = _wval(level.find(_wq("lvlText")))
            if num_fmt is not None:
                item["numFmt"] = num_fmt
            if lvl_text is not None:
                item["lvlText"] = lvl_text
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
                num_fmt = _wval(override_level.find(_wq("numFmt")))
                lvl_text = _wval(override_level.find(_wq("lvlText")))
                if num_fmt is not None:
                    override_item["numFmt"] = num_fmt
                if lvl_text is not None:
                    override_item["lvlText"] = lvl_text
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
            for key in ("numFmt", "lvlText"):
                if isinstance(level.get(key), str):
                    pattern[key] = level[key]

    override = num.get("overrides", {}).get(ilvl, {}) if isinstance(num, dict) else {}
    if isinstance(override, dict):
        for key in ("numFmt", "lvlText", "startOverride"):
            if isinstance(override.get(key), str):
                pattern[key] = override[key]
    return pattern


def _numbering_role_candidates(
    target_pattern: Optional[Dict[str, str]],
    role_specs: Optional[Dict[str, Dict[str, Any]]],
    available_roles: List[str],
) -> List[str]:
    """Return roles whose portable rendered numbering signature is exact.

    ``numId`` and ``abstractNumId`` are package-local identifiers, so matching
    those across the architect and target packages would be incorrect.  A
    deterministic match requires both rendered signals (format and level text),
    plus every other portable constraint declared by the role.
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

        # This is an exact rendered signature.  In particular, a restart is
        # not interchangeable with a non-restarted level (or a different
        # restart value), even when format and level text match.
        keys = ["ilvl", "numFmt", "lvlText", "startOverride"]
        if all(
            str(target_pattern.get(key, "")) == str(expected.get(key, ""))
            for key in keys
        ):
            matches.append(role)
    return matches


def _resolve_role(preferred: str, available_roles: List[str]) -> Optional[str]:
    fallback_chain = {
        "SectionID": ["SectionID", "SectionTitle"],
        "END_OF_SECTION": ["END_OF_SECTION"],
        "SUBSUBPARAGRAPH": ["SUBSUBPARAGRAPH", "SUBPARAGRAPH", "PARAGRAPH"],
        "SUBPARAGRAPH": ["SUBPARAGRAPH", "PARAGRAPH"],
        "PARAGRAPH": ["PARAGRAPH"],
        "PART": ["PART"],
        "ARTICLE": ["ARTICLE"],
        "SectionTitle": ["SectionTitle"],
    }
    for candidate in fallback_chain.get(preferred, [preferred]):
        if candidate in available_roles:
            return candidate
    return None


def _deterministic_role_for_paragraph(paragraph: Dict[str, Any], prev_text: str = "") -> Optional[str]:
    text = paragraph.get("text", "")
    if not text or paragraph.get("in_table"):
        return None
    if _SECTION_ID_RX.match(text):
        return "SectionID"
    if _END_OF_SECTION_RX.match(text):
        return "END_OF_SECTION"
    if _PART_RX.match(text):
        return "PART"
    if _ARTICLE_RX.match(text):
        return "ARTICLE"
    if prev_text and _SECTION_ID_RX.match(prev_text) and _ALL_CAPS_RX.match(text):
        return "SectionTitle"

    # Automatic numbering is deterministic only when its effective rendered
    # signature uniquely matches one architect role.  Ambiguous or incomplete
    # numbering metadata must remain in the LLM input.
    if paragraph.get("effective_numPr"):
        numbering_role = paragraph.get("numbering_role")
        if isinstance(numbering_role, str):
            return numbering_role
        pattern = paragraph.get("numbering_pattern")
        if isinstance(pattern, dict):
            return role_from_numbering_signature(
                pattern.get("numFmt"),
                pattern.get("lvlText"),
                pattern.get("ilvl"),
            )
        return None

    marker_type = paragraph.get("marker_type")
    if marker_type == "upper_alpha":
        return "PARAGRAPH"
    if marker_type == "number":
        return "SUBPARAGRAPH"
    if marker_type == "lower_alpha":
        return "SUBSUBPARAGRAPH"
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
            "PART",
            "ARTICLE",
            "PARAGRAPH",
            "SUBPARAGRAPH",
            "SUBSUBPARAGRAPH",
        ]

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
        if contains_drawing_or_textbox:
            filter_report["paragraphs_out_of_scope"].append({
                "paragraph_index": idx,
                "reason": "drawing_or_textbox_subtree",
                "original_text_preview": raw_text[:120],
            })

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
            continue

        if tags:
            filter_report["paragraphs_stripped"].append({
                "paragraph_index": idx,
                "tags": tags
            })

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
                continue

        numpr = paragraph_numpr_from_block(analysis_xml)
        effective_numpr = _effective_numpr(analysis_xml, styles_text)
        numbering_pattern = _resolve_numbering_pattern(effective_numpr, numbering_catalog)
        numbering_candidates = _numbering_role_candidates(
            numbering_pattern,
            role_specs,
            available_roles,
        )
        pstyle = paragraph_pstyle_from_block(analysis_xml)
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
            "numbering_role": numbering_candidates[0] if len(numbering_candidates) == 1 else None,
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
        ]
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


def validate_phase2_llm_payload(bundle: Dict[str, Any], classifications: Dict[str, Any], allowed_roles: List[str]) -> None:
    unresolved = {p["paragraph_index"] for p in bundle.get("paragraphs", [])}
    seen = _validate_payload_shape(classifications, allowed_roles, unresolved)
    missing = sorted(unresolved - set(seen.keys()))
    if missing:
        raise ValueError(f"missing coverage for paragraph indices: {missing[:20]}")


def coerce_to_final_classifications(
    bundle: Dict[str, Any],
    classifications: Dict[str, Any],
    allowed_roles: List[str],
) -> Dict[str, Any]:
    unresolved = {p["paragraph_index"] for p in bundle.get("paragraphs", [])}
    deterministic: Dict[int, str] = {
        item["paragraph_index"]: item["csi_role"]
        for item in bundle.get("deterministic_classifications", [])
        if isinstance(item, dict)
        and isinstance(item.get("paragraph_index"), int)
        and isinstance(item.get("csi_role"), str)
    }
    total = set(deterministic.keys()) | unresolved

    incoming = _validate_payload_shape(classifications, allowed_roles, total)
    incoming_indices = set(incoming.keys())

    if incoming_indices <= unresolved:
        missing_unresolved = sorted(unresolved - incoming_indices)
        if missing_unresolved:
            raise ValueError(f"missing coverage for paragraph indices: {missing_unresolved[:20]}")
        merged = dict(deterministic)
        merged.update(incoming)
    elif incoming_indices == total:
        for idx, expected in deterministic.items():
            actual = incoming.get(idx)
            if actual != expected:
                raise ValueError(
                    f"deterministic override attempted at paragraph_index={idx}: "
                    f"expected {expected!r}, got {actual!r}"
                )
        merged = incoming
    else:
        raise ValueError("payload is neither unresolved-only nor valid final coverage")

    notes = classifications.get("notes", []) if isinstance(classifications, dict) else []
    return {
        "classifications": [
            {"paragraph_index": idx, "csi_role": role}
            for idx, role in sorted(merged.items())
        ],
        "notes": notes if isinstance(notes, list) else [],
    }


def validate_phase2_final_payload(bundle: Dict[str, Any], classifications: Dict[str, Any], allowed_roles: List[str]) -> None:
    coerce_to_final_classifications(bundle, classifications, allowed_roles)


def validate_phase2_classification_contract(bundle: Dict[str, Any], classifications: Dict[str, Any], allowed_roles: List[str]) -> None:
    # Backward-compatible alias: strict unresolved-only LLM payload validation.
    validate_phase2_llm_payload(bundle, classifications, allowed_roles)


def _normalize_paragraph_for_contract_unprotected(p_xml: str) -> str:
    """
    Normalize paragraph for contract comparison.
    Strips elements we're allowed to change: pStyle, numPr, and
    run-level font formatting (rFonts, sz, szCs).  Also removes
    empty pPr / rPr shells so paragraphs that originally lacked
    these blocks compare equal after stripping.
    """
    out = p_xml
    # Strip pStyle (we change this)
    out = re.sub(r"<w:pStyle\b[^>]*/>", "", out)
    # Strip numPr (inline numbering overrides are allowed to change)
    out = re.sub(r"<w:numPr\b[^>]*/>", "", out)
    out = re.sub(r"<w:numPr\b[^>]*>[\s\S]*?</w:numPr>", "", out, flags=re.S)
    # Strip direct pPr overrides now allowed to be removed during apply
    out = re.sub(r"<w:jc\b[^>]*/>", "", out)
    out = re.sub(r"<w:jc\b[^>]*>[\s\S]*?</w:jc>", "", out, flags=re.S)
    out = re.sub(r"<w:ind\b[^>]*/>", "", out)
    out = re.sub(r"<w:ind\b[^>]*>[\s\S]*?</w:ind>", "", out, flags=re.S)
    out = re.sub(r"<w:spacing\b[^>]*/>", "", out)
    out = re.sub(r"<w:spacing\b[^>]*>[\s\S]*?</w:spacing>", "", out, flags=re.S)
    # Strip run-level font formatting (we now strip this too)
    out = re.sub(r"<w:rFonts\b[^>]*/>", "", out)
    out = re.sub(r"<w:rFonts\b[^>]*>[\s\S]*?</w:rFonts>", "", out, flags=re.S)
    out = re.sub(r"<w:sz\b[^>]*/>", "", out)
    out = re.sub(r"<w:szCs\b[^>]*/>", "", out)
    # Clean up empty rPr blocks that might result
    out = re.sub(r"<w:rPr>\s*</w:rPr>", "", out)
    out = re.sub(r"<w:rPr\s*/>", "", out)
    # Clean up empty pPr blocks that might result
    out = re.sub(r"<w:pPr>\s*</w:pPr>", "", out)
    out = re.sub(r"<w:pPr\s*/>", "", out)
    return out


def _normalize_paragraph_for_contract(p_xml: str) -> str:
    return edit_preserving_out_of_scope_subtrees(
        p_xml,
        _normalize_paragraph_for_contract_unprotected,
    )


def _strip_direct_numpr_only_unprotected(paragraph_xml: str) -> str:
    out = re.sub(r"<w:numPr\b[^>]*/>", "", paragraph_xml)
    out = re.sub(r"<w:numPr\b[^>]*>[\s\S]*?</w:numPr>", "", out, flags=re.S)
    out = re.sub(r"<w:pPr>\s*</w:pPr>", "", out)
    out = re.sub(r"<w:pPr\s*/>", "", out)
    return out


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
    if "<w:numPr" in analysis_xml:
        return True
    style_id = paragraph_pstyle_from_block(analysis_xml)
    return bool(style_id and _find_style_numpr_in_chain(styles_xml, style_id))


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
        r"(?:SECTION\s+\d|PART\s+\d|\d+\.\d{2,}\b|[A-Z]\.\s|\d+\.\s|[a-z]\.\s)",
        text,
        flags=re.IGNORECASE,
    ))


def apply_phase2_classifications(
    extract_dir: Path,
    classifications: Dict[str, Any],
    arch_style_registry: Dict[str, str],
    log: List[str],
    role_specs: Optional[Dict[str, Dict[str, Any]]] = None,
    role_numpr_remap: Optional[Dict[str, Dict[str, Any]]] = None,
) -> "ApplyReport":
    """
    Apply CSI role classifications to paragraphs by setting pStyle.

    Also strips run-level font formatting so the style's fonts take effect.
    This handles MasterSpec/ARCOM documents that have hardcoded fonts in every run.
    """
    doc_path = extract_dir / "word" / "document.xml"
    doc_text = read_xml_text(doc_path)

    # Load styles once so we can verify architect style IDs exist in target styles.xml
    styles_xml_text = read_xml_text(extract_dir / "word" / "styles.xml")
    style_ids_in_styles = set(re.findall(r'w:styleId="([^"]+)"', styles_xml_text))

    blocks = list(iter_paragraph_xml_blocks(doc_text))
    para_blocks = [b[2] for b in blocks]

    report = ApplyReport(requested=0)

    # Contract check: normalize paragraphs for comparison
    contract_before = [_normalize_paragraph_for_contract(p) for p in para_blocks]

    items = classifications.get("classifications", [])
    if not isinstance(items, list):
        raise ValueError("phase2 classifications: 'classifications' must be a list")
    report.requested = len(items)

    style_xml_by_id = _build_style_xml_map(styles_xml_text)

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

        if provenance in {"style_numpr", "direct_numpr"}:
            pb = strip_conflicting_direct_ppr(pb)
            report.stripped_direct_ppr += 1
        elif provenance in {"none", "text_literal"}:
            if (
                provenance == "text_literal"
                and _paragraph_uses_automatic_numbering(pb, styles_xml_text)
                and not _has_literal_numbering_marker(pb)
            ):
                raise ValueError(
                    f"Role {role} requires literal-text numbering, but target paragraph {idx} "
                    "uses automatic numbering and has no literal marker"
                )
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

        # Strip run-level font formatting so style fonts take effect
        pb = strip_run_font_formatting(pb)
        report.stripped_run_fonts += 1

        # Now safely swap pStyle
        pb = apply_pstyle_to_paragraph_block(pb, style_id)
        if provenance == "direct_numpr":
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
    log.append(f"Stripped run-level font formatting from {report.stripped_run_fonts} paragraphs")

    # Enforce the diff contract.
    contract_after = [_normalize_paragraph_for_contract(p) for p in para_blocks]
    if len(contract_before) != len(contract_after):
        raise RuntimeError("Internal error: paragraph count changed during Phase 2 application")

    for i, (b, a) in enumerate(zip(contract_before, contract_after)):
        if b != a:
            diff = "\n".join(difflib.unified_diff(
                b.splitlines(),
                a.splitlines(),
                fromfile=f"before:p[{i}]",
                tofile=f"after:p[{i}]",
                lineterm=""
            ))
            raise ValueError(
                "Phase 2 invariant violation: paragraph content changed outside allowed edits "
                f"(pStyle/numPr/run fonts) at paragraph index {i}.\n" + diff[:4000]
            )

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
    stripped_run_fonts: int = 0


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
