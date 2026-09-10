"""Fail-closed CSI-to-Canadian PageFormat hierarchy conversion.

The converter deliberately changes only leading, typed CSI/CSC numbering
markers.  The selected architect template remains the sole source of the
rendered Canadian numbering, styles, and layout.  Consequently Canadian mode
requires the architect's numbered roles to use true Word automatic numbering
with Canadian numeric signatures.
"""

from __future__ import annotations

import re
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Dict, Optional, Tuple

from spec_formatter.role_contract import (
    BODY_HIERARCHY_ROLES,
    ROLE_LEVEL,
)

from .conversion_modes import (  # re-exported: long-standing import site
    CSI_TO_CANADIAN,
    FORMAT_ONLY,
    VALID_CONVERSION_MODES,
    validate_conversion_mode,
)
from .marker_tools import (
    NUMBERED_ROLES,
    _detect_any_literal_marker,
    _detect_literal_marker,
    _find_numbering_level,
    _has_heading_like_article_body,
    _marker_markup_delimiter,
    _paragraph_locator,
    _remove_literal_marker,
    _SourceEvidence,
    _validate_automatic_source,
    _validate_numbering_start,
    _validate_source_sequence,
    _verify_changed_paragraph,
)
from .untrusted_xml import parse_untrusted_xml
from .errors import EngineError
from .classification import (
    _build_numbering_catalog,
    _effective_numpr,
    _resolve_numbering_pattern,
)
from .ooxml_text import prepare_xml_text_for_utf8, read_xml_text, write_xml_text
from .sectpr_tools import extract_all_sectpr_blocks
from .xml_helpers import (
    iter_paragraph_xml_blocks,
    paragraph_text_from_block,
    strip_out_of_scope_subtrees,
)


PRESERVED_UNNUMBERED_ROLE = "unproven_numbered_role_preserved"


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
    #: Whether the source document had ``<w:trackRevisions/>`` set, and whether
    #: the markers this run wrote were therefore written as tracked insertions.
    #: Both are recorded, always, so "the markers are plain text" is a visible
    #: decision in the run record rather than something a reader has to infer
    #: from the absence of a field.
    source_tracks_revisions: bool = False
    markers_tracked: bool = False
    marker_author: Optional[str] = None

    def as_dict(self) -> Dict[str, Any]:
        return {
            "paragraphs_examined": self.paragraphs_examined,
            "paragraphs_converted": self.paragraphs_converted,
            "literal_markers_removed": self.literal_markers_removed,
            "automatic_numbering_retargeted": self.automatic_numbering_retargeted,
            "unnumbered_paragraphs_numbered": self.unnumbered_paragraphs_numbered,
            "source_tracks_revisions": self.source_tracks_revisions,
            "markers_tracked": self.markers_tracked,
            "marker_author": self.marker_author,
            "edits": [item.as_dict() for item in self.edits],
            "warnings": [item.as_dict() for item in self.warnings],
        }


@dataclass(frozen=True)
class ConversionPlan:
    document_xml: str
    report: CanadianConversionReport


def classifications_for_canadian_application(
    classifications: Dict[str, Any],
    report: CanadianConversionReport,
) -> Dict[str, Any]:
    """Exclude markerless numbered guesses before architect styles are applied.

    The classifier still retains complete coverage for its audit and for
    format-only mode.  Canadian conversion, however, must not let a semantic
    LLM guess apply an automatically numbered architect style to source prose
    that has no typed or automatic numbering evidence.
    """

    preserved = {
        issue.paragraph_index
        for issue in report.warnings
        if issue.code == PRESERVED_UNNUMBERED_ROLE
    }
    if not preserved:
        return classifications

    filtered = dict(classifications)
    filtered["classifications"] = [
        item
        for item in classifications.get("classifications", [])
        if not (
            isinstance(item, dict)
            and item.get("paragraph_index") in preserved
        )
    ]
    return filtered




def _validate_canadian_role_contract(
    role: str,
    spec: object,
    *,
    allow_numeric_part: bool = False,
) -> None:
    if not isinstance(spec, dict):
        raise EngineError("canadian_architect_contract", 
            f"Architect template: role {role} has no complete role contract; "
            "Canadian conversion requires a strict analyzed profile."
        )
    provenance = spec.get("numbering_provenance")
    if provenance not in {"style_numpr", "direct_numpr"}:
        raise EngineError("canadian_architect_contract", 
            f"Architect template: role {role} must use true Word automatic "
            f"numbering for Canadian conversion; found {provenance!r}."
        )
    pattern = spec.get("numbering_pattern")
    if not isinstance(pattern, dict):
        raise EngineError("canadian_architect_contract", 
            f"Architect template: role {role} has no numbering pattern; "
            "Canadian conversion cannot proceed."
        )
    num_fmt = str(pattern.get("numFmt") or "")
    lvl_text = str(pattern.get("lvlText") or "")
    for key in ("start", "startOverride"):
        value = pattern.get(key)
        if value is not None and str(value) != "1":
            raise EngineError("canadian_architect_contract", 
                f"Architect template: role {role} starts at {value!r}; Canadian conversion "
                "requires numbering that starts at 1."
            )
    if pattern.get("lvlRestart") is not None:
        raise EngineError("canadian_architect_contract", 
            f"Architect template: role {role} uses an explicit numbering restart rule that "
            "Canadian conversion cannot yet prove safe."
        )
    if num_fmt != "decimal":
        raise EngineError("canadian_architect_contract", 
            f"Architect template: role {role} is not Canadian numeric numbering "
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
        raise EngineError("canadian_architect_contract", 
            f"Architect template: role {role} does not demonstrate Canadian PageFormat "
            f"numbering (lvlText={lvl_text!r}; expected a pattern like {expected!r})."
        )


def _validate_complete_article_hierarchy(
    role_specs: Dict[str, Dict[str, Any]],
    roles_in_target: set[str],
) -> None:
    """Require one coherent multilevel list when articles are converted."""

    if "PART" not in roles_in_target:
        raise EngineError("canadian_target_hierarchy", 
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
            raise EngineError("canadian_architect_contract", 
                f"Architect template: role {role} has ilvl={ilvl!r}, "
                f"lvlText={lvl_text!r}; Canadian article conversion requires a "
                "coherent PART/article/list hierarchy."
            )
        if not num_id:
            raise EngineError("canadian_architect_contract", 
                f"Architect template: role {role} is missing its Word numbering list "
                "identifier."
            )
        if reference_num_id is None:
            reference_num_id = num_id
        elif num_id != reference_num_id:
            raise EngineError("canadian_architect_contract", 
                "Architect template: Canadian PART, article, and subordinate roles "
                "must share one Word multilevel numbering list."
            )


_ROLE_LEVEL = ROLE_LEVEL

#: Roles a designer counts as headings when locating a paragraph in Word.
_LOCATOR_HEADING_ROLES = frozenset(NUMBERED_ROLES) | {"PART"}


def _validate_architect_numbering(
    numbering_xml: str,
    role_specs: Dict[str, Dict[str, Any]],
    roles: set[str],
) -> None:
    if not numbering_xml.strip():
        raise EngineError("canadian_architect_contract", 
            "Architect template: numbering.xml is missing; Canadian conversion "
            "requires the architect's numbering definitions."
        )
    root = parse_untrusted_xml(
        prepare_xml_text_for_utf8(numbering_xml),
        "architect numbering.xml",
    )
    for role in sorted(roles):
        spec = role_specs[role]
        pattern = spec["numbering_pattern"]
        num_id = str(pattern.get("numId") or "")
        ilvl = str(pattern.get("ilvl") or "0")
        level, override = _find_numbering_level(root, num_id, ilvl)
        _validate_numbering_start(
            level,
            override,
            context=f"Architect template: role {role} numbering",
            reject_override=False,
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
        raise EngineError("canadian_target_hierarchy", 
            "Canadian conversion requires a strict current architect template profile."
        )
    items = classifications.get("classifications")
    if not isinstance(items, list):
        raise EngineError("canadian_target_hierarchy", "Canadian conversion requires final paragraph classifications")

    role_by_index: Dict[int, str] = {}
    for item in items:
        if not isinstance(item, dict):
            raise EngineError("canadian_target_hierarchy", "Canadian conversion classification entries must be objects")
        index = item.get("paragraph_index")
        role = item.get("csi_role")
        if not isinstance(index, int) or index < 0 or not isinstance(role, str):
            raise EngineError("canadian_target_hierarchy", "Canadian conversion received an invalid classification entry")
        if index in role_by_index:
            raise EngineError("canadian_target_hierarchy", f"Duplicate Canadian conversion classification for paragraph {index}")
        role_by_index[index] = role

    blocks = list(iter_paragraph_xml_blocks(document_xml))
    warnings: list[ConversionIssue] = []

    # The Phase 2 classifier intentionally gives every content paragraph a
    # semantic role.  In Canadian mode, a numbered role is actionable only
    # when the source also proves that the paragraph is a list item.  Preserve
    # markerless guesses unchanged and keep them out of the later numbered
    # style application; otherwise the architect style itself would silently
    # create numbering that was not present in the source.
    numbered_candidates = set(NUMBERED_ROLES)
    part_spec = role_specs.get("PART")
    if isinstance(part_spec, dict) and part_spec.get(
        "numbering_provenance"
    ) in {"style_numpr", "direct_numpr"}:
        numbered_candidates.add("PART")

    preserved_indices: set[int] = set()
    for index, role in sorted(role_by_index.items()):
        if role not in numbered_candidates:
            continue
        if index >= len(blocks):
            raise EngineError("canadian_target_hierarchy", f"Canadian conversion paragraph index is out of range: {index}")
        paragraph = blocks[index][2]
        text = paragraph_text_from_block(paragraph)
        literal = _detect_literal_marker(text, role)
        any_literal = _detect_any_literal_marker(text)
        automatic = _effective_numpr(
            strip_out_of_scope_subtrees(paragraph),
            styles_xml,
        ) is not None
        if literal is None and any_literal is None and not automatic:
            preserved_indices.add(index)
            warnings.append(
                ConversionIssue(
                    index,
                    PRESERVED_UNNUMBERED_ROLE,
                    f"Classified as numbered role {role}, but the source has neither "
                    "a recognized typed marker nor Word automatic numbering. The "
                    "paragraph was preserved unchanged and excluded from numbered "
                    "architect style application.",
                    text.strip()[:120],
                )
            )

    effective_role_by_index = {
        index: role
        for index, role in role_by_index.items()
        if index not in preserved_indices
    }
    locate = _paragraph_locator(blocks, effective_role_by_index)
    roles_in_target = set(effective_role_by_index.values())
    used_numbered_roles = sorted(roles_in_target & NUMBERED_ROLES)
    for role in used_numbered_roles:
        _validate_canadian_role_contract(role, role_specs.get(role))
    if "ARTICLE" in roles_in_target:
        _validate_complete_article_hierarchy(role_specs, roles_in_target)

    roles_to_convert = set(used_numbered_roles)
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
        parse_untrusted_xml(
            prepare_xml_text_for_utf8(numbering_xml),
            "architect numbering.xml",
        )
        if numbering_xml.strip()
        else None
    )

    replacements: Dict[int, str] = {}
    evidence: list[_SourceEvidence] = []
    edits: list[MarkerEdit] = []
    literal_removed = 0
    automatic_retargeted = 0

    for index, role in sorted(effective_role_by_index.items()):
        if role not in roles_to_convert:
            continue
        if index >= len(blocks):
            raise EngineError("canadian_target_hierarchy", f"Canadian conversion paragraph index is out of range: {index}")
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
            raise EngineError("canadian_target_hierarchy", 
                f"Paragraph {index}{locate(index)} is classified as {role} but starts "
                "with incompatible "
                f"marker {any_literal!r}.",
                locate.at(index),
            )
        if literal is not None and automatic:
            raise EngineError("canadian_target_hierarchy", 
                f"Paragraph {index}{locate(index)} has both automatic numbering and "
                "typed marker "
                f"{literal.marker!r}; remove the doubled numbering before conversion.",
                locate.at(index),
            )

        if literal is not None:
            delimiter = _marker_markup_delimiter(analysis, role)
            if delimiter == "line_break":
                raise EngineError("canadian_target_hierarchy", 
                    f"Paragraph {index}{locate(index)} uses a line break after typed marker "
                    f"{literal.marker!r}; only a normal space or Word list tab is safe.",
                    locate.at(index),
                )
            # A two-component article marker such as ``1.1`` is proven later
            # against its active PART and contiguous article sequence.  The
            # single-component ``.1`` family remains ambiguous with decimal
            # values and therefore still requires an actual Word list tab.
            if literal.family == "csc_dot_decimal" and delimiter != "tab":
                raise EngineError("canadian_target_hierarchy", 
                    f"Paragraph {index}{locate(index)} begins with ambiguous decimal text "
                    f"{literal.marker!r}. A typed Canadian marker is converted only "
                    "when followed by a structural Word tab.",
                    locate.at(index),
                )
            if (
                literal.family == "csc_article"
                and delimiter != "tab"
                and not _has_heading_like_article_body(literal.body_text)
            ):
                raise EngineError("canadian_target_hierarchy", 
                    f"Paragraph {index}{locate(index)} begins with ambiguous decimal text "
                    f"{literal.marker!r}. A spaced Canadian article marker must "
                    "be followed by heading-like text or a structural Word tab.",
                    locate.at(index),
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
            raise EngineError("canadian_target_hierarchy", 
                f"Paragraph {index}{locate(index)} is classified as numbered role {role}, "
                "but it has "
                "neither a recognized typed marker nor Word automatic numbering. "
                "Canadian conversion will not insert an unproven list item.",
                locate.at(index),
            )

    automatic_ids: Dict[str, set[str]] = {}
    automatic_by_index: Dict[int, _SourceEvidence] = {}
    for item in evidence:
        if item.source_kind != "automatic":
            continue
        _validate_automatic_source(
            item,
            numbering_root,
            numbering_catalog,
            set(role_specs),
            describe=locate,
        )
        assert item.automatic_numpr is not None
        automatic_ids.setdefault(item.role, set()).add(item.automatic_numpr["numId"])
        automatic_by_index[item.paragraph_index] = item
    for role, num_ids in automatic_ids.items():
        if len(num_ids) > 1:
            raise EngineError("canadian_target_hierarchy", 
                f"Automatic {role} numbering changes list instances ({sorted(num_ids)}); "
                "Canadian conversion cannot prove that the restart will be preserved."
            )
    converted_num_ids = set().union(*automatic_ids.values()) if automatic_ids else set()
    if len(converted_num_ids) > 1:
        raise EngineError("canadian_target_hierarchy", 
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
            raise EngineError("canadian_target_hierarchy", 
                f"Unconverted paragraph {index} shares automatic source list "
                f"numId={effective_numpr.get('numId')!r}; removing it would change "
                f"following counters. It is paragraph {index}{locate(index)}.",
                locate.at(index),
            )
    _validate_source_sequence(evidence, describe=locate)

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
    parse_untrusted_xml(
        prepare_xml_text_for_utf8(converted_document),
        "word/document.xml (converted)",
    )

    report = CanadianConversionReport(
        paragraphs_examined=sum(
            1 for role in effective_role_by_index.values() if role in roles_to_convert
        ),
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
    "classifications_for_canadian_application",
    "plan_csi_to_canadian",
    "validate_conversion_mode",
]
