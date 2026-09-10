"""Fail-closed Canadian PageFormat -> CSI hierarchy conversion.

The inverse of :mod:`csi_to_canadian`. Where the forward converter strips a
typed marker and lets an automatic Word list render the number, this one goes
the other way: it works out what number each paragraph *currently* shows, then
writes that number into the paragraph as literal text and takes the automatic
numbering away.

Typed markers are the deliberate output, not a shortcut. A US specification is
normally issued with ``PART 1``, ``1.1``, ``A.``, ``1.``, ``a.`` typed into the
text, and a document whose numbering is literal renders identically everywhere,
survives being pasted into another editor, and can be proven paragraph by
paragraph. Automatic numbering in the CSI scheme would re-introduce exactly the
list-instance fragility the forward converter spends most of its length
refusing to guess at.

**Writing a number is harder than removing one.** Removing a typed marker only
needs the marker; writing one needs the counter Word would have rendered, which
means walking the list. That walk is only sound inside the same fence the
forward converter builds, so this module reuses it: every converted paragraph
must sit on one list instance, that instance must start at 1 with no
``lvlRestart`` and no level override, and every paragraph on it must be
converted. Under those conditions a counter is a plain per-level tally --
increment this level, reset the deeper ones -- and anything outside them fails
closed rather than publishing a document whose numbers are a guess.
"""

from __future__ import annotations

import re
from pathlib import Path
from typing import Any, Callable, Dict, List

from spec_formatter.role_contract import (
    BODY_HIERARCHY_ROLES,
    ROLE_LEVEL,
    ROLE_ORDER,
)

from .classification import (
    _build_numbering_catalog,
    _effective_numpr,
    _resolve_numbering_pattern,
)
from .csi_to_canadian import (
    CanadianConversionReport,
    ConversionIssue,
    ConversionPlan,
    MarkerEdit,
    PRESERVED_UNNUMBERED_ROLE,
)
from .errors import EngineError
from .marker_tools import (
    _TRACKED_OR_FIELD_RX,
    _detect_any_literal_marker,
    _detect_literal_marker,
    _find_numbering_level,
    _paragraph_locator,
    _remove_literal_marker,
    _SourceEvidence,
    _validate_automatic_source,
    _validate_numbering_start,
    _validate_source_sequence,
)
from .ooxml_text import prepare_xml_text_for_utf8, read_xml_text, write_xml_text
from .sectpr_tools import extract_all_sectpr_blocks
from .untrusted_xml import parse_untrusted_xml
from .xml_helpers import (
    edit_preserving_out_of_scope_subtrees,
    iter_paragraph_xml_blocks,
    paragraph_text_from_block,
    strip_out_of_scope_subtrees,
)


_HIERARCHY = "canadian_to_csi_hierarchy"
_UNPROVABLE = "canadian_to_csi_numbering_unprovable"

#: Roles this converter writes a marker for. ``PART`` is included
#: unconditionally, unlike in the forward direction: a CSI ``PART 1`` heading
#: is typed text in every template, so there is no architect contract to
#: consult about whether it should be.
_CONVERTIBLE_ROLES = frozenset(BODY_HIERARCHY_ROLES)

_PPR_RX = re.compile(r"<w:pPr\b[^>]*(?:/>|>.*?</w:pPr>)", re.S)
_NUMPR_RX = re.compile(r"<w:numPr\b[^>]*(?:/>|>.*?</w:numPr>)", re.S)
#: Word's "this paragraph is not in a list" numbering reference. Written as a
#: direct property so it also cancels numbering inherited from a style, which
#: simply deleting a direct ``numPr`` would not.
_NUMBERING_OFF = '<w:numPr><w:ilvl w:val="0"/><w:numId w:val="0"/></w:numPr>'
_FIRST_TEXT_RX = re.compile(r"<w:t\b[^>]*>", re.S)


def _escape(value: str) -> str:
    """Escape marker text for an XML text node.

    Markers are generated from integers and fixed punctuation, so nothing here
    needs escaping today. It is done anyway because a text node is a text node,
    and the next person to widen the marker vocabulary should not have to
    notice this.
    """

    return value.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")


def _int_to_alpha(value: int) -> str:
    """Render 1 -> ``a`` ... 26 -> ``z``, refusing anything past ``z``.

    Word renders a 27th alphabetic item as ``aa`` and a 28th as ``bb``, but the
    marker vocabulary both converters share (``_ROLE_MARKERS``) only ever
    matches a *single* letter. Writing ``aa.`` would therefore produce a
    document this application could no longer read back as a marker at all, so
    a level that runs past ``z`` fails closed instead. Twenty-seven items at
    one alphabetic level is already outside what the CSI scheme expresses.
    """

    if not 1 <= value <= 26:
        raise EngineError(
            _UNPROVABLE,
            f"An alphabetic CSI level reached item {value}; only a. through z. "
            "can be written as a single-letter marker.",
        )
    return chr(ord("a") + value - 1)


def _csi_marker(role: str, counters: Dict[int, int]) -> str:
    """Render the typed CSI marker for *role* from the active counter stack."""

    level = ROLE_LEVEL[role]
    counter = counters[level]
    if role == "PART":
        return f"PART {counter}"
    if role == "ARTICLE":
        return f"{counters[ROLE_LEVEL['PART']]}.{counter}"
    if role == "PARAGRAPH":
        return f"{_int_to_alpha(counter).upper()}."
    if role == "SUBPARAGRAPH":
        return f"{counter}."
    if role == "SUBSUBPARAGRAPH":
        return f"{_int_to_alpha(counter)}."
    if role == "SUBPARAGRAPH_LEVEL_5":
        return f"{counter})"
    if role == "SUBPARAGRAPH_LEVEL_6":
        return f"{_int_to_alpha(counter)})"
    if role == "SUBPARAGRAPH_LEVEL_7":
        return f"({counter})"
    if role == "SUBPARAGRAPH_LEVEL_8":
        return f"({_int_to_alpha(counter)})"
    raise AssertionError(f"Unhandled convertible role: {role}")


#: The ``CT_PPr`` children that must precede ``w:numPr`` in schema order.
#: ``w:numPr`` written before ``w:pStyle`` is invalid OOXML even though Word
#: often renders it anyway, and a stricter consumer is entitled to reject the
#: file, so the insertion point is chosen rather than assumed to be the front.
_PPR_BEFORE_NUMPR = (
    "w:pStyle",
    "w:keepNext",
    "w:keepLines",
    "w:pageBreakBefore",
    "w:framePr",
    "w:widowControl",
)


def _numpr_insertion_point(ppr_inner: str) -> int:
    """Offset inside a ``w:pPr`` body where ``w:numPr`` may legally be added."""

    offset = 0
    for name in _PPR_BEFORE_NUMPR:
        for match in re.finditer(rf"<{name}\b[^>]*(?:/>|>.*?</{name}>)", ppr_inner, re.S):
            offset = max(offset, match.end())
    return offset


def _suppress_automatic_numbering(paragraph_xml: str) -> str:
    """Return *paragraph_xml* with its effective list membership cancelled."""

    match = _PPR_RX.search(paragraph_xml)
    if match is None:
        insert_at = paragraph_xml.index(">") + 1
        return (
            paragraph_xml[:insert_at]
            + f"<w:pPr>{_NUMBERING_OFF}</w:pPr>"
            + paragraph_xml[insert_at:]
        )
    ppr = match.group(0)
    if ppr.endswith("/>"):
        rebuilt = ppr[:-2] + ">" + _NUMBERING_OFF + "</w:pPr>"
    else:
        without = _NUMPR_RX.sub("", ppr, count=1)
        open_end = without.index(">") + 1
        inner = without[open_end : without.rindex("</w:pPr>")]
        cut = open_end + _numpr_insertion_point(inner)
        rebuilt = without[:cut] + _NUMBERING_OFF + without[cut:]
    return paragraph_xml[: match.start()] + rebuilt + paragraph_xml[match.end():]


def _insert_marker(paragraph_xml: str, marker: str) -> str:
    """Prepend *marker* and a tab inside the paragraph's first text run.

    The marker goes *into* the existing run rather than into new runs of its
    own, for two reasons. It inherits that run's character formatting, so a
    bold heading gets a bold number instead of a bare one in the document
    default; and the paragraph's run structure is unchanged, which is what the
    run-property invariant in ``phase2_invariants`` checks. Adding runs would
    trip that check for a change that loses no formatting at all -- the right
    answer is not to widen the invariant but not to add the runs.

    A run may hold several ``w:t`` and ``w:tab`` children, so the result is
    ordinary OOXML.
    """

    def _edit(unprotected: str) -> str:
        match = _FIRST_TEXT_RX.search(unprotected)
        if match is None:
            raise EngineError(
                _HIERARCHY,
                f"A paragraph classified for CSI marker {marker!r} has no text run "
                "to carry it.",
            )
        # The numbering suppression goes on ``w:pPr``, outside any field or
        # revision subtree. If the marker went *inside* one, updating the field
        # or rejecting the insertion would delete the marker and leave the
        # paragraph with no number at all -- the automatic numbering that used
        # to supply it is gone by then. The forward converter refuses the same
        # markup for the mirror-image reason.
        if _TRACKED_OR_FIELD_RX.search(unprotected[: match.start()]):
            raise EngineError(
                _HIERARCHY,
                "A paragraph's leading text is inside a tracked change or field "
                f"result, so CSI marker {marker!r} could not be written where it "
                "would survive. Accept the changes or unlink the field first.",
            )
        prefix = f'<w:t xml:space="preserve">{_escape(marker)}</w:t><w:tab/>'
        return unprotected[: match.start()] + prefix + unprotected[match.start():]

    return edit_preserving_out_of_scope_subtrees(paragraph_xml, _edit)


def _verify_marked_paragraph(
    index: int,
    after: str,
    marker: str,
    expected_body: str,
    *,
    describe: Callable[[int], str],
) -> None:
    """Prove the edit added the marker and changed nothing else in the text."""

    actual = paragraph_text_from_block(after)
    if not actual.startswith(marker):
        raise EngineError(
            _HIERARCHY,
            f"Paragraph {index}{describe(index)} did not receive its CSI marker "
            f"{marker!r} as leading text.",
        )
    remainder = actual[len(marker):].lstrip(" \t ")
    if remainder != expected_body.lstrip(" \t "):
        raise EngineError(
            _HIERARCHY,
            f"Paragraph {index}{describe(index)} changed beyond its numbering "
            "marker; the conversion was withheld.",
        )


def _validate_converted_list(
    numbering_xml: str,
    num_ids: set[str],
    ilvls: Dict[str, set[str]],
) -> None:
    """Prove every converted list level has a countable, unoverridden counter."""

    if not num_ids:
        return
    if not numbering_xml.strip():
        raise EngineError(
            _UNPROVABLE,
            "Paragraphs use automatic numbering but the target has no "
            "numbering.xml, so their numbers cannot be read.",
        )
    root = parse_untrusted_xml(
        prepare_xml_text_for_utf8(numbering_xml),
        "word/numbering.xml",
    )
    for num_id in sorted(num_ids):
        for ilvl in sorted(ilvls.get(num_id, set())):
            level, override = _find_numbering_level(root, num_id, ilvl)
            # ``reject_override=True``: a level override changes the counter in
            # a way this walker does not model, and a wrong number written as
            # literal text is unrecoverable for the reader.
            _validate_numbering_start(
                level,
                override,
                context=f"Target numbering numId={num_id} ilvl={ilvl}",
                reject_override=True,
            )


def _advance(counters: Dict[int, int], level: int) -> None:
    """Increment *level* and reset every deeper level, as Word would."""

    counters[level] = counters.get(level, 0) + 1
    for deeper in list(counters):
        if deeper > level:
            del counters[deeper]


def plan_canadian_to_csi(
    document_xml: str,
    styles_xml: str,
    classifications: Dict[str, Any],
    *,
    numbering_xml: str = "",
) -> ConversionPlan:
    """Validate and build the complete document edit before writing anything."""

    items = classifications.get("classifications")
    if not isinstance(items, list):
        raise EngineError(_HIERARCHY, "CSI conversion requires final paragraph classifications")

    role_by_index: Dict[int, str] = {}
    for item in items:
        if not isinstance(item, dict):
            raise EngineError(_HIERARCHY, "CSI conversion classification entries must be objects")
        index = item.get("paragraph_index")
        role = item.get("csi_role")
        if not isinstance(index, int) or index < 0 or not isinstance(role, str):
            raise EngineError(_HIERARCHY, "CSI conversion received an invalid classification entry")
        if index in role_by_index:
            raise EngineError(_HIERARCHY, f"Duplicate CSI conversion classification for paragraph {index}")
        role_by_index[index] = role

    blocks = list(iter_paragraph_xml_blocks(document_xml))
    warnings: List[ConversionIssue] = []
    numbering_catalog = _build_numbering_catalog(numbering_xml)

    # A numbered role with no number in the source is a semantic guess, not a
    # list item. The forward converter preserves those untouched and so does
    # this one: inventing "A." for a paragraph that never carried a marker
    # would put a number in the reader's document that nobody wrote.
    preserved_indices: set[int] = set()
    for index, role in sorted(role_by_index.items()):
        if role not in _CONVERTIBLE_ROLES:
            continue
        if index >= len(blocks):
            raise EngineError(_HIERARCHY, f"CSI conversion paragraph index is out of range: {index}")
        paragraph = blocks[index][2]
        text = paragraph_text_from_block(paragraph)
        automatic = _effective_numpr(strip_out_of_scope_subtrees(paragraph), styles_xml)
        if (
            _detect_literal_marker(text, role) is None
            and _detect_any_literal_marker(text) is None
            and automatic is None
        ):
            preserved_indices.add(index)
            warnings.append(
                ConversionIssue(
                    index,
                    PRESERVED_UNNUMBERED_ROLE,
                    f"Classified as numbered role {role}, but the source has neither "
                    "a recognized typed marker nor Word automatic numbering. The "
                    "paragraph was preserved unchanged and no CSI marker was "
                    "invented for it.",
                    text.strip()[:120],
                )
            )

    effective_role_by_index = {
        index: role
        for index, role in role_by_index.items()
        if index not in preserved_indices
    }
    locate = _paragraph_locator(blocks, effective_role_by_index)

    # --- Pass 1: gather evidence and prove the source is countable ----------
    evidence: List[_SourceEvidence] = []
    automatic_indices: Dict[int, Dict[str, str]] = {}
    literal_indices: Dict[int, Any] = {}
    num_ids: set[str] = set()
    ilvls: Dict[str, set[str]] = {}

    for index, role in sorted(effective_role_by_index.items()):
        if role not in _CONVERTIBLE_ROLES:
            continue
        paragraph = blocks[index][2]
        text = paragraph_text_from_block(paragraph)
        literal = _detect_literal_marker(text, role)
        any_literal = _detect_any_literal_marker(text)
        analysis = strip_out_of_scope_subtrees(paragraph)
        automatic_numpr = _effective_numpr(analysis, styles_xml)

        if literal is None and any_literal is not None:
            raise EngineError(
                _HIERARCHY,
                f"Paragraph {index}{locate(index)} is classified as {role} but starts "
                f"with incompatible marker {any_literal!r}.",
            )
        if literal is not None and automatic_numpr is not None:
            raise EngineError(
                _HIERARCHY,
                f"Paragraph {index}{locate(index)} has both automatic numbering and "
                f"typed marker {literal.marker!r}; remove the doubled numbering "
                "before conversion.",
            )

        if literal is not None:
            literal_indices[index] = literal
            evidence.append(_SourceEvidence(index, role, "literal", literal, None, None))
        elif automatic_numpr is not None:
            num_id = str(automatic_numpr["numId"])
            ilvl = str(automatic_numpr.get("ilvl", "0"))
            automatic_indices[index] = {"numId": num_id, "ilvl": ilvl}
            num_ids.add(num_id)
            ilvls.setdefault(num_id, set()).add(ilvl)
            evidence.append(
                _SourceEvidence(
                    index,
                    role,
                    "automatic",
                    None,
                    automatic_numpr,
                    _resolve_numbering_pattern(automatic_numpr, numbering_catalog),
                )
            )
        else:  # pragma: no cover - the preserve pass already removed these
            raise EngineError(
                _HIERARCHY,
                f"Paragraph {index}{locate(index)} is classified as numbered role "
                f"{role} but carries no number to convert.",
            )

    if len(num_ids) > 1:
        raise EngineError(
            _HIERARCHY,
            "Converted paragraphs span more than one Word list instance "
            f"({sorted(num_ids)}); their counters cannot be proven together.",
        )
    _validate_converted_list(numbering_xml, num_ids, ilvls)

    # The counter walk below is driven by the *classified* role's level, so a
    # role that disagrees with the level Word is actually rendering would
    # produce a number the document never showed -- writing "1.1" onto what
    # Word renders as "PART 2", and writing it as permanent text. Prove the two
    # agree first, with the same check the forward converter makes.
    numbering_root = (
        parse_untrusted_xml(
            prepare_xml_text_for_utf8(numbering_xml),
            "word/numbering.xml",
        )
        if numbering_xml.strip()
        else None
    )
    for item in evidence:
        if item.source_kind != "automatic":
            continue
        _validate_automatic_source(
            item,
            numbering_root,
            numbering_catalog,
            set(ROLE_ORDER),
            describe=locate,
            error_code=_UNPROVABLE,
        )
        expected_ilvl = str(ROLE_LEVEL[item.role])
        actual_ilvl = str((item.automatic_numpr or {}).get("ilvl", "0"))
        if actual_ilvl != expected_ilvl:
            raise EngineError(
                _UNPROVABLE,
                f"Paragraph {item.paragraph_index}{locate(item.paragraph_index)} is "
                f"classified as {item.role} (list level {expected_ilvl}) but sits at "
                f"list level {actual_ilvl}; the CSI marker would not match the "
                "number the document currently shows.",
            )

    # Every paragraph on a converted list must be converted. One left behind
    # keeps its automatic number while its neighbours get literal ones, and the
    # two sequences drift apart from that point on.
    for index, (_start, _end, block) in enumerate(blocks):
        effective = _effective_numpr(strip_out_of_scope_subtrees(block), styles_xml)
        if effective is None or str(effective.get("numId")) not in num_ids:
            continue
        if index not in automatic_indices:
            raise EngineError(
                _HIERARCHY,
                f"Unconverted paragraph {index}{locate(index)} shares automatic list "
                f"numId={effective.get('numId')!r} with converted paragraphs; "
                "leaving it numbered would desynchronise the sequence.",
            )

    _validate_source_sequence(evidence, describe=locate, error_code=_HIERARCHY)

    # --- Pass 2: walk the counters and build the edits ----------------------
    counters: Dict[int, int] = {}
    replacements: Dict[int, str] = {}
    edits: List[MarkerEdit] = []
    literal_removed = 0
    automatic_converted = 0

    for index, role in sorted(effective_role_by_index.items()):
        if role not in _CONVERTIBLE_ROLES:
            continue
        level = ROLE_LEVEL[role]
        _advance(counters, level)
        if role == "ARTICLE" and ROLE_LEVEL["PART"] not in counters:
            raise EngineError(
                _HIERARCHY,
                f"Paragraph {index}{locate(index)} is an ARTICLE with no preceding "
                "PART, so its CSI article number cannot be formed.",
            )
        marker = _csi_marker(role, counters)
        paragraph = blocks[index][2]

        literal = literal_indices.get(index)
        if literal is not None:
            stripped, _tab_removed = _remove_literal_marker(paragraph, role)
            body = literal.body_text
            source_kind = "literal"
            literal_removed += 1
        else:
            stripped = _suppress_automatic_numbering(paragraph)
            body = paragraph_text_from_block(paragraph)
            source_kind = "automatic"
            automatic_converted += 1

        converted = _insert_marker(stripped, marker)
        _verify_marked_paragraph(index, converted, marker, body, describe=locate)
        replacements[index] = converted
        edits.append(
            MarkerEdit(
                index,
                role,
                source_kind,
                "literal",
                literal.marker if literal is not None else None,
                marker,
            )
        )

    # --- Assemble and re-prove the whole document ---------------------------
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
        raise RuntimeError("CSI conversion invariant failed: paragraph count changed")
    for index, (_start, _end, before) in enumerate(blocks):
        if index not in replacements and after_blocks[index][2] != before:
            raise RuntimeError(
                f"CSI conversion invariant failed: untouched paragraph {index} changed"
            )
    if extract_all_sectpr_blocks(document_xml) != extract_all_sectpr_blocks(converted_document):
        raise RuntimeError("CSI conversion invariant failed: section properties changed")
    parse_untrusted_xml(
        prepare_xml_text_for_utf8(converted_document),
        "word/document.xml (converted)",
    )

    report = CanadianConversionReport(
        paragraphs_examined=sum(
            1 for role in effective_role_by_index.values() if role in _CONVERTIBLE_ROLES
        ),
        paragraphs_converted=len(edits),
        literal_markers_removed=literal_removed,
        automatic_numbering_retargeted=automatic_converted,
        unnumbered_paragraphs_numbered=0,
        edits=tuple(edits),
        warnings=tuple(warnings),
    )
    return ConversionPlan(converted_document, report)


def apply_canadian_to_csi(
    extract_dir: Path,
    classifications: Dict[str, Any],
    log: List[str],
) -> CanadianConversionReport:
    """Apply a validated Canadian-to-CSI conversion to one extracted DOCX."""

    document_path = Path(extract_dir) / "word" / "document.xml"
    styles_path = Path(extract_dir) / "word" / "styles.xml"
    numbering_path = Path(extract_dir) / "word" / "numbering.xml"
    plan = plan_canadian_to_csi(
        read_xml_text(document_path),
        read_xml_text(styles_path),
        classifications,
        numbering_xml=(
            read_xml_text(numbering_path) if numbering_path.is_file() else ""
        ),
    )
    write_xml_text(document_path, plan.document_xml)
    report = plan.report
    log.append(
        "CSI conversion: "
        f"{report.paragraphs_converted} numbered paragraphs; "
        f"rewrote {report.literal_markers_removed} typed markers; "
        f"converted {report.automatic_numbering_retargeted} automatic paragraphs "
        "to typed CSI markers"
    )
    for issue in report.warnings:
        log.append(f"CSI conversion warning p[{issue.paragraph_index}]: {issue.message}")
    return report


__all__ = [
    "apply_canadian_to_csi",
    "plan_canadian_to_csi",
]
