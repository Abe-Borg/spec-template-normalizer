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
import xml.etree.ElementTree as ET
from datetime import datetime, timezone
from pathlib import Path
from typing import Any, Callable, Dict, List, Optional, Set, Tuple

from spec_formatter.role_contract import (
    BODY_HIERARCHY_ROLES,
    ROLE_LEVEL,
    ROLE_ORDER,
)

from .classification import (
    _build_numbering_catalog,
    _effective_numpr,
    _resolve_numbering_pattern,
    _style_replacement_ppr_properties,
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
    _wq,
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
_TRACKED = "canadian_to_csi_tracked_hierarchy"

#: Roles this converter writes a marker for. ``PART`` is included
#: unconditionally, unlike in the forward direction: a CSI ``PART 1`` heading
#: is typed text in every template, so there is no architect contract to
#: consult about whether it should be.
_CONVERTIBLE_ROLES = frozenset(BODY_HIERARCHY_ROLES)

_PPR_RX = re.compile(r"<w:pPr\b[^>]*(?:/>|>.*?</w:pPr>)", re.S)
_NUMPR_RX = re.compile(r"<w:numPr\b[^>]*(?:/>|>.*?</w:numPr>)", re.S)
_FIRST_TEXT_RX = re.compile(r"<w:t\b[^>]*>", re.S)



_MARK_RPR_RX = re.compile(r"<w:rPr\b[^>]*(?:/>|>(.*?)</w:rPr>)", re.S)


def _paragraph_mark_revision(paragraph_xml: str) -> Optional[str]:
    """Return ``"insertion"``/``"deletion"`` if the paragraph *mark* is tracked.

    Only ``w:pPr/w:rPr`` counts. A ``w:ins`` anywhere else in the paragraph
    marks inserted *text*, which leaves the paragraph -- and therefore the
    sequence -- intact under both accept and reject, so the search is bounded
    to the ``w:pPr`` element rather than run against the whole paragraph.
    """

    ppr = _PPR_RX.search(paragraph_xml)
    if ppr is None:
        return None
    mark = _MARK_RPR_RX.search(ppr.group(0))
    if mark is None or not mark.group(1):
        return None
    if re.search(r"<w:ins\b", mark.group(1)):
        return "insertion"
    if re.search(r"<w:del\b", mark.group(1)):
        return "deletion"
    return None



#: Revision ids must be unique within the document. The source's own ids are
#: Word's, in the low thousands at most; starting well above that keeps the
#: application's insertions from colliding with them without needing to scan.
_MARKER_REVISION_ID_BASE = 900000


def _utc_revision_date() -> str:
    return datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")


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


#: The complete ``CT_PPr`` child sequence, in schema order, from
#: ``ISO-IEC29500-4_2016/wml.xsd`` (``CT_PPrBase`` supplies 1-33 and ``CT_PPr``
#: appends the final three).
#:
#: This is deliberately the *whole* table rather than a prefix trimmed to the
#: elements written today. An abbreviated order table is only correct for the
#: one element it was abbreviated for: a six-entry "everything before
#: ``w:numPr``" list silently places ``w:ind`` immediately after ``w:pStyle``,
#: which is invalid even where Word renders it, and only an XSD check catches
#: it. Insert through :func:`_ppr_insertion_point` and this stays true for any
#: element a later change needs to write.
_PPR_CHILD_ORDER = (
    "w:pStyle",
    "w:keepNext",
    "w:keepLines",
    "w:pageBreakBefore",
    "w:framePr",
    "w:widowControl",
    "w:numPr",
    "w:suppressLineNumbers",
    "w:pBdr",
    "w:shd",
    "w:tabs",
    "w:suppressAutoHyphens",
    "w:kinsoku",
    "w:wordWrap",
    "w:overflowPunct",
    "w:topLinePunct",
    "w:autoSpaceDE",
    "w:autoSpaceDN",
    "w:bidi",
    "w:adjustRightInd",
    "w:snapToGrid",
    "w:spacing",
    "w:ind",
    "w:contextualSpacing",
    "w:mirrorIndents",
    "w:suppressOverlap",
    "w:jc",
    "w:textDirection",
    "w:textAlignment",
    "w:textboxTightWrap",
    "w:outlineLvl",
    "w:divId",
    "w:cnfStyle",
    "w:rPr",
    "w:sectPr",
    "w:pPrChange",
)


def _ppr_insertion_point(ppr_inner: str, element_name: str) -> int:
    """Offset inside a ``w:pPr`` body where *element_name* may legally be added.

    The offset is the end of the last present child that must precede
    *element_name*, so the new element lands after its predecessors and before
    everything that must follow it.
    """

    try:
        position = _PPR_CHILD_ORDER.index(element_name)
    except ValueError:  # pragma: no cover - guarded by the caller's constants
        raise AssertionError(f"Unknown w:pPr child: {element_name}") from None
    offset = 0
    for name in _PPR_CHILD_ORDER[:position]:
        for match in re.finditer(rf"<{name}\b[^>]*(?:/>|>.*?</{name}>)", ppr_inner, re.S):
            offset = max(offset, match.end())
    return offset


def _numbering_off(ilvl: str) -> str:
    """Word's "this paragraph is not in a list" reference at *ilvl*.

    The level is carried through rather than flattened to ``0``. With
    ``numId=0`` the value does not render, but a paragraph that states level 0
    while sitting at level 3 is simply describing itself falsely to the next
    reader of the file -- including this application's own converters.
    """

    return f'<w:numPr><w:ilvl w:val="{ilvl}"/><w:numId w:val="0"/></w:numPr>'


def _suppress_automatic_numbering(
    paragraph_xml: str,
    ilvl: str = "0",
    level_geometry: Tuple[str, ...] = (),
) -> str:
    """Return *paragraph_xml* with its effective list membership cancelled.

    Cancelling the list is only half the edit. A numbering level's ``w:pPr``
    -- its ``w:ind`` and its ``w:tabs`` num stop -- applies *only* while the
    paragraph is a member of that list, so ``numId=0`` discards the paragraph's
    indentation along with its number whenever the geometry lived in
    ``numbering.xml`` rather than in the style. Templates written that way are
    normal, not exotic: a ``Cdn*``-style stylesheet carries ``w:numPr`` and no
    ``w:ind`` at all, so *every* indent in the document comes from the level.
    Dropping it flattens the whole outline into one column while leaving the
    text, the numbers and the run structure provably intact -- which is exactly
    why no text-level check can see it happen.

    So the level's geometry is materialized onto the paragraph in the same
    edit, which is what Word itself writes when a user turns numbering off on
    one paragraph by hand. A property the paragraph already sets directly is
    left alone: it already outranks the level and is the author's own choice.

    The caller decides *whether* there is anything to restore -- see
    :func:`_restorable_level_geometry`. This function only places what it is
    given.
    """

    numbering_off = _numbering_off(ilvl)
    match = _PPR_RX.search(paragraph_xml)
    if match is None:
        inserted = numbering_off + "".join(level_geometry)
        insert_at = paragraph_xml.index(">") + 1
        return (
            paragraph_xml[:insert_at]
            + f"<w:pPr>{inserted}</w:pPr>"
            + paragraph_xml[insert_at:]
        )
    ppr = match.group(0)
    if ppr.endswith("/>"):
        inserted = numbering_off + "".join(level_geometry)
        rebuilt = ppr[:-2] + ">" + inserted + "</w:pPr>"
    else:
        rebuilt = _NUMPR_RX.sub("", ppr, count=1)
        for fragment in (numbering_off,) + tuple(level_geometry):
            name = re.match(r"<(w:\w+)", fragment).group(1)
            open_end = rebuilt.index(">") + 1
            inner = rebuilt[open_end : rebuilt.rindex("</w:pPr>")]
            if re.search(rf"<{name}\b", inner):
                # Already set directly on the paragraph: the author's own
                # value wins over the numbering level's.
                continue
            cut = open_end + _ppr_insertion_point(inner, name)
            rebuilt = rebuilt[:cut] + fragment + rebuilt[cut:]
    return paragraph_xml[: match.start()] + rebuilt + paragraph_xml[match.end():]


#: Author recorded on marker insertions when the source is under review.
#:
#: Deliberately *not* the document author. Three things depend on it being
#: distinguishable: Word's markup pane should show a machine conversion apart
#: from a person's own edits; the run-structure invariant projects the app's
#: revisions back out by author; and a reviewer validating that nothing changed
#: the author's content outside a revision must not have that check pass
#: trivially because the app signed the author's name to its own work.
MARKER_REVISION_AUTHOR = "Specification Formatter"

_TRACK_REVISIONS_RX = re.compile(
    r"<w:trackRevisions\b(?![^>]*\bw:val=\"(?:0|false|off)\")"
)


def source_tracks_revisions(settings_xml: str) -> bool:
    """Whether ``settings.xml`` has revision tracking switched on.

    ``<w:trackRevisions w:val="false"/>`` is the off state written explicitly,
    so a bare element test would read it backwards.
    """

    return bool(_TRACK_REVISIONS_RX.search(settings_xml or ""))


def _tracked_marker_run(marker: str, revision_id: int, date: str) -> str:
    return (
        f'<w:ins w:id="{revision_id}" w:author="{_escape_attribute(MARKER_REVISION_AUTHOR)}"'
        f' w:date="{date}"><w:r><w:t xml:space="preserve">{_escape(marker)}</w:t>'
        f"<w:tab/></w:r></w:ins>"
    )


def _insert_marker(
    paragraph_xml: str,
    marker: str,
    *,
    tracked: bool = False,
    revision_id: int = 0,
    revision_date: str = "",
) -> str:
    """Prepend *marker* and a tab at the start of the paragraph's text.

    Untracked, the marker goes *into* the paragraph's existing first run rather
    than into a run of its own, for two reasons. It inherits that run's
    character formatting, so a bold heading gets a bold number instead of a
    bare one in the document default; and the paragraph's run structure is
    unchanged, which is what the run-property invariant in ``phase2_invariants``
    checks. Adding runs would trip that check for a change that loses no
    formatting at all.

    Tracked, it cannot: a revision is a subtree, so the marker must be its own
    run inside ``w:ins``. That does shift every following run index, and the
    answer is still not to widen the invariant -- ``phase2_invariants`` runs the
    unchanged check against the document with this application's own revisions
    projected back out, where the run structure is identical again.

    A run may hold several ``w:t`` and ``w:tab`` children, so both results are
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
        if tracked:
            # The revision wraps its own run, so it is placed before the run
            # holding the first text rather than inside it.
            run_start = unprotected.rfind("<w:r", 0, match.start())
            if run_start < 0:
                raise EngineError(
                    _HIERARCHY,
                    f"A paragraph classified for CSI marker {marker!r} has no run "
                    "to place a tracked marker before.",
                )
            insertion = _tracked_marker_run(marker, revision_id, revision_date)
            return unprotected[:run_start] + insertion + unprotected[run_start:]
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


#: Numbering-level ``w:pPr`` children that carry the geometry a paragraph
#: loses when it stops being a list member, in schema order. ``w:ind`` is the
#: indentation itself; ``w:tabs`` is the level's num tab stop, which decides
#: where the text after the marker's tab actually lands.
_LEVEL_GEOMETRY_ELEMENTS = ("tabs", "ind")


def _serialize_w_element(element: ET.Element) -> str:
    """Serialize a WordprocessingML element as a prefixed, declaration-free fragment.

    ``ET.tostring`` would emit an ``xmlns`` declaration on the fragment root (or
    an ``ns0`` prefix), neither of which can be spliced into a paragraph that
    already declares ``w``. These elements are small and attribute-only apart
    from ``w:tabs``'s children, so the fragment is built directly.
    """

    local = element.tag.split("}", 1)[-1]
    attributes = "".join(
        f' w:{name.split("}", 1)[-1]}="{_escape_attribute(value)}"'
        for name, value in sorted(element.attrib.items())
    )
    children = "".join(_serialize_w_element(child) for child in element)
    if not children:
        return f"<w:{local}{attributes}/>"
    return f"<w:{local}{attributes}>{children}</w:{local}>"


def _escape_attribute(value: str) -> str:
    return (
        value.replace("&", "&amp;")
        .replace("<", "&lt;")
        .replace(">", "&gt;")
        .replace('"', "&quot;")
    )


def _restorable_level_geometry(
    geometry: Tuple[str, ...],
    paragraph_xml: str,
    styles_xml: str,
) -> Tuple[str, ...]:
    """Narrow *geometry* to the properties whose loss is actually provable.

    OOXML precedence between a paragraph style's own ``w:ind`` and the ``w:ind``
    of the numbering level it references is genuinely ambiguous -- the spec's
    style hierarchy puts paragraph styles after numbering, while Word's
    observed behaviour for a directly referenced list is the reverse. So this
    restores geometry only where nothing else could have supplied it: when the
    paragraph's effective style chain sets the property, that value was
    available before the edit and is still available after it, and guessing
    which of the two Word preferred would risk *changing* a rendering in order
    to protect it.

    That leaves the case this exists for -- a stylesheet whose list styles
    carry ``w:numPr`` and no ``w:ind`` at all, so the level was unambiguously
    the only source of indentation. Anything more ambitious is a guess, and
    the differential geometry invariant is what catches the residue.
    """

    if not geometry:
        return ()
    style_id = _paragraph_style_id(paragraph_xml)
    supplied: Set[str] = (
        _style_replacement_ppr_properties(styles_xml, style_id) if style_id else set()
    )
    return tuple(
        fragment
        for fragment in geometry
        if re.match(r"<w:(\w+)", fragment).group(1) not in supplied
    )


def _paragraph_style_id(paragraph_xml: str) -> Optional[str]:
    match = re.search(r'<w:pStyle\b[^>]*w:val="([^"]+)"', paragraph_xml)
    return match.group(1) if match else None


def _level_geometry(level: Optional[ET.Element]) -> Tuple[str, ...]:
    """Serialize the geometry a level's ``w:pPr`` contributes to its paragraphs.

    Returned in ``CT_PPr`` order so a caller can insert them in sequence.
    """

    if level is None:
        return ()
    level_ppr = level.find(_wq("pPr"))
    if level_ppr is None:
        return ()
    fragments: List[str] = []
    for local in _LEVEL_GEOMETRY_ELEMENTS:
        element = level_ppr.find(_wq(local))
        if element is not None:
            fragments.append(_serialize_w_element(element))
    return tuple(fragments)


def _validate_converted_list(
    numbering_xml: str,
    num_ids: set[str],
    ilvls: Dict[str, set[str]],
) -> Dict[Tuple[str, str], Tuple[str, ...]]:
    """Prove every converted list level has a countable, unoverridden counter.

    Also returns each level's geometry, keyed by ``(numId, ilvl)``. It is
    collected here rather than in a second pass because this is already the one
    place that resolves every converted level, and the suppression step needs
    exactly what this loop already holds.
    """

    if not num_ids:
        return {}
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
    geometry: Dict[Tuple[str, str], Tuple[str, ...]] = {}
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
            # ``level`` is already the effective one: _find_numbering_level
            # resolves an override's own ``w:lvl`` in preference to the
            # abstract level and returns that, so its geometry is the geometry
            # Word applies.
            geometry[(num_id, ilvl)] = _level_geometry(level)
    return geometry


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
    settings_xml: str = "",
    revision_date: str = "",
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
    level_geometry = _validate_converted_list(numbering_xml, num_ids, ilvls)

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

    # A paragraph whose *mark* is an unresolved tracked revision does not have
    # one position in the sequence, it has two: reject an inserted mark and the
    # paragraph disappears, accept a deleted one and it merges into the next.
    # Automatic numbering renumbers itself either way. A literal marker cannot,
    # so every marker after such a paragraph would silently become wrong the
    # moment somebody resolved the revision -- in a document a reader trusts,
    # long after this run is forgotten.
    for index in sorted(effective_role_by_index):
        if effective_role_by_index[index] not in _CONVERTIBLE_ROLES:
            continue
        revision = _paragraph_mark_revision(blocks[index][2])
        if revision is not None:
            raise EngineError(
                _TRACKED,
                f"Paragraph {index}{locate(index)} is classified as a numbered "
                f"role but its paragraph mark is a tracked {revision}; its CSI "
                "number would depend on whether that revision is accepted.",
            )

    _validate_source_sequence(evidence, describe=locate, error_code=_HIERARCHY)

    # --- Pass 2: walk the counters and build the edits ----------------------
    # A document with revision tracking on is in an active review cycle, and
    # writing 122 numbers into it as plain accepted text would put the
    # application's own work beyond the reach of the review everything else in
    # the document is subject to. Where the source says edits are tracked, the
    # markers are tracked too.
    tracked = source_tracks_revisions(settings_xml)
    date = revision_date or _utc_revision_date()
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
            source_numpr = automatic_indices[index]
            source_ilvl = str(source_numpr["ilvl"])
            stripped = _suppress_automatic_numbering(
                paragraph,
                source_ilvl,
                _restorable_level_geometry(
                    level_geometry.get((str(source_numpr["numId"]), source_ilvl), ()),
                    paragraph,
                    styles_xml,
                ),
            )
            body = paragraph_text_from_block(paragraph)
            source_kind = "automatic"
            automatic_converted += 1

        converted = _insert_marker(
            stripped,
            marker,
            tracked=tracked,
            revision_id=_MARKER_REVISION_ID_BASE + len(edits),
            revision_date=date,
        )
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
        source_tracks_revisions=tracked,
        markers_tracked=tracked and automatic_converted + literal_removed > 0,
        marker_author=MARKER_REVISION_AUTHOR if tracked else None,
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
    settings_path = Path(extract_dir) / "word" / "settings.xml"
    plan = plan_canadian_to_csi(
        read_xml_text(document_path),
        read_xml_text(styles_path),
        classifications,
        numbering_xml=(
            read_xml_text(numbering_path) if numbering_path.is_file() else ""
        ),
        settings_xml=(
            read_xml_text(settings_path) if settings_path.is_file() else ""
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
    "MARKER_REVISION_AUTHOR",
    "apply_canadian_to_csi",
    "plan_canadian_to_csi",
]
