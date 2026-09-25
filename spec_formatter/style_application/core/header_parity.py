"""The one document-global header switch: ``w:evenAndOddHeaders``.

A section's ``even`` header and footer render only while ``word/settings.xml``
carries ``<w:evenAndOddHeaders/>``; without it every page shows the
``default`` part and an ``even`` reference stays dormant. The switch is global,
so the header/footer importer cannot carry it with the parts it wires in. An
architect with distinct odd and even headers used to reach the output with its
default header on every page, and an architect without them, applied to a
target whose switch was on, left even pages with no header at all. Both were
silent.

This module is the one place the switch is read and written:

- :func:`even_and_odd_headers` reads it namespace-aware and fails closed when
  it cannot say which way Word renders: the element repeated, or a ``w:val``
  that is not one of the six ``ST_OnOff`` spellings.
- :func:`architect_even_and_odd_headers` derives the architect's setting from
  the registry's captured ``settings.settings_xml``, so neither the profile
  contract nor the engine fingerprint changes.
- :func:`set_even_and_odd_headers` sets or clears it in a settings part
  lexically, placing a new element by the complete ``CT_Settings`` sequence and
  proving afterwards that nothing else in the part moved.
"""

from __future__ import annotations

from typing import Any, List, Mapping, Optional, Tuple, Union

from .ooxml_namespaces import W_NS
from .untrusted_xml import parse_untrusted_xml
from .xml_helpers import (
    _root_open_tag,
    _tag_declarations,
    iter_direct_child_xml_blocks,
    root_namespace_declarations,
    root_opening_tag,
)

M_NS = "http://schemas.openxmlformats.org/officeDocument/2006/math"
SL_NS = "http://schemas.openxmlformats.org/schemaLibrary/2006/main"

_ORDER_TABLE_NAMESPACES = {"w": W_NS, "m": M_NS, "sl": SL_NS}

#: The complete child sequence of ``w:settings`` (``CT_Settings``) in the
#: transitional WordprocessingML schema. Checked element by element, on
#: 2026-09-25, against two copies of that schema text that agree with each
#: other: ISO/IEC 29500-4:2012 as printed in the standard (``wml.xsd``,
#: ``CT_Settings`` from schema line 2896, page 922), and the transitional
#: ``wml.xsd`` that python-docx ships in ``ref/xsd`` (commit ``e454546``,
#: SHA-256 ``753c5c80...38b310bb2c``). All 98 children are optional;
#: ``activeWritingStyle``, ``attachedSchema`` and ``smartTagType`` may repeat.
#: Do not abbreviate it for one element: a table shortened around the element
#: being inserted places it wrongly relative to whatever it left out.
CT_SETTINGS_CHILD_ORDER: Tuple[str, ...] = (
    "w:writeProtection",
    "w:view",
    "w:zoom",
    "w:removePersonalInformation",
    "w:removeDateAndTime",
    "w:doNotDisplayPageBoundaries",
    "w:displayBackgroundShape",
    "w:printPostScriptOverText",
    "w:printFractionalCharacterWidth",
    "w:printFormsData",
    "w:embedTrueTypeFonts",
    "w:embedSystemFonts",
    "w:saveSubsetFonts",
    "w:saveFormsData",
    "w:mirrorMargins",
    "w:alignBordersAndEdges",
    "w:bordersDoNotSurroundHeader",
    "w:bordersDoNotSurroundFooter",
    "w:gutterAtTop",
    "w:hideSpellingErrors",
    "w:hideGrammaticalErrors",
    "w:activeWritingStyle",
    "w:proofState",
    "w:formsDesign",
    "w:attachedTemplate",
    "w:linkStyles",
    "w:stylePaneFormatFilter",
    "w:stylePaneSortMethod",
    "w:documentType",
    "w:mailMerge",
    "w:revisionView",
    "w:trackRevisions",
    "w:doNotTrackMoves",
    "w:doNotTrackFormatting",
    "w:documentProtection",
    "w:autoFormatOverride",
    "w:styleLockTheme",
    "w:styleLockQFSet",
    "w:defaultTabStop",
    "w:autoHyphenation",
    "w:consecutiveHyphenLimit",
    "w:hyphenationZone",
    "w:doNotHyphenateCaps",
    "w:showEnvelope",
    "w:summaryLength",
    "w:clickAndTypeStyle",
    "w:defaultTableStyle",
    "w:evenAndOddHeaders",
    "w:bookFoldRevPrinting",
    "w:bookFoldPrinting",
    "w:bookFoldPrintingSheets",
    "w:drawingGridHorizontalSpacing",
    "w:drawingGridVerticalSpacing",
    "w:displayHorizontalDrawingGridEvery",
    "w:displayVerticalDrawingGridEvery",
    "w:doNotUseMarginsForDrawingGridOrigin",
    "w:drawingGridHorizontalOrigin",
    "w:drawingGridVerticalOrigin",
    "w:doNotShadeFormData",
    "w:noPunctuationKerning",
    "w:characterSpacingControl",
    "w:printTwoOnOne",
    "w:strictFirstAndLastChars",
    "w:noLineBreaksAfter",
    "w:noLineBreaksBefore",
    "w:savePreviewPicture",
    "w:doNotValidateAgainstSchema",
    "w:saveInvalidXml",
    "w:ignoreMixedContent",
    "w:alwaysShowPlaceholderText",
    "w:doNotDemarcateInvalidXml",
    "w:saveXmlDataOnly",
    "w:useXSLTWhenSaving",
    "w:saveThroughXslt",
    "w:showXMLTags",
    "w:alwaysMergeEmptyNamespace",
    "w:updateFields",
    "w:hdrShapeDefaults",
    "w:footnotePr",
    "w:endnotePr",
    "w:compat",
    "w:docVars",
    "w:rsids",
    "m:mathPr",
    "w:attachedSchema",
    "w:themeFontLang",
    "w:clrSchemeMapping",
    "w:doNotIncludeSubdocsInStats",
    "w:doNotAutoCompressPictures",
    "w:forceUpgrade",
    "w:captions",
    "w:readModeInkLockDown",
    "w:smartTagType",
    "sl:schemaLibrary",
    "w:shapeDefaults",
    "w:doNotEmbedSmartTags",
    "w:decimalSymbol",
    "w:listSeparator",
)


def _expanded_order_name(qualified_name: str) -> Tuple[str, str]:
    prefix, _, local = qualified_name.partition(":")
    return _ORDER_TABLE_NAMESPACES[prefix], local


#: Expanded name (namespace URI, local name) -> position in the sequence.
#: Keyed by namespace rather than by the prefix a part happens to use, so a
#: target that binds WordprocessingML to another prefix is ordered correctly.
_ORDER_INDEX = {
    _expanded_order_name(name): position
    for position, name in enumerate(CT_SETTINGS_CHILD_ORDER)
}

_SETTINGS = (W_NS, "settings")
_SWITCH = (W_NS, "evenAndOddHeaders")
_SWITCH_TAG = f"{{{W_NS}}}evenAndOddHeaders"
_VAL = f"{{{W_NS}}}val"

# ``ST_OnOff`` is the union of ``xsd:boolean`` and ``on``/``off``. Only the
# boolean branch collapses whitespace, but no spelling has inner space, so
# trimming both ends accepts exactly the valid values.
_ON_VALUES = frozenset({"1", "true", "on"})
_OFF_VALUES = frozenset({"0", "false", "off"})

XmlPayload = Union[str, bytes]


class HeaderParityError(ValueError):
    """The even/odd header switch cannot be read or written with certainty."""


def _switch_value(element: Any, part_name: str) -> bool:
    raw = element.get(_VAL)
    if raw is None:
        return True
    token = raw.strip()
    if token in _ON_VALUES:
        return True
    if token in _OFF_VALUES:
        return False
    # The value is not quoted back: it is document content, and a message
    # about it may travel further than the part it came from.
    raise HeaderParityError(
        f"{part_name}: w:evenAndOddHeaders has a w:val that is not an on/off "
        "value, so whether even-page headers render cannot be determined"
    )


def _settings_root(settings_xml: XmlPayload, part_name: str) -> Any:
    root = parse_untrusted_xml(settings_xml, part_name)
    if root.tag != f"{{{W_NS}}}settings":
        raise HeaderParityError(
            f"{part_name}: the root element is not a WordprocessingML w:settings"
        )
    return root


def even_and_odd_headers(settings_xml: Optional[XmlPayload], part_name: str) -> bool:
    """Whether a settings part switches even-page headers and footers on.

    ``None`` is a document without a settings part, which Word reads as off.
    A bare ``<w:evenAndOddHeaders/>`` is on; ``w:val`` may spell either state.
    The element written twice, or a value that is neither state, raises
    :class:`HeaderParityError` rather than guessing which reading Word takes.
    """

    if settings_xml is None:
        return False
    root = _settings_root(settings_xml, part_name)
    switches = [child for child in root if child.tag == _SWITCH_TAG]
    if not switches:
        return False
    if len(switches) > 1:
        raise HeaderParityError(
            f"{part_name}: w:evenAndOddHeaders is written {len(switches)} times, "
            "so whether even-page headers render cannot be determined"
        )
    return _switch_value(switches[0], part_name)


def even_and_odd_headers_as_written(
    settings_xml: Optional[XmlPayload], part_name: str
) -> Tuple[Optional[str], ...]:
    """The switch exactly as written, uninterpreted.

    One entry per ``w:evenAndOddHeaders`` child of the root, holding its raw
    ``w:val`` or ``None``. Comparing this before and after proves a part left
    its switch alone even where :func:`even_and_odd_headers` would refuse to
    read it.
    """

    if settings_xml is None:
        return ()
    root = _settings_root(settings_xml, part_name)
    return tuple(child.get(_VAL) for child in root if child.tag == _SWITCH_TAG)


def architect_even_and_odd_headers(template_registry: Mapping[str, Any]) -> bool:
    """The architect's even/odd header setting, from the template registry.

    Phase 1 captures the architect's whole settings part, canonicalized, as
    ``settings.settings_xml`` (``None`` when the template has none). A registry
    that does not record the field at all says nothing about the switch, and
    that is not the same as saying it is off.
    """

    settings = (
        template_registry.get("settings")
        if isinstance(template_registry, Mapping)
        else None
    )
    if not isinstance(settings, Mapping) or "settings_xml" not in settings:
        raise HeaderParityError(
            "The template registry does not record the architect's settings "
            "part (settings.settings_xml), so its even/odd header setting is "
            "unknown"
        )
    settings_xml = settings["settings_xml"]
    if settings_xml is not None and not isinstance(settings_xml, str):
        raise HeaderParityError(
            "settings.settings_xml in the template registry must be a string or null"
        )
    return even_and_odd_headers(settings_xml, "architect settings (settings.settings_xml)")


def _expanded_child_name(
    qualified_name: str,
    block: str,
    root_declarations: Mapping[str, str],
) -> Tuple[Optional[str], str]:
    prefix, _, local = qualified_name.rpartition(":")
    scope = dict(root_declarations)
    scope.update(_tag_declarations(root_opening_tag(block)))
    return scope.get(prefix), local


def _children(settings_xml: str) -> List[Tuple[int, int, Tuple[Optional[str], str], str]]:
    declarations = root_namespace_declarations(settings_xml)
    return [
        (start, end, _expanded_child_name(name, block, declarations), block)
        for start, end, name, block in iter_direct_child_xml_blocks(settings_xml)
    ]


def _main_namespace_prefix(settings_xml: str, part_name: str) -> str:
    declarations = root_namespace_declarations(settings_xml)
    if declarations.get("w") == W_NS:
        return "w"
    for prefix, uri in sorted(declarations.items()):
        if uri == W_NS:
            return prefix
    raise HeaderParityError(
        f"{part_name}: the root does not declare the WordprocessingML namespace, "
        "so w:evenAndOddHeaders cannot be written into it"
    )


def _insertion_point(
    settings_xml: str,
    children: List[Tuple[int, int, Tuple[Optional[str], str], str]],
    part_name: str,
) -> int:
    """Where a new ``w:evenAndOddHeaders`` goes: after everything that precedes it.

    The position is the end of the last child the sequence places before the
    switch, or just inside the root when there is none. An element the table
    does not know (an extension such as ``w14:docId``, which Word writes last)
    neither pulls the switch forward nor pushes it back. A known child that
    must *follow* the switch but sits before that point means the part is out
    of schema order already, and no position would be valid; that fails
    rather than adding a second ordering error to it.
    """

    rank = _ORDER_INDEX[_SWITCH]
    position: Optional[int] = None
    for _start, end, name, _block in children:
        child_rank = _ORDER_INDEX.get(name)
        if child_rank is not None and child_rank < rank:
            position = end
    if position is None:
        _open_start, position = _root_open_tag(settings_xml)
    for start, _end, name, _block in children:
        child_rank = _ORDER_INDEX.get(name)
        if child_rank is not None and child_rank > rank and start < position:
            raise HeaderParityError(
                f"{part_name}: the settings children are not in schema order, "
                "so w:evenAndOddHeaders has no valid position"
            )
    return position


def _expand_self_closing_root(settings_xml: str) -> str:
    """``<w:settings .../>`` as an open and a close tag, so a child fits inside."""

    start, end = _root_open_tag(settings_xml)
    tag = settings_xml[start:end]
    if not tag[:-1].rstrip().endswith("/"):
        return settings_xml
    name = tag[1:].split(None, 1)[0].rstrip("/>")
    opened = tag[: tag.rindex("/")].rstrip() + ">"
    return f"{settings_xml[:start]}{opened}</{name}>{settings_xml[end:]}"


def _without_switches(settings_xml: str) -> str:
    for start, end, name, _block in reversed(_children(settings_xml)):
        if name == _SWITCH:
            settings_xml = settings_xml[:start] + settings_xml[end:]
    return settings_xml


def set_even_and_odd_headers(
    settings_xml: str,
    enabled: bool,
    *,
    part_name: str = "word/settings.xml",
) -> str:
    """Return ``settings_xml`` with the switch set (``enabled``) or cleared.

    Off is written the way Word writes it, by leaving the element out. On is a
    single bare ``<w:evenAndOddHeaders/>`` at its schema position; one already
    reading on is left exactly as written. Every other byte of the part is
    kept, and the result is proven afterwards: it parses, it reads as
    ``enabled``, and every other child of the root is unchanged and in the
    same order.
    """

    _settings_root(settings_xml, part_name)
    children = _children(settings_xml)
    switches = [child for child in children if child[2] == _SWITCH]
    if not switches and not enabled:
        return settings_xml
    if enabled and len(switches) == 1:
        try:
            if even_and_odd_headers(settings_xml, part_name):
                return settings_xml
        except HeaderParityError:
            pass  # An unreadable value is replaced below.

    updated = settings_xml
    for start, end, _name, _block in reversed(switches):
        updated = updated[:start] + updated[end:]
    if enabled:
        prefix = _main_namespace_prefix(updated, part_name)
        element = f"<{prefix}:evenAndOddHeaders/>" if prefix else "<evenAndOddHeaders/>"
        updated = _expand_self_closing_root(updated)
        position = _insertion_point(updated, _children(updated), part_name)
        updated = updated[:position] + element + updated[position:]

    _verify_switch_edit(settings_xml, updated, enabled, part_name)
    return updated


def _verify_switch_edit(before: str, after: str, enabled: bool, part_name: str) -> None:
    """Prove the edit changed the switch and nothing else.

    Read back independently of how it was written: the result must read as
    ``enabled`` through the namespace-aware reader, must be byte-for-byte the
    original once every switch is taken out of both (the only other edit
    allowed is opening a self-closing root), and a switch that was written must
    sit after every child the sequence puts before it and before every child it
    puts after.
    """

    if even_and_odd_headers(after, part_name) is not enabled:
        raise HeaderParityError(
            f"{part_name}: w:evenAndOddHeaders does not read as intended after the edit"
        )
    original = _expand_self_closing_root(before) if enabled else before
    if _without_switches(after) != _without_switches(original):
        raise HeaderParityError(
            f"{part_name}: setting w:evenAndOddHeaders changed more of the part than the switch"
        )
    if not enabled:
        return
    rank = _ORDER_INDEX[_SWITCH]
    ranks = [_ORDER_INDEX.get(name) for _s, _e, name, _block in _children(after)]
    switch_at = next(
        index
        for index, (_s, _e, name, _block) in enumerate(_children(after))
        if name == _SWITCH
    )
    if any(
        child_rank is not None and child_rank > rank for child_rank in ranks[:switch_at]
    ) or any(
        child_rank is not None and child_rank < rank for child_rank in ranks[switch_at + 1:]
    ):
        raise HeaderParityError(
            f"{part_name}: w:evenAndOddHeaders is not at its schema position"
        )


__all__ = [
    "CT_SETTINGS_CHILD_ORDER",
    "HeaderParityError",
    "architect_even_and_odd_headers",
    "even_and_odd_headers",
    "even_and_odd_headers_as_written",
    "set_even_and_odd_headers",
]
