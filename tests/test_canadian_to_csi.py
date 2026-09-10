"""Unit coverage for the Canadian -> CSI converter.

Writing a number is a stronger claim than removing one: a wrong marker becomes
literal text in a document a reader will trust. Most of what follows therefore
tests the refusals rather than the conversions.
"""

from __future__ import annotations

import re

import pytest

from spec_formatter import builtin_scheme
from spec_formatter.role_contract import ROLE_TO_ARCH_STYLE
from spec_formatter.style_application.core.canadian_to_csi import (
    _csi_marker,
    _verify_prediction,
    _int_to_alpha,
    plan_canadian_to_csi,
)
from spec_formatter.style_application.core.errors import EngineError
from spec_formatter.style_application.core.xml_helpers import (
    iter_paragraph_xml_blocks,
    paragraph_text_from_block,
)

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"


def _styled(role: str, text: str) -> str:
    return (
        f'<w:p><w:pPr><w:pStyle w:val="{ROLE_TO_ARCH_STYLE[role]}"/></w:pPr>'
        f'<w:r><w:t xml:space="preserve">{text}</w:t></w:r></w:p>'
    )


def _document(paragraphs: list[str]) -> str:
    return f'<w:document xmlns:w="{W_NS}"><w:body>{"".join(paragraphs)}</w:body></w:document>'


def _classifications(roles: list[str]) -> dict:
    return {
        "classifications": [
            {"paragraph_index": index, "csi_role": role}
            for index, role in enumerate(roles)
        ]
    }


def _convert(rows: list[tuple[str, str]], **kwargs):
    roles = [role for role, _text in rows]
    document = _document([_styled(role, text) for role, text in rows])
    return plan_canadian_to_csi(
        document,
        builtin_scheme.build_styles_xml(),
        _classifications(roles),
        numbering_xml=kwargs.pop("numbering_xml", builtin_scheme.build_numbering_xml()),
        **kwargs,
    )


def _lines(document_xml: str) -> list[str]:
    return [
        paragraph_text_from_block(block)
        for _start, _end, block in iter_paragraph_xml_blocks(document_xml)
    ]


def test_counters_nest_and_reset_like_word() -> None:
    plan = _convert(
        [
            ("PART", "GENERAL"),
            ("ARTICLE", "SUMMARY"),
            ("PARAGRAPH", "First requirement."),
            ("PARAGRAPH", "Second requirement."),
            ("SUBPARAGRAPH", "A nested item."),
            ("ARTICLE", "REFERENCES"),
            ("PARAGRAPH", "Back to A under a new article."),
            ("PART", "PRODUCTS"),
            ("ARTICLE", "SPRINKLERS"),
            ("PARAGRAPH", "First again under part two."),
        ]
    )
    assert _lines(plan.document_xml) == [
        "PART 1 GENERAL",
        "1.1 SUMMARY",
        "A. First requirement.",
        "B. Second requirement.",
        "1. A nested item.",
        "1.2 REFERENCES",
        "A. Back to A under a new article.",
        "PART 2 PRODUCTS",
        "2.1 SPRINKLERS",
        "A. First again under part two.",
    ]
    assert plan.report.paragraphs_converted == 10
    assert plan.report.automatic_numbering_retargeted == 10


def test_deep_levels_use_their_csi_marker_shapes() -> None:
    plan = _convert(
        [
            ("PART", "GENERAL"),
            ("ARTICLE", "SUMMARY"),
            ("PARAGRAPH", "Alpha."),
            ("SUBPARAGRAPH", "Numeric."),
            ("SUBSUBPARAGRAPH", "Lower alpha."),
            ("SUBPARAGRAPH_LEVEL_5", "Numeric paren."),
            ("SUBPARAGRAPH_LEVEL_6", "Alpha paren."),
            ("SUBPARAGRAPH_LEVEL_7", "Bracketed numeric."),
            ("SUBPARAGRAPH_LEVEL_8", "Bracketed alpha."),
        ]
    )
    assert _lines(plan.document_xml)[2:] == [
        "A. Alpha.",
        "1. Numeric.",
        "a. Lower alpha.",
        "1) Numeric paren.",
        "a) Alpha paren.",
        "(1) Bracketed numeric.",
        "(a) Bracketed alpha.",
    ]


def test_typed_canadian_markers_are_rewritten_as_csi_markers() -> None:
    """A Canadian document may carry typed markers rather than a Word list."""

    rows = [
        ("PART", "PART 1", "GENERAL"),
        ("ARTICLE", "1.1", "SUMMARY"),
        ("PARAGRAPH", ".1", "First Canadian requirement."),
        ("PARAGRAPH", ".2", "Second Canadian requirement."),
        ("SUBPARAGRAPH", ".1", "Nested Canadian item."),
    ]
    document = _document(
        [
            "<w:p><w:pPr/>"
            f'<w:r><w:t xml:space="preserve">{marker}</w:t></w:r><w:r><w:tab/></w:r>'
            f'<w:r><w:t xml:space="preserve">{text}</w:t></w:r></w:p>'
            for _role, marker, text in rows
        ]
    )
    plan = plan_canadian_to_csi(
        document,
        builtin_scheme.build_styles_xml(),
        _classifications([role for role, _m, _t in rows]),
        numbering_xml="",
    )
    assert plan.report.literal_markers_removed == 5
    assert plan.report.automatic_numbering_retargeted == 0
    assert _lines(plan.document_xml) == [
        "PART 1 GENERAL",
        "1.1 SUMMARY",
        "A. First Canadian requirement.",
        "B. Second Canadian requirement.",
        "1. Nested Canadian item.",
    ]


def test_non_contiguous_typed_markers_fail_closed() -> None:
    rows = [
        ("PART", "PART 1", "GENERAL"),
        ("ARTICLE", "1.1", "SUMMARY"),
        ("PARAGRAPH", ".1", "First."),
        # .3 skips .2, so the source sequence cannot be trusted.
        ("PARAGRAPH", ".3", "Third."),
    ]
    document = _document(
        [
            "<w:p><w:pPr/>"
            f'<w:r><w:t xml:space="preserve">{marker}</w:t></w:r><w:r><w:tab/></w:r>'
            f'<w:r><w:t xml:space="preserve">{text}</w:t></w:r></w:p>'
            for _role, marker, text in rows
        ]
    )
    with pytest.raises(EngineError) as raised:
        plan_canadian_to_csi(
            document,
            builtin_scheme.build_styles_xml(),
            _classifications([role for role, _m, _t in rows]),
            numbering_xml="",
        )
    assert raised.value.code == "canadian_to_csi_hierarchy"
    assert "non-contiguous" in str(raised.value)


def test_paragraph_without_a_number_is_preserved_not_invented() -> None:
    document = _document(
        [
            _styled("PART", "GENERAL"),
            _styled("ARTICLE", "SUMMARY"),
            # No style, so no automatic numbering and no typed marker.
            '<w:p><w:pPr/><w:r><w:t>A stray prose line.</w:t></w:r></w:p>',
        ]
    )
    plan = plan_canadian_to_csi(
        document,
        builtin_scheme.build_styles_xml(),
        _classifications(["PART", "ARTICLE", "PARAGRAPH"]),
        numbering_xml=builtin_scheme.build_numbering_xml(),
    )
    assert _lines(plan.document_xml)[2] == "A stray prose line."
    assert [issue.code for issue in plan.report.warnings] == [
        "unproven_numbered_role_preserved"
    ]


def test_article_without_a_part_fails_closed() -> None:
    with pytest.raises(EngineError) as raised:
        _convert([("ARTICLE", "SUMMARY")])
    assert raised.value.code == "canadian_to_csi_hierarchy"


def test_doubled_numbering_fails_closed() -> None:
    document = _document(
        [
            _styled("PART", "GENERAL"),
            _styled("ARTICLE", "SUMMARY"),
            # Automatic numbering from the style *and* a typed marker.
            f'<w:p><w:pPr><w:pStyle w:val="{ROLE_TO_ARCH_STYLE["PARAGRAPH"]}"/></w:pPr>'
            '<w:r><w:t xml:space="preserve">A.</w:t></w:r><w:r><w:tab/></w:r>'
            '<w:r><w:t xml:space="preserve">Doubled.</w:t></w:r></w:p>',
        ]
    )
    with pytest.raises(EngineError) as raised:
        plan_canadian_to_csi(
            document,
            builtin_scheme.build_styles_xml(),
            _classifications(["PART", "ARTICLE", "PARAGRAPH"]),
            numbering_xml=builtin_scheme.build_numbering_xml(),
        )
    assert raised.value.code == "canadian_to_csi_hierarchy"
    assert "doubled numbering" in str(raised.value)


def test_a_paragraph_left_on_the_converted_list_fails_closed() -> None:
    """One unconverted list member would desynchronise every later counter."""

    document = _document(
        [
            _styled("PART", "GENERAL"),
            _styled("ARTICLE", "SUMMARY"),
            _styled("PARAGRAPH", "Classified."),
            # Same list, but never classified, so never converted.
            _styled("PARAGRAPH", "Unclassified but still numbered."),
        ]
    )
    with pytest.raises(EngineError) as raised:
        plan_canadian_to_csi(
            document,
            builtin_scheme.build_styles_xml(),
            _classifications(["PART", "ARTICLE", "PARAGRAPH"]),
            numbering_xml=builtin_scheme.build_numbering_xml(),
        )
    assert raised.value.code == "canadian_to_csi_hierarchy"
    assert "desynchronise" in str(raised.value)


def test_numbering_that_does_not_start_at_one_fails_closed() -> None:
    numbering = builtin_scheme.build_numbering_xml().replace(
        '<w:start w:val="1"/>', '<w:start w:val="4"/>', 1
    )
    with pytest.raises(EngineError) as raised:
        _convert(
            [("PART", "GENERAL"), ("ARTICLE", "SUMMARY")],
            numbering_xml=numbering,
        )
    assert raised.value.code == "canadian_numbering_unprovable"


def test_missing_numbering_part_fails_closed() -> None:
    with pytest.raises(EngineError) as raised:
        _convert([("PART", "GENERAL"), ("ARTICLE", "SUMMARY")], numbering_xml="")
    assert raised.value.code == "canadian_to_csi_numbering_unprovable"


def test_marker_inherits_the_run_it_joins() -> None:
    """The marker goes into the existing run, so it keeps its formatting."""

    document = _document(
        [
            _styled("PART", "GENERAL"),
            f'<w:p><w:pPr><w:pStyle w:val="{ROLE_TO_ARCH_STYLE["ARTICLE"]}"/></w:pPr>'
            "<w:r><w:rPr><w:b/></w:rPr>"
            '<w:t xml:space="preserve">SUMMARY</w:t></w:r></w:p>',
        ]
    )
    plan = plan_canadian_to_csi(
        document,
        builtin_scheme.build_styles_xml(),
        _classifications(["PART", "ARTICLE"]),
        numbering_xml=builtin_scheme.build_numbering_xml(),
    )
    article = list(iter_paragraph_xml_blocks(plan.document_xml))[1][2]
    assert article.count("<w:r>") == 1
    assert "<w:b/>" in article
    assert article.index("<w:b/>") < article.index("1.1")


def test_section_properties_survive_the_conversion() -> None:
    sect = '<w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>'
    document = (
        f'<w:document xmlns:w="{W_NS}"><w:body>'
        + _styled("PART", "GENERAL")
        + _styled("ARTICLE", "SUMMARY")
        + f"{sect}</w:body></w:document>"
    )
    plan = plan_canadian_to_csi(
        document,
        builtin_scheme.build_styles_xml(),
        _classifications(["PART", "ARTICLE"]),
        numbering_xml=builtin_scheme.build_numbering_xml(),
    )
    assert sect in plan.document_xml


def test_alphabetic_levels_refuse_to_run_past_z() -> None:
    """Only a single letter is a marker this application can read back."""

    assert _int_to_alpha(1) == "a"
    assert _int_to_alpha(26) == "z"
    with pytest.raises(EngineError) as raised:
        _int_to_alpha(27)
    assert raised.value.code == "canadian_to_csi_numbering_unprovable"


def test_a_role_that_disagrees_with_its_list_level_fails_closed() -> None:
    """Regression: the marker must match the number the document shows.

    Two paragraphs sit at ``ilvl`` 0, so Word renders "PART 1" and "PART 2".
    Classifying the second as ARTICLE would previously write "1.1" over it --
    a number the document never displayed, written as permanent text.
    """

    document = _document([_styled("PART", "GENERAL"), _styled("PART", "PRODUCTS")])
    with pytest.raises(EngineError) as raised:
        plan_canadian_to_csi(
            document,
            builtin_scheme.build_styles_xml(),
            _classifications(["PART", "ARTICLE"]),
            numbering_xml=builtin_scheme.build_numbering_xml(),
        )
    assert raised.value.code == "canadian_to_csi_numbering_unprovable"


def test_marker_is_refused_inside_a_tracked_insertion() -> None:
    """Regression: a marker inside a revision dies when the revision is rejected.

    The numbering suppression lives on ``w:pPr``, outside the revision, so the
    paragraph would be left with no number at all.
    """

    document = _document(
        [
            _styled("PART", "GENERAL"),
            _styled("ARTICLE", "SUMMARY"),
            f'<w:p><w:pPr><w:pStyle w:val="{ROLE_TO_ARCH_STYLE["PARAGRAPH"]}"/></w:pPr>'
            '<w:ins w:id="7" w:author="a">'
            '<w:r><w:t xml:space="preserve">Inserted requirement.</w:t></w:r>'
            "</w:ins></w:p>",
        ]
    )
    with pytest.raises(EngineError) as raised:
        plan_canadian_to_csi(
            document,
            builtin_scheme.build_styles_xml(),
            _classifications(["PART", "ARTICLE", "PARAGRAPH"]),
            numbering_xml=builtin_scheme.build_numbering_xml(),
        )
    assert raised.value.code == "canadian_to_csi_hierarchy"
    assert "tracked change or field" in str(raised.value)


def test_suppressed_numbering_keeps_ooxml_ppr_child_order() -> None:
    """Regression: ``w:numPr`` must not be written before ``w:pStyle``.

    ``CT_PPr`` is a sequence, so the reverse order is invalid OOXML even where
    Word tolerates it, and a stricter consumer may reject the file.
    """

    plan = _convert([("PART", "GENERAL"), ("ARTICLE", "SUMMARY")])
    blocks = re.findall(r"<w:pPr>[\s\S]*?</w:pPr>", plan.document_xml)
    assert blocks
    for block in blocks:
        if "<w:numPr" in block and "<w:pStyle" in block:
            assert block.index("<w:pStyle") < block.index("<w:numPr"), block


def test_article_number_follows_the_active_part() -> None:
    assert _csi_marker("ARTICLE", {0: 3, 1: 7}) == "3.7"
    assert _csi_marker("PART", {0: 4}) == "PART 4"


# --- Numbering-derived geometry ------------------------------------------
#
# The stylesheet below is the shape that made this class of defect invisible
# for so long, and it is the ordinary shape rather than an exotic one: the
# list styles carry ``w:numPr`` and no ``w:ind`` at all, so *every* indent in
# the document is supplied by the numbering level. Cancelling the list without
# putting that geometry back flattens the whole outline into one column while
# leaving text, numbers and run structure provably intact -- which is why the
# fixtures that assert on marker text could never see it.

#: ``(left, hanging, numFmt, lvlText)`` per level, copied from the real
#: Canadian PageFormat template this defect was found in, so the fixture
#: exercises the same numbering signatures the role validator checks.
_GEOMETRY_LEVELS = {
    0: (720, 720, "decimal", "%1"),
    1: (720, 720, "decimalZero", "%1.%2"),
    2: (1259, 539, "decimal", ".%3"),
    3: (1797, 538, "decimal", ".%4"),
}


def _geometry_numbering_xml() -> str:
    levels = "".join(
        f'<w:lvl w:ilvl="{ilvl}"><w:start w:val="1"/>'
        f'<w:numFmt w:val="{fmt}"/><w:lvlText w:val="{text}"/>'
        f"<w:pPr><w:tabs><w:tab w:val=\"num\" w:pos=\"{left}\"/></w:tabs>"
        f'<w:ind w:left="{left}" w:hanging="{hanging}"/></w:pPr></w:lvl>'
        for ilvl, (left, hanging, fmt, text) in sorted(_GEOMETRY_LEVELS.items())
    )
    return (
        f'<w:numbering xmlns:w="{W_NS}">'
        f'<w:abstractNum w:abstractNumId="7">{levels}</w:abstractNum>'
        f'<w:num w:numId="4"><w:abstractNumId w:val="7"/></w:num>'
        f"</w:numbering>"
    )


def _geometry_styles_xml(*, list1_indent: str = "") -> str:
    """Styles that source their indentation only from the numbering level."""

    styles = "".join(
        f'<w:style w:type="paragraph" w:styleId="Geo{ilvl}">'
        f'<w:name w:val="Geo {ilvl}"/><w:basedOn w:val="Normal"/>'
        f'<w:pPr><w:numPr><w:ilvl w:val="{ilvl}"/><w:numId w:val="4"/></w:numPr>'
        f"{list1_indent if ilvl == 2 else ''}</w:pPr></w:style>"
        for ilvl in sorted(_GEOMETRY_LEVELS)
    )
    return (
        f'<w:styles xmlns:w="{W_NS}">'
        f'<w:style w:type="paragraph" w:styleId="Normal"><w:name w:val="Normal"/></w:style>'
        f"{styles}</w:styles>"
    )


_GEOMETRY_ROLE_STYLE = {
    "PART": "Geo0",
    "ARTICLE": "Geo1",
    "PARAGRAPH": "Geo2",
    "SUBPARAGRAPH": "Geo3",
}


def _convert_geometry(rows, *, styles_xml=None, extra_ppr=""):
    roles = [role for role, _text in rows]
    paragraphs = [
        f'<w:p><w:pPr><w:pStyle w:val="{_GEOMETRY_ROLE_STYLE[role]}"/>{extra_ppr}</w:pPr>'
        f'<w:r><w:t xml:space="preserve">{text}</w:t></w:r></w:p>'
        for role, text in rows
    ]
    return plan_canadian_to_csi(
        _document(paragraphs),
        styles_xml if styles_xml is not None else _geometry_styles_xml(),
        _classifications(roles),
        numbering_xml=_geometry_numbering_xml(),
    )


def _ppr_of(document_xml: str, index: int) -> str:
    block = list(iter_paragraph_xml_blocks(document_xml))[index][2]
    match = re.search(r"<w:pPr>[\s\S]*?</w:pPr>", block)
    return match.group(0) if match else ""


def test_cancelled_numbering_restores_the_level_indent() -> None:
    """The indent the level supplied survives the loss of the level."""

    plan = _convert_geometry(
        [
            ("PART", "GENERAL"),
            ("ARTICLE", "SUMMARY"),
            ("PARAGRAPH", "Section includes."),
            ("SUBPARAGRAPH", "A related requirement."),
        ]
    )
    for index, ilvl in enumerate((0, 1, 2, 3)):
        left, hanging, _fmt, _text = _GEOMETRY_LEVELS[ilvl]
        ppr = _ppr_of(plan.document_xml, index)
        assert f'w:left="{left}"' in ppr, ppr
        assert f'w:hanging="{hanging}"' in ppr, ppr
        # The level's num tab stop decides where the text after the marker's
        # tab lands; without it the tab falls through to w:defaultTabStop.
        assert f'w:pos="{left}"' in ppr, ppr


def test_cancelled_numbering_records_the_real_level() -> None:
    """``w:ilvl`` states the level the paragraph actually sat at.

    With ``numId=0`` the value does not render, but a paragraph claiming level
    0 while sitting at level 3 describes itself falsely to every later reader,
    this application's own converters included.
    """

    plan = _convert_geometry(
        [("PART", "GENERAL"), ("ARTICLE", "SUMMARY"), ("PARAGRAPH", "Body.")]
    )
    levels = re.findall(r'<w:numPr><w:ilvl w:val="(\d+)"/><w:numId w:val="0"/>', plan.document_xml)
    assert levels == ["0", "1", "2"]


def test_restored_geometry_keeps_ooxml_ppr_child_order() -> None:
    """``w:tabs`` (#11) and ``w:ind`` (#23) land at their schema positions."""

    plan = _convert_geometry([("PARAGRAPH", "Body.")])
    ppr = _ppr_of(plan.document_xml, 0)
    order = [
        ppr.index(tag)
        for tag in ("<w:pStyle", "<w:numPr", "<w:tabs", "<w:ind")
        if tag in ppr
    ]
    assert order == sorted(order), ppr


def test_style_supplied_indent_is_left_alone() -> None:
    """A style that sets ``w:ind`` keeps supplying it; nothing is injected.

    Precedence between a style's ``w:ind`` and its numbering level's is
    genuinely ambiguous, so the converter restores geometry only where nothing
    else could have supplied it rather than guessing which Word preferred.
    """

    styles = _geometry_styles_xml(list1_indent='<w:ind w:left="999" w:hanging="111"/>')
    plan = _convert_geometry([("PARAGRAPH", "Body.")], styles_xml=styles)
    ppr = _ppr_of(plan.document_xml, 0)
    assert "<w:ind" not in ppr, ppr
    # The level's tab stop is still restored: the style supplies no w:tabs.
    assert "<w:tabs>" in ppr, ppr


def test_direct_paragraph_indent_outranks_the_level() -> None:
    """A direct ``w:ind`` is the author's own choice and is never overwritten."""

    plan = _convert_geometry(
        [("PARAGRAPH", "Body.")],
        extra_ppr='<w:ind w:left="2500"/>',
    )
    ppr = _ppr_of(plan.document_xml, 0)
    assert 'w:left="2500"' in ppr, ppr
    assert 'w:left="1259"' not in ppr, ppr


# --- Tracked revisions on the paragraph mark ------------------------------


def test_tracked_inserted_paragraph_mark_is_refused() -> None:
    """A heading that is itself an unresolved insertion has no single number.

    Reject the insertion and the paragraph disappears; every literal marker
    after it is then silently wrong, in a document a reader trusts, long after
    this run is forgotten. Automatic numbering renumbers itself either way,
    which is exactly the safety net a literal marker gives up.
    """

    rows = [("PART", "GENERAL"), ("ARTICLE", "SUMMARY"), ("ARTICLE", "REFERENCES")]
    paragraphs = [
        f'<w:p><w:pPr><w:pStyle w:val="{_GEOMETRY_ROLE_STYLE[role]}"/>'
        + (
            '<w:rPr><w:ins w:id="9" w:author="A" w:date="2026-01-01T00:00:00Z"/></w:rPr>'
            if index == 1
            else ""
        )
        + f'</w:pPr><w:r><w:t xml:space="preserve">{text}</w:t></w:r></w:p>'
        for index, (role, text) in enumerate(rows)
    ]
    with pytest.raises(EngineError) as raised:
        plan_canadian_to_csi(
            _document(paragraphs),
            _geometry_styles_xml(),
            _classifications([role for role, _t in rows]),
            numbering_xml=_geometry_numbering_xml(),
        )
    assert raised.value.code == "canadian_to_csi_tracked_hierarchy"
    assert "tracked insertion" in str(raised.value)


def test_tracked_deleted_paragraph_mark_is_refused() -> None:
    """Accepting a deleted mark merges the paragraph away -- same hazard."""

    rows = [("PART", "GENERAL"), ("ARTICLE", "SUMMARY")]
    paragraphs = [
        f'<w:p><w:pPr><w:pStyle w:val="{_GEOMETRY_ROLE_STYLE[role]}"/>'
        + (
            '<w:rPr><w:del w:id="9" w:author="A" w:date="2026-01-01T00:00:00Z"/></w:rPr>'
            if index == 1
            else ""
        )
        + f'</w:pPr><w:r><w:t xml:space="preserve">{text}</w:t></w:r></w:p>'
        for index, (role, text) in enumerate(rows)
    ]
    with pytest.raises(EngineError) as raised:
        plan_canadian_to_csi(
            _document(paragraphs),
            _geometry_styles_xml(),
            _classifications([role for role, _t in rows]),
            numbering_xml=_geometry_numbering_xml(),
        )
    assert raised.value.code == "canadian_to_csi_tracked_hierarchy"
    assert "tracked deletion" in str(raised.value)


def test_inline_tracked_insertion_still_converts() -> None:
    """Tracked *text* inside a paragraph is not a tracked paragraph mark.

    This is the ordinary case in a spec under review -- a sentence added to an
    existing requirement -- and it changes no paragraph's position in the
    sequence, so the mark guard is bounded to ``w:pPr`` precisely to let it
    through instead of failing a whole document closed.
    """

    paragraphs = [
        '<w:p><w:pPr><w:pStyle w:val="Geo0"/></w:pPr>'
        '<w:r><w:t xml:space="preserve">GENERAL</w:t></w:r></w:p>',
        '<w:p><w:pPr><w:pStyle w:val="Geo1"/></w:pPr>'
        '<w:r><w:t xml:space="preserve">SUMMARY</w:t></w:r>'
        '<w:ins w:id="9" w:author="A" w:date="2026-01-01T00:00:00Z">'
        '<w:r><w:t xml:space="preserve"> AND SCOPE</w:t></w:r></w:ins></w:p>',
    ]
    plan = plan_canadian_to_csi(
        _document(paragraphs),
        _geometry_styles_xml(),
        _classifications(["PART", "ARTICLE"]),
        numbering_xml=_geometry_numbering_xml(),
    )
    assert plan.report.paragraphs_converted == 2
    assert "1.1" in plan.document_xml


# --- Tracked markers ------------------------------------------------------


_TRACKING_ON = f'<w:settings xmlns:w="{W_NS}"><w:trackRevisions/></w:settings>'
_TRACKING_OFF = f'<w:settings xmlns:w="{W_NS}"><w:trackRevisions w:val="false"/></w:settings>'


def _convert_tracked(rows, settings_xml):
    roles = [role for role, _text in rows]
    paragraphs = [
        f'<w:p><w:pPr><w:pStyle w:val="{_GEOMETRY_ROLE_STYLE[role]}"/></w:pPr>'
        f'<w:r><w:t xml:space="preserve">{text}</w:t></w:r></w:p>'
        for role, text in rows
    ]
    return plan_canadian_to_csi(
        _document(paragraphs),
        _geometry_styles_xml(),
        _classifications(roles),
        numbering_xml=_geometry_numbering_xml(),
        settings_xml=settings_xml,
        revision_date="2026-01-01T00:00:00Z",
    )


def test_markers_are_tracked_when_the_source_tracks_revisions() -> None:
    """A document under review gets the application's work reviewed too.

    Writing numbers into a spec with tracking on as plain accepted text puts
    them beyond the reach of the review every other change in the file is
    subject to -- and a validator asking "was anything changed outside a
    revision" would name all of them.
    """

    plan = _convert_tracked([("PART", "GENERAL"), ("ARTICLE", "SUMMARY")], _TRACKING_ON)
    assert plan.report.source_tracks_revisions is True
    assert plan.report.markers_tracked is True
    assert plan.report.marker_author == "Specification Formatter"
    insertions = re.findall(
        r'<w:ins\b[^>]*w:author="Specification Formatter"[^>]*>', plan.document_xml
    )
    assert len(insertions) == 2
    # Each paragraph carries two revisions, not one: the marker insertion and
    # the property change that took its automatic numbering away. Tracking
    # only the first would let a rejection strip the marker and leave the
    # suppression, giving the paragraph no number at all.
    changes = re.findall(
        r'<w:pPrChange\b[^>]*w:author="Specification Formatter"[^>]*>', plan.document_xml
    )
    assert len(changes) == 2
    assert plan.document_xml.count('w:date="2026-01-01T00:00:00Z"') == 4
    ids = re.findall(r'w:id="(9\d+)"', plan.document_xml)
    assert len(ids) == len(set(ids)), "revision ids must be unique"


def test_markers_are_plain_text_when_tracking_is_off() -> None:
    """Untracked stays untracked: the marker joins the paragraph's own run."""

    plan = _convert_tracked([("PART", "GENERAL")], _TRACKING_OFF)
    assert plan.report.source_tracks_revisions is False
    assert plan.report.markers_tracked is False
    assert plan.report.marker_author is None
    assert "<w:ins" not in plan.document_xml


def test_tracking_absent_from_settings_is_not_tracking() -> None:
    plan = _convert_tracked([("PART", "GENERAL")], "")
    assert plan.report.source_tracks_revisions is False


def _reject_all(document_xml: str) -> str:
    """Apply Word's reject-all to this application's revisions.

    Rejecting drops the content of a ``w:ins`` and restores the properties a
    ``w:pPrChange`` recorded -- both halves, which is the whole point of the
    test below.

    Splits paragraphs by regex, which is sound only because these fixtures
    contain no text boxes. A real document nests ``w:p`` inside ``w:p`` through
    ``w:txbxContent``, and a lazy ``</w:p>`` stops at the inner one. Verify a
    real package over the DOM, not with this.
    """

    without_insertions = re.sub(
        r"<w:ins\b[^>]*>.*?</w:ins>", "", document_xml, flags=re.S
    )

    # Restore per paragraph, never across the document. A lazy match anchored
    # on w:pPrChange will happily start at one paragraph's w:pPr and end at a
    # later paragraph's, deleting everything between -- which looks like a
    # clean pass right up until the indices no longer line up.
    def _restore_paragraph(paragraph: re.Match) -> str:
        return re.sub(
            r"<w:pPr\b[^>]*>.*?<w:pPrChange\b[^>]*>"
            r"(<w:pPr\b[^>]*(?:/>|>.*?</w:pPr>))</w:pPrChange></w:pPr>",
            lambda match: match.group(1),
            paragraph.group(0),
            flags=re.S,
        )

    return re.sub(
        r"<w:p\b[^>]*(?:/>|>.*?</w:p>)",
        _restore_paragraph,
        without_insertions,
        flags=re.S,
    )


def test_rejecting_the_tracked_conversion_restores_text_and_numbering() -> None:
    """The whole conversion is reversible in Word, which is the point.

    Text alone is not the test. The numbering suppression is a *property*
    edit, so a conversion that tracked only its marker would pass a text
    comparison while leaving every paragraph unnumbered on rejection -- worse
    than either the source or the output.
    """

    rows = [("PART", "GENERAL"), ("ARTICLE", "SUMMARY"), ("PARAGRAPH", "Body text.")]
    plan = _convert_tracked(rows, _TRACKING_ON)
    rejected = _reject_all(plan.document_xml)

    blocks = [block for _s, _e, block in iter_paragraph_xml_blocks(rejected)]
    assert [paragraph_text_from_block(b) for b in blocks] == [t for _r, t in rows]
    for block in blocks:
        assert "<w:numPr" not in block, block
        assert "<w:ind" not in block, block
    assert [
        re.search(r'<w:pStyle w:val="([^"]+)"', b).group(1) for b in blocks
    ] == [_GEOMETRY_ROLE_STYLE[role] for role, _text in rows]


def test_tracked_marker_run_precedes_the_original_run() -> None:
    """The revision wraps its own run and is placed before, never inside.

    A marker inside the paragraph's existing run could not be rejected
    independently of the text around it.
    """

    plan = _convert_tracked([("PART", "GENERAL")], _TRACKING_ON)
    body = plan.document_xml
    assert body.index("<w:ins") < body.index("GENERAL")
    assert re.search(r"<w:ins\b[^>]*><w:r><w:t[^>]*>PART 1</w:t><w:tab/></w:r></w:ins>", body)


# --- Predict-first --------------------------------------------------------


def _blocks(*texts: str):
    """``(start, end, xml)`` triples shaped like ``iter_paragraph_xml_blocks``."""

    return [
        (0, 0, f'<w:p><w:r><w:t xml:space="preserve">{text}</w:t></w:r></w:p>')
        for text in texts
    ]


def test_prediction_catches_a_marker_that_was_never_written() -> None:
    """A predicted paragraph that did not receive its marker fails closed.

    The per-paragraph check inside the edit loop cannot see this: it compares
    the marker it just wrote against the paragraph it just wrote it into, so it
    is silent about a paragraph the loop never reached.
    """

    with pytest.raises(EngineError) as raised:
        _verify_prediction(
            _blocks("GENERAL", "SUMMARY"),
            _blocks("PART 1\tGENERAL", "SUMMARY"),
            [(0, "PART", "PART 1"), (1, "ARTICLE", "1.1")],
            describe=lambda index: "",
        )
    assert raised.value.code == "conversion_prediction_mismatch"
    assert "predicted to lead with" in str(raised.value)


def test_prediction_catches_an_edit_nobody_asked_for() -> None:
    """A paragraph outside the prediction must not change text at all."""

    with pytest.raises(EngineError) as raised:
        _verify_prediction(
            _blocks("GENERAL", "Untouched prose."),
            _blocks("PART 1\tGENERAL", "Quietly rewritten."),
            [(0, "PART", "PART 1")],
            describe=lambda index: "",
        )
    assert raised.value.code == "conversion_prediction_mismatch"
    assert "no CSI marker was predicted" in str(raised.value)


def test_prediction_accepts_the_document_it_predicted() -> None:
    _verify_prediction(
        _blocks("GENERAL", "Untouched prose."),
        _blocks("PART 1\tGENERAL", "Untouched prose."),
        [(0, "PART", "PART 1")],
        describe=lambda index: "",
    )


def test_unchanged_text_still_satisfies_its_prediction() -> None:
    """A typed ``PART 1`` converting to ``PART 1`` changes nothing, correctly.

    The check therefore tests what the paragraph leads with, not whether it
    moved -- a prediction keyed on "did this change" would fail every document
    whose Canadian and CSI markers happen to coincide.
    """

    rows = [("PART", "PART 1", "GENERAL"), ("ARTICLE", "1.1", "SUMMARY")]
    document = _document(
        [
            "<w:p><w:pPr/>"
            f'<w:r><w:t xml:space="preserve">{marker}</w:t></w:r><w:r><w:tab/></w:r>'
            f'<w:r><w:t xml:space="preserve">{text}</w:t></w:r></w:p>'
            for _role, marker, text in rows
        ]
    )
    plan = plan_canadian_to_csi(
        document,
        builtin_scheme.build_styles_xml(),
        _classifications([role for role, _m, _t in rows]),
        numbering_xml="",
    )
    texts = [
        paragraph_text_from_block(block)
        for _s, _e, block in iter_paragraph_xml_blocks(plan.document_xml)
    ]
    assert texts[0].startswith("PART 1")
    assert texts[1].startswith("1.1")


def test_typed_marker_replacement_under_tracking_fails_closed() -> None:
    """Half a tracked edit is the defect this mode exists to avoid.

    The marker insertion can be tracked; the removal of the typed marker it
    replaces cannot, without restructuring shared code both converters use.
    Publishing a reviewable insertion beside an unrecorded deletion would be
    exactly the silent, untracked text change the tracked path was added to
    prevent, so the target is refused with an actionable remedy instead.
    """

    rows = [("PART", "PART 1", "GENERAL"), ("ARTICLE", "1.1", "SUMMARY")]
    document = _document(
        [
            "<w:p><w:pPr/>"
            f'<w:r><w:t xml:space="preserve">{marker}</w:t></w:r><w:r><w:tab/></w:r>'
            f'<w:r><w:t xml:space="preserve">{text}</w:t></w:r></w:p>'
            for _role, marker, text in rows
        ]
    )
    with pytest.raises(EngineError) as raised:
        plan_canadian_to_csi(
            document,
            builtin_scheme.build_styles_xml(),
            _classifications([role for role, _m, _t in rows]),
            numbering_xml="",
            settings_xml=_TRACKING_ON,
        )
    assert raised.value.code == "canadian_to_csi_tracked_hierarchy"
    assert "Turn Track Changes off" in str(raised.value)


def test_tracked_marker_inherits_the_formatting_of_the_run_it_precedes() -> None:
    """A bold heading gets a bold number, tracked or not.

    The untracked path gets this by joining the existing run. A tracked marker
    must be its own run, and a bare one renders in the document defaults --
    which the run-property invariant cannot catch, because it removes this
    application's insertions before comparing.
    """

    paragraphs = [
        '<w:p><w:pPr><w:pStyle w:val="Geo0"/></w:pPr>'
        '<w:r><w:rPr><w:b/><w:sz w:val="28"/></w:rPr>'
        '<w:t xml:space="preserve">GENERAL</w:t></w:r></w:p>',
    ]
    plan = plan_canadian_to_csi(
        _document(paragraphs),
        _geometry_styles_xml(),
        _classifications(["PART"]),
        numbering_xml=_geometry_numbering_xml(),
        settings_xml=_TRACKING_ON,
        revision_date="2026-01-01T00:00:00Z",
    )
    marker_run = re.search(r"<w:ins\b[^>]*>(.*?)</w:ins>", plan.document_xml, re.S)
    assert marker_run is not None
    assert "<w:b/>" in marker_run.group(1)
    assert '<w:sz w:val="28"/>' in marker_run.group(1)


def test_tracked_marker_run_is_bare_when_the_source_run_is() -> None:
    """Nothing is invented: a run with no properties contributes none."""

    plan = _convert_tracked([("PART", "GENERAL")], _TRACKING_ON)
    marker_run = re.search(r"<w:ins\b[^>]*>(.*?)</w:ins>", plan.document_xml, re.S)
    assert marker_run is not None
    assert "<w:rPr" not in marker_run.group(1)


def test_reject_simulation_does_not_swallow_neighbouring_paragraphs() -> None:
    """Guard on the test helper itself.

    A document-wide lazy match anchored on ``w:pPrChange`` will start at one
    paragraph's ``w:pPr`` and end at a later paragraph's, deleting everything
    between. That reads as a clean pass until the paragraph indices stop lining
    up, so the helper restores per paragraph and this proves it.
    """

    unconverted = '<w:p><w:pPr><w:pStyle w:val="Geo2"/></w:pPr>'\
                  '<w:r><w:t xml:space="preserve">Front matter.</w:t></w:r></w:p>'
    plan = _convert_tracked([("PART", "GENERAL"), ("ARTICLE", "SUMMARY")], _TRACKING_ON)
    document = plan.document_xml.replace("<w:body>", f"<w:body>{unconverted}", 1)

    rejected = _reject_all(document)
    texts = [
        paragraph_text_from_block(block)
        for _s, _e, block in iter_paragraph_xml_blocks(rejected)
    ]
    assert texts == ["Front matter.", "GENERAL", "SUMMARY"]
