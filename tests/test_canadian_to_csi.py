"""Unit coverage for the Canadian -> CSI converter.

Writing a number is a stronger claim than removing one: a wrong marker becomes
literal text in a document a reader will trust. Most of what follows therefore
tests the refusals rather than the conversions.
"""

from __future__ import annotations

import pytest

from spec_formatter import builtin_scheme
from spec_formatter.role_contract import ROLE_TO_ARCH_STYLE
from spec_formatter.style_application.core.canadian_to_csi import (
    _csi_marker,
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


def test_article_number_follows_the_active_part() -> None:
    assert _csi_marker("ARTICLE", {0: 3, 1: 7}) == "3.7"
    assert _csi_marker("PART", {0: 4}) == "PART 4"
