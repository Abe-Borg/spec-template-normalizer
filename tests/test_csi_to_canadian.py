from __future__ import annotations

from pathlib import Path

import pytest

from spec_formatter.style_application.core.csi_to_canadian import (
    CSI_TO_CANADIAN,
    FORMAT_ONLY,
    apply_csi_to_canadian,
    classifications_for_canadian_application,
    plan_csi_to_canadian,
    validate_conversion_mode,
)
from spec_formatter.numbering_roles import role_from_numbering_signature
from spec_formatter.style_application.core.classification import (
    apply_phase2_classifications,
)
from spec_formatter.style_application.core.xml_helpers import (
    iter_paragraph_xml_blocks,
    paragraph_pstyle_from_block,
    paragraph_text_from_block,
)


W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"


def _document(*paragraphs: str) -> str:
    return (
        '<?xml version="1.0" encoding="UTF-8"?>'
        f'<w:document xmlns:w="{W_NS}"><w:body>'
        + "".join(paragraphs)
        + "<w:sectPr/></w:body></w:document>"
    )


def _paragraph(text: str, ppr: str = "") -> str:
    props = f"<w:pPr>{ppr}</w:pPr>" if ppr else ""
    return f"<w:p>{props}<w:r><w:t>{text}</w:t></w:r></w:p>"


def _styles(extra: str = "") -> str:
    return f'<w:styles xmlns:w="{W_NS}">{extra}</w:styles>'


def _source_numbering(*num_ids: str, override: str = "") -> str:
    nums = "".join(
        f'<w:num w:numId="{num_id}"><w:abstractNumId w:val="{num_id}"/>'
        f'{override}</w:num>'
        for num_id in num_ids
    )
    abstracts = "".join(
        f'<w:abstractNum w:abstractNumId="{num_id}">'
        '<w:lvl w:ilvl="0"><w:start w:val="1"/>'
        '<w:numFmt w:val="upperLetter"/><w:lvlText w:val="%1."/>'
        '</w:lvl></w:abstractNum>'
        for num_id in num_ids
    )
    return f'<w:numbering xmlns:w="{W_NS}">{abstracts}{nums}</w:numbering>'


def _architect_numbering(
    *,
    start: str = "1",
    restart: str = "",
) -> str:
    return (
        f'<w:numbering xmlns:w="{W_NS}">'
        '<w:abstractNum w:abstractNumId="9">'
        f'<w:lvl w:ilvl="2"><w:start w:val="{start}"/>{restart}'
        '<w:numFmt w:val="decimal"/><w:lvlText w:val=".%3"/>'
        '</w:lvl></w:abstractNum>'
        '<w:num w:numId="9"><w:abstractNumId w:val="9"/></w:num>'
        '</w:numbering>'
    )


def _canadian_role_specs(*roles: str):
    patterns = {
        "PART": "PART %1",
        "ARTICLE": "%1.%2",
        "PARAGRAPH": ".%3",
        "SUBPARAGRAPH": ".%4",
        "SUBSUBPARAGRAPH": ".%5",
        "SUBPARAGRAPH_LEVEL_5": ".%6",
        "SUBPARAGRAPH_LEVEL_6": ".%7",
        "SUBPARAGRAPH_LEVEL_7": ".%8",
        "SUBPARAGRAPH_LEVEL_8": ".%9",
    }
    levels = {
        "PART": "0",
        "ARTICLE": "1",
        "PARAGRAPH": "2",
        "SUBPARAGRAPH": "3",
        "SUBSUBPARAGRAPH": "4",
        "SUBPARAGRAPH_LEVEL_5": "5",
        "SUBPARAGRAPH_LEVEL_6": "6",
        "SUBPARAGRAPH_LEVEL_7": "7",
        "SUBPARAGRAPH_LEVEL_8": "8",
    }
    return {
        role: {
            "style_id": f"Canadian{role.title()}",
            "numbering_provenance": "style_numpr",
            "numbering_pattern": {
                "numId": "9",
                "ilvl": levels[role],
                "numFmt": "decimal",
                "lvlText": patterns[role],
            },
        }
        for role in roles
    }


def _classifications(*roles: str):
    return {
        "classifications": [
            {"paragraph_index": index, "csi_role": role}
            for index, role in enumerate(roles)
        ],
        "notes": [],
    }


def _texts(document_xml: str) -> list[str]:
    return [
        paragraph_text_from_block(block)
        for _start, _end, block in iter_paragraph_xml_blocks(document_xml)
    ]


def test_mode_validation_is_strict_and_backwards_compatible():
    assert validate_conversion_mode(FORMAT_ONLY) == FORMAT_ONLY
    assert validate_conversion_mode(CSI_TO_CANADIAN) == CSI_TO_CANADIAN
    with pytest.raises(ValueError, match="conversion_mode"):
        validate_conversion_mode("canadian-ish")


def test_literal_csi_hierarchy_is_removed_for_architect_automatic_numbering():
    roles = ("PART", "ARTICLE", "PARAGRAPH", "SUBPARAGRAPH", "SUBSUBPARAGRAPH")
    source = _document(
        _paragraph("PART 1 - GENERAL"),
        _paragraph("1.01 SUMMARY"),
        _paragraph("A. Work Included"),
        _paragraph("1. Pumps"),
        _paragraph("a. Steel"),
    )

    plan = plan_csi_to_canadian(
        source,
        _styles(),
        _classifications(*roles),
        _canadian_role_specs(*roles),
    )

    assert _texts(plan.document_xml)[:5] == [
        "GENERAL",
        "SUMMARY",
        "Work Included",
        "Pumps",
        "Steel",
    ]
    assert _texts(source)[:5] == [
        "PART 1 - GENERAL",
        "1.01 SUMMARY",
        "A. Work Included",
        "1. Pumps",
        "a. Steel",
    ]
    assert plan.report.literal_markers_removed == 5
    assert plan.report.paragraphs_converted == 5
    assert [edit.source_marker for edit in plan.report.edits] == [
        "PART 1", "1.01", "A.", "1.", "a."
    ]


def test_fifth_level_csi_marker_converts_to_sixth_pageformat_level():
    roles = (
        "PART",
        "ARTICLE",
        "PARAGRAPH",
        "SUBPARAGRAPH",
        "SUBSUBPARAGRAPH",
        "SUBPARAGRAPH_LEVEL_5",
    )
    source = _document(
        _paragraph("PART 1 GENERAL"),
        _paragraph("1.01 SUMMARY"),
        _paragraph("A. Work Included"),
        _paragraph("1. Pumps"),
        _paragraph("a. Steel"),
        _paragraph("1) Factory welds"),
    )
    specs = _canadian_role_specs(*roles)
    # The failing architect sample uses numeric CSC Part numbering rather than
    # a literal PART prefix in its automatic level text.
    specs["PART"]["numbering_pattern"]["lvlText"] = "%1."

    plan = plan_csi_to_canadian(
        source,
        _styles(),
        _classifications(*roles),
        specs,
    )

    assert _texts(plan.document_xml)[:6] == [
        "GENERAL",
        "SUMMARY",
        "Work Included",
        "Pumps",
        "Steel",
        "Factory welds",
    ]
    assert plan.report.literal_markers_removed == 6
    assert plan.report.edits[-1].role == "SUBPARAGRAPH_LEVEL_5"
    assert plan.report.edits[-1].source_marker == "1)"


def test_canonical_counters_restart_under_each_new_parent():
    roles = (
        "PART",
        "ARTICLE",
        "PARAGRAPH",
        "PARAGRAPH",
        "PART",
        "ARTICLE",
        "PARAGRAPH",
    )
    source = _document(
        _paragraph("PART 1 GENERAL"),
        _paragraph("1.01 SUMMARY"),
        _paragraph("A. First"),
        _paragraph("B. Second"),
        _paragraph("PART 2 PRODUCTS"),
        _paragraph("2.01 EQUIPMENT"),
        _paragraph("A. First product"),
    )

    plan = plan_csi_to_canadian(
        source,
        _styles(),
        _classifications(*roles),
        _canadian_role_specs(*dict.fromkeys(roles)),
    )

    assert _texts(plan.document_xml)[:7] == [
        "GENERAL",
        "SUMMARY",
        "First",
        "Second",
        "PRODUCTS",
        "EQUIPMENT",
        "First product",
    ]


def test_split_run_marker_tab_and_escaped_content_are_handled_without_losing_formatting():
    source = _document(
        '<w:p><w:r><w:rPr><w:b/></w:rPr><w:t>A</w:t></w:r>'
        '<w:r><w:t>.</w:t><w:tab/></w:r>'
        '<w:r><w:rPr><w:i/></w:rPr><w:t>Scope &amp; coordination</w:t></w:r></w:p>'
    )

    plan = plan_csi_to_canadian(
        source,
        _styles(),
        _classifications("PARAGRAPH"),
        _canadian_role_specs("PARAGRAPH"),
    )

    assert _texts(plan.document_xml)[0] == "Scope & coordination"
    assert "<w:tab" not in plan.document_xml
    assert "<w:b/>" in plan.document_xml
    assert "<w:i/>" in plan.document_xml
    assert "Scope &amp; coordination" in plan.document_xml


def test_part_marker_followed_by_word_tab_is_removed():
    source = _document(
        '<w:p><w:r><w:t>PART 1</w:t><w:tab/></w:r>'
        '<w:r><w:t>GENERAL</w:t></w:r></w:p>'
    )

    plan = plan_csi_to_canadian(
        source,
        _styles(),
        _classifications("PART"),
        _canadian_role_specs("PART"),
    )

    assert _texts(plan.document_xml)[0] == "GENERAL"
    assert "<w:tab" not in plan.document_xml


def test_article_requires_part_and_one_coherent_canadian_multilevel_list():
    with pytest.raises(ValueError, match="requires a classified PART"):
        plan_csi_to_canadian(
            _document(_paragraph("1.01 SUMMARY")),
            _styles(),
            _classifications("ARTICLE"),
            _canadian_role_specs("ARTICLE"),
        )

    roles = ("PART", "ARTICLE")
    split_lists = _canadian_role_specs(*roles)
    split_lists["ARTICLE"]["numbering_pattern"]["numId"] = "10"
    with pytest.raises(ValueError, match="share one Word multilevel numbering list"):
        plan_csi_to_canadian(
            _document(_paragraph("PART 1 GENERAL"), _paragraph("1.01 SUMMARY")),
            _styles(),
            _classifications(*roles),
            split_lists,
        )


def test_automatic_csi_numbering_is_retargeted_without_text_edits():
    style = (
        '<w:style w:type="paragraph" w:styleId="TargetAlpha"><w:pPr><w:numPr>'
        '<w:ilvl w:val="0"/><w:numId w:val="7"/>'
        '</w:numPr></w:pPr></w:style>'
    )
    source = _document(
        _paragraph("Scope", '<w:pStyle w:val="TargetAlpha"/>')
    )

    plan = plan_csi_to_canadian(
        source,
        _styles(style),
        _classifications("PARAGRAPH"),
        _canadian_role_specs("PARAGRAPH"),
        numbering_xml=_source_numbering("7"),
    )

    assert plan.document_xml == source
    assert plan.report.automatic_numbering_retargeted == 1
    assert plan.report.literal_markers_removed == 0
    assert plan.report.edits[0].source_kind == "automatic"


def test_mixed_automatic_and_literal_numbering_fails_closed():
    source = _document(
        _paragraph(
            "A. Scope",
            '<w:numPr><w:ilvl w:val="0"/><w:numId w:val="7"/></w:numPr>',
        )
    )
    with pytest.raises(ValueError, match="both automatic numbering and typed marker"):
        plan_csi_to_canadian(
            source,
            _styles(),
            _classifications("PARAGRAPH"),
            _canadian_role_specs("PARAGRAPH"),
        )


def test_incompatible_marker_role_and_non_canadian_template_fail_closed():
    source = _document(_paragraph("a. Wrong level"))
    with pytest.raises(ValueError, match="incompatible marker"):
        plan_csi_to_canadian(
            source,
            _styles(),
            _classifications("PARAGRAPH"),
            _canadian_role_specs("PARAGRAPH"),
        )

    with pytest.raises(ValueError, match="incompatible marker"):
        plan_csi_to_canadian(
            _document(_paragraph("A.B. designation is unchanged technical text")),
            _styles(),
            _classifications("PARAGRAPH"),
            _canadian_role_specs("PARAGRAPH"),
        )

    american = _canadian_role_specs("PARAGRAPH")
    american["PARAGRAPH"]["numbering_pattern"].update(
        {"numFmt": "upperLetter", "lvlText": "%1."}
    )
    with pytest.raises(ValueError, match="not Canadian numeric"):
        plan_csi_to_canadian(
            _document(_paragraph("A. Scope")),
            _styles(),
            _classifications("PARAGRAPH"),
            american,
        )


def test_unclassified_marker_like_text_and_section_properties_are_byte_exact():
    source = _document(
        _paragraph("A. This resembles a marker but is not classified"),
        _paragraph("A. Classified requirement"),
    )
    classifications = {
        "classifications": [{"paragraph_index": 1, "csi_role": "PARAGRAPH"}],
        "notes": [],
    }
    plan = plan_csi_to_canadian(
        source,
        _styles(),
        classifications,
        _canadian_role_specs("PARAGRAPH"),
    )
    before_blocks = list(iter_paragraph_xml_blocks(source))
    after_blocks = list(iter_paragraph_xml_blocks(plan.document_xml))
    assert before_blocks[0][2] == after_blocks[0][2]
    assert "<w:sectPr/>" in plan.document_xml


def test_tracked_marker_and_multi_paragraph_failure_do_not_partially_write(tmp_path: Path):
    extract = tmp_path / "extract"
    (extract / "word").mkdir(parents=True)
    document_path = extract / "word" / "document.xml"
    document_path.write_text(
        _document(
            _paragraph("A. Valid first paragraph"),
            '<w:p><w:ins><w:r><w:t>B. Tracked marker</w:t></w:r></w:ins></w:p>',
        ),
        encoding="utf-8",
    )
    (extract / "word" / "styles.xml").write_text(_styles(), encoding="utf-8")
    original = document_path.read_bytes()

    with pytest.raises(ValueError, match="tracked-change/field"):
        apply_csi_to_canadian(
            extract,
            _classifications("PARAGRAPH", "PARAGRAPH"),
            _canadian_role_specs("PARAGRAPH"),
            [],
        )

    assert document_path.read_bytes() == original


def test_typed_canadian_marker_requires_word_tab_and_canonical_counter():
    source = _document(
        '<w:p><w:r><w:t>.1</w:t><w:tab/></w:r>'
        '<w:r><w:t>Existing Canadian item</w:t></w:r></w:p>'
    )
    plan = plan_csi_to_canadian(
        source,
        _styles(),
        _classifications("PARAGRAPH"),
        _canadian_role_specs("PARAGRAPH"),
    )
    assert _texts(plan.document_xml)[0] == "Existing Canadian item"
    assert plan.report.edits[0].source_kind == "already_canadian"

    with pytest.raises(ValueError, match="ambiguous decimal text"):
        plan_csi_to_canadian(
            _document(_paragraph(".125 mm thick")),
            _styles(),
            _classifications("PARAGRAPH"),
            _canadian_role_specs("PARAGRAPH"),
        )


def test_spaced_canadian_articles_convert_when_part_sequence_proves_them():
    roles = ("PART", "ARTICLE", "ARTICLE")
    source = _document(
        _paragraph("PART 1 GENERAL"),
        _paragraph("1.1 SUMMARY"),
        _paragraph("1.2 REFERENCES"),
    )

    plan = plan_csi_to_canadian(
        source,
        _styles(),
        _classifications(*roles),
        _canadian_role_specs(*dict.fromkeys(roles)),
    )

    assert _texts(plan.document_xml)[:3] == ["GENERAL", "SUMMARY", "REFERENCES"]
    assert [edit.source_kind for edit in plan.report.edits] == [
        "literal",
        "already_canadian",
        "already_canadian",
    ]


def test_spaced_decimal_measurement_is_not_accepted_as_canadian_article():
    roles = ("PART", "ARTICLE")
    with pytest.raises(ValueError, match="ambiguous decimal text"):
        plan_csi_to_canadian(
            _document(
                _paragraph("PART 1 GENERAL"),
                _paragraph("1.1 mm thick"),
            ),
            _styles(),
            _classifications(*roles),
            _canadian_role_specs(*roles),
        )


def test_deep_numbering_signatures_are_not_deterministically_misclassified():
    assert role_from_numbering_signature("decimal", "%1.%2", "1") == "ARTICLE"
    assert role_from_numbering_signature("decimal", "%1.%2.%3", "2") is None
    assert role_from_numbering_signature("decimal", ".%3", "2") is None


@pytest.mark.parametrize(
    ("paragraphs", "roles", "message"),
    [
        (
            (_paragraph("A. First"), _paragraph("C. Gap")),
            ("PARAGRAPH", "PARAGRAPH"),
            "non-contiguous PARAGRAPH",
        ),
        (
            (_paragraph("PART 2 PRODUCTS"),),
            ("PART",),
            "expected counter 1",
        ),
        (
            (_paragraph("1. Orphan"),),
            ("SUBPARAGRAPH",),
            "without a preceding PARAGRAPH",
        ),
    ],
)
def test_noncanonical_typed_sequences_fail_closed(paragraphs, roles, message):
    with pytest.raises(ValueError, match=message):
        plan_csi_to_canadian(
            _document(*paragraphs),
            _styles(),
            _classifications(*roles),
            _canadian_role_specs(*dict.fromkeys(roles)),
        )


def test_unnumbered_classification_is_not_inserted_into_known_sequence(tmp_path: Path):
    classifications = _classifications("PARAGRAPH", "PARAGRAPH", "PARAGRAPH")
    plan = plan_csi_to_canadian(
        _document(
            _paragraph("A. First"),
            _paragraph("Unproven list item"),
            _paragraph("B. Second"),
        ),
        _styles(),
        classifications,
        _canadian_role_specs("PARAGRAPH"),
    )

    assert _texts(plan.document_xml)[:3] == [
        "First",
        "Unproven list item",
        "Second",
    ]
    assert len(plan.report.warnings) == 1
    assert plan.report.warnings[0].paragraph_index == 1
    assert plan.report.warnings[0].code == "unproven_numbered_role_preserved"

    application = classifications_for_canadian_application(
        classifications,
        plan.report,
    )
    assert [
        item["paragraph_index"] for item in application["classifications"]
    ] == [0, 2]
    assert len(classifications["classifications"]) == 3

    extract = tmp_path / "extract"
    (extract / "word").mkdir(parents=True)
    (extract / "word" / "document.xml").write_text(
        plan.document_xml,
        encoding="utf-8",
    )
    style = (
        '<w:style w:type="paragraph" w:styleId="CanadianParagraph">'
        '<w:pPr><w:numPr><w:ilvl w:val="2"/><w:numId w:val="9"/>'
        '</w:numPr></w:pPr></w:style>'
    )
    (extract / "word" / "styles.xml").write_text(
        _styles(style),
        encoding="utf-8",
    )
    apply_report = apply_phase2_classifications(
        extract,
        application,
        {"PARAGRAPH": "CanadianParagraph"},
        [],
        role_specs=_canadian_role_specs("PARAGRAPH"),
    )

    applied_xml = (extract / "word" / "document.xml").read_text(encoding="utf-8")
    applied_blocks = list(iter_paragraph_xml_blocks(applied_xml))
    assert apply_report.modified == 2
    assert paragraph_pstyle_from_block(applied_blocks[0][2]) == "CanadianParagraph"
    assert paragraph_pstyle_from_block(applied_blocks[1][2]) is None
    assert paragraph_pstyle_from_block(applied_blocks[2][2]) == "CanadianParagraph"


def test_automatic_start_override_and_list_instance_change_fail_closed():
    style = (
        '<w:style w:type="paragraph" w:styleId="Auto"><w:pPr><w:numPr>'
        '<w:ilvl w:val="0"/><w:numId w:val="7"/>'
        '</w:numPr></w:pPr></w:style>'
    )
    source = _document(_paragraph("Scope", '<w:pStyle w:val="Auto"/>'))
    override = (
        '<w:lvlOverride w:ilvl="0"><w:startOverride w:val="5"/>'
        '</w:lvlOverride>'
    )
    with pytest.raises(ValueError, match="list-level override"):
        plan_csi_to_canadian(
            source,
            _styles(style),
            _classifications("PARAGRAPH"),
            _canadian_role_specs("PARAGRAPH"),
            numbering_xml=_source_numbering("7", override=override),
        )

    two_lists = _document(
        _paragraph(
            "First",
            '<w:numPr><w:ilvl w:val="0"/><w:numId w:val="7"/></w:numPr>',
        ),
        _paragraph(
            "Second",
            '<w:numPr><w:ilvl w:val="0"/><w:numId w:val="8"/></w:numPr>',
        ),
    )
    with pytest.raises(ValueError, match="changes list instances"):
        plan_csi_to_canadian(
            two_lists,
            _styles(),
            _classifications("PARAGRAPH", "PARAGRAPH"),
            _canadian_role_specs("PARAGRAPH"),
            numbering_xml=_source_numbering("7", "8"),
        )


def test_unconverted_paragraph_sharing_automatic_list_is_rejected():
    numpr = '<w:numPr><w:ilvl w:val="0"/><w:numId w:val="7"/></w:numPr>'
    source = _document(
        _paragraph("First", numpr),
        _paragraph("Filtered table or boilerplate item", numpr),
        _paragraph("Third", numpr),
    )
    classifications = {
        "classifications": [
            {"paragraph_index": 0, "csi_role": "PARAGRAPH"},
            {"paragraph_index": 2, "csi_role": "PARAGRAPH"},
        ],
        "notes": [],
    }

    with pytest.raises(ValueError, match="Unconverted paragraph 1 shares"):
        plan_csi_to_canadian(
            source,
            _styles(),
            classifications,
            _canadian_role_specs("PARAGRAPH"),
            numbering_xml=_source_numbering("7"),
        )


def test_dependent_automatic_roles_must_share_one_source_list_instance():
    source = _document(
        _paragraph(
            "GENERAL",
            '<w:numPr><w:ilvl w:val="0"/><w:numId w:val="7"/></w:numPr>',
        ),
        _paragraph(
            "SUMMARY",
            '<w:numPr><w:ilvl w:val="1"/><w:numId w:val="8"/></w:numPr>',
        ),
    )
    numbering = (
        f'<w:numbering xmlns:w="{W_NS}">'
        '<w:abstractNum w:abstractNumId="7"><w:lvl w:ilvl="0">'
        '<w:start w:val="1"/><w:numFmt w:val="decimal"/>'
        '<w:lvlText w:val="PART %1"/></w:lvl></w:abstractNum>'
        '<w:abstractNum w:abstractNumId="8"><w:lvl w:ilvl="1">'
        '<w:start w:val="1"/><w:numFmt w:val="decimalZero"/>'
        '<w:lvlText w:val="%1.%2"/></w:lvl></w:abstractNum>'
        '<w:num w:numId="7"><w:abstractNumId w:val="7"/></w:num>'
        '<w:num w:numId="8"><w:abstractNumId w:val="8"/></w:num>'
        '</w:numbering>'
    )

    with pytest.raises(ValueError, match="different Word list instances"):
        plan_csi_to_canadian(
            source,
            _styles(),
            _classifications("PART", "ARTICLE"),
            _canadian_role_specs("PART", "ARTICLE"),
            numbering_xml=numbering,
        )


def test_direct_numid_zero_suppresses_inherited_numbering():
    style = (
        '<w:style w:type="paragraph" w:styleId="Auto"><w:pPr><w:numPr>'
        '<w:ilvl w:val="0"/><w:numId w:val="7"/>'
        '</w:numPr></w:pPr></w:style>'
    )
    source = _document(
        _paragraph(
            "A. Scope",
            '<w:pStyle w:val="Auto"/><w:numPr><w:numId w:val="0"/></w:numPr>',
        )
    )
    plan = plan_csi_to_canadian(
        source,
        _styles(style),
        _classifications("PARAGRAPH"),
        _canadian_role_specs("PARAGRAPH"),
    )
    assert _texts(plan.document_xml)[0] == "Scope"


def test_line_break_after_marker_is_rejected_without_reflowing_content():
    source = _document(
        '<w:p><w:r><w:t>A.</w:t><w:br/><w:t>Scope</w:t></w:r></w:p>'
    )
    with pytest.raises(ValueError, match="line break after typed marker"):
        plan_csi_to_canadian(
            source,
            _styles(),
            _classifications("PARAGRAPH"),
            _canadian_role_specs("PARAGRAPH"),
        )


def test_architect_counter_start_and_restart_are_validated_from_numbering_xml():
    kwargs = {
        "document_xml": _document(_paragraph("A. Scope")),
        "styles_xml": _styles(),
        "classifications": _classifications("PARAGRAPH"),
        "role_specs": _canadian_role_specs("PARAGRAPH"),
    }
    with pytest.raises(ValueError, match="starts at '5'"):
        plan_csi_to_canadian(
            **kwargs,
            architect_numbering_xml=_architect_numbering(start="5"),
        )
    with pytest.raises(ValueError, match="explicit restart rule"):
        plan_csi_to_canadian(
            **kwargs,
            architect_numbering_xml=_architect_numbering(
                restart='<w:lvlRestart w:val="1"/>',
            ),
        )
