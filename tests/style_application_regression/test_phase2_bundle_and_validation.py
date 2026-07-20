import json
from pathlib import Path

import pytest

from spec_formatter.style_application.core.classification import (
    build_phase2_slim_bundle,
    preclassify_paragraphs,
    validate_phase2_classification_contract,
    validate_phase2_final_payload,
    coerce_to_final_classifications,
    PHASE2_MASTER_PROMPT,
    PHASE2_RUN_INSTRUCTION,
)


def _write_document_xml(tmp_path: Path, body_xml: str) -> Path:
    word = tmp_path / "word"
    word.mkdir(parents=True)
    doc_xml = (
        '<?xml version="1.0" encoding="UTF-8"?>\n'
        '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">\n'
        f'<w:body>{body_xml}</w:body></w:document>'
    )
    (word / "document.xml").write_text(doc_xml, encoding="utf-8")
    return tmp_path


def test_prompts_loaded_from_files():
    assert "CSI STRUCTURE CLASSIFIER" in PHASE2_MASTER_PROMPT
    assert "Output schema" in PHASE2_RUN_INSTRUCTION


def test_bundle_enrichment_and_table_filtering(tmp_path: Path):
    extract_dir = _write_document_xml(
        tmp_path,
        (
            '<w:p><w:pPr><w:pStyle w:val="Body"/><w:ind w:left="720"/><w:spacing w:after="120"/></w:pPr>'
            '<w:r><w:rPr><w:b/></w:rPr><w:t>PART 1 GENERAL</w:t></w:r></w:p>'
            '<w:tbl><w:tr><w:tc><w:p><w:r><w:t>A. In table</w:t></w:r></w:p></w:tc></w:tr></w:tbl>'
            '<w:p><w:r><w:t>1.01 SUMMARY</w:t></w:r></w:p>'
        ),
    )

    bundle = build_phase2_slim_bundle(extract_dir)
    assert bundle["deterministic_classifications"]
    assert len(bundle["paragraphs"]) == 0


def test_nested_tables_are_fully_out_of_scope(tmp_path: Path):
    extract_dir = _write_document_xml(
        tmp_path,
        (
            '<w:tbl><w:tr><w:tc>'
            '<w:p><w:r><w:t>A. Outer cell</w:t></w:r></w:p>'
            '<w:tbl><w:tr><w:tc><w:p><w:r><w:t>1. Nested cell</w:t></w:r></w:p>'
            '</w:tc></w:tr></w:tbl>'
            '</w:tc></w:tr></w:tbl>'
            '<w:p><w:r><w:t>Ordinary requirement</w:t></w:r></w:p>'
        ),
    )

    bundle = build_phase2_slim_bundle(extract_dir, available_roles=["PARAGRAPH"])
    assert [p["paragraph_index"] for p in bundle["paragraphs"]] == [2]
    assert bundle["filter_report"]["paragraphs_out_of_scope"] == [
        {
            "paragraph_index": 0,
            "reason": "table",
            "original_text_preview": "A. Outer cell",
        },
        {
            "paragraph_index": 1,
            "reason": "table",
            "original_text_preview": "1. Nested cell",
        },
    ]


def test_drawing_textbox_host_is_reported_and_preserved_out_of_scope(tmp_path: Path):
    extract_dir = _write_document_xml(
        tmp_path,
        (
            '<w:p><w:r><w:t>Host</w:t><w:drawing><w:txbxContent>'
            '<w:p><w:r><w:t>A. Nested text box</w:t></w:r></w:p>'
            '</w:txbxContent></w:drawing></w:r></w:p>'
            '<w:p><w:r><w:t>Ordinary requirement</w:t></w:r></w:p>'
        ),
    )

    bundle = build_phase2_slim_bundle(extract_dir, available_roles=["PARAGRAPH"])
    assert [p["paragraph_index"] for p in bundle["paragraphs"]] == [0, 1]
    assert bundle["paragraphs"][0]["text"] == "Host"
    assert bundle["filter_report"]["paragraphs_out_of_scope"] == [
        {
            "paragraph_index": 0,
            "reason": "drawing_or_textbox_subtree",
            "original_text_preview": "Host",
        }
    ]


def test_drawing_markup_inside_comment_does_not_hide_visible_paragraph(tmp_path: Path):
    extract_dir = _write_document_xml(
        tmp_path,
        '<w:p><!-- <w:drawing/> --><w:r><w:t>Ordinary requirement</w:t></w:r></w:p>',
    )

    bundle = build_phase2_slim_bundle(extract_dir, available_roles=["PARAGRAPH"])

    assert [p["paragraph_index"] for p in bundle["paragraphs"]] == [0]
    assert bundle["filter_report"]["paragraphs_out_of_scope"] == []


def test_visible_sectpr_paragraph_is_classifiable_but_empty_break_is_skipped(tmp_path: Path):
    extract_dir = _write_document_xml(
        tmp_path,
        (
            '<w:p><w:pPr><w:sectPr><w:pgSz/></w:sectPr></w:pPr>'
            '<w:r><w:t>Visible requirement</w:t></w:r></w:p>'
            '<w:p><w:pPr><w:sectPr/></w:pPr></w:p>'
        ),
    )

    bundle = build_phase2_slim_bundle(extract_dir, available_roles=["PARAGRAPH"])
    assert len(bundle["paragraphs"]) == 1
    assert bundle["paragraphs"][0]["paragraph_index"] == 0
    assert bundle["paragraphs"][0]["contains_sectPr"] is True
    assert bundle["filter_report"]["paragraphs_removed_entirely"] == [
        {
            "paragraph_index": 1,
            "tags": ["empty_structural_section_break"],
            "original_text_preview": "",
        }
    ]


def _write_numbering_parts(extract_dir: Path, styles_xml: str, numbering_xml: str) -> None:
    (extract_dir / "word" / "styles.xml").write_text(styles_xml, encoding="utf-8")
    (extract_dir / "word" / "numbering.xml").write_text(numbering_xml, encoding="utf-8")


def test_direct_numbering_override_uniquely_matches_role_signature(tmp_path: Path):
    extract_dir = _write_document_xml(
        tmp_path,
        '<w:p><w:pPr><w:numPr><w:ilvl w:val="1"/><w:numId w:val="7"/>'
        '</w:numPr></w:pPr><w:r><w:t>Requirement</w:t></w:r></w:p>',
    )
    _write_numbering_parts(
        extract_dir,
        '<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"/>',
        '<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
        '<w:abstractNum w:abstractNumId="3"><w:lvl w:ilvl="1">'
        '<w:numFmt w:val="decimal"/><w:lvlText w:val="%2."/></w:lvl></w:abstractNum>'
        '<w:num w:numId="7"><w:abstractNumId w:val="3"/><w:lvlOverride w:ilvl="1">'
        '<w:startOverride w:val="5"/><w:lvl w:ilvl="1"><w:numFmt w:val="upperLetter"/>'
        '<w:lvlText w:val="%2)"/></w:lvl></w:lvlOverride></w:num></w:numbering>',
    )
    role_specs = {
        "PARAGRAPH": {
            "numbering_provenance": "style_numpr",
            "numbering_pattern": {
                "numId": "99",
                "abstractNumId": "44",
                "ilvl": "1",
                "numFmt": "upperLetter",
                "lvlText": "%2)",
                "startOverride": "5",
            },
        },
        "SUBPARAGRAPH": {
            "numbering_provenance": "style_numpr",
            "numbering_pattern": {
                "numId": "100", "ilvl": "1", "numFmt": "decimal", "lvlText": "%2."
            },
        },
    }

    bundle = build_phase2_slim_bundle(
        extract_dir,
        available_roles=["PARAGRAPH", "SUBPARAGRAPH"],
        role_specs=role_specs,
    )
    assert bundle["paragraphs"] == []
    assert bundle["deterministic_classifications"] == [
        {"paragraph_index": 0, "csi_role": "PARAGRAPH"}
    ]


def test_typical_automatic_csi_signature_is_recognized_against_canadian_template(
    tmp_path: Path,
):
    extract_dir = _write_document_xml(
        tmp_path,
        '<w:p><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="7"/>'
        '</w:numPr></w:pPr><w:r><w:t>Scope</w:t></w:r></w:p>',
    )
    _write_numbering_parts(
        extract_dir,
        '<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"/>',
        '<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
        '<w:abstractNum w:abstractNumId="3"><w:lvl w:ilvl="0">'
        '<w:numFmt w:val="upperLetter"/><w:lvlText w:val="%1."/>'
        '</w:lvl></w:abstractNum><w:num w:numId="7">'
        '<w:abstractNumId w:val="3"/></w:num></w:numbering>',
    )
    canadian_role = {
        "numbering_provenance": "style_numpr",
        "numbering_pattern": {
            "numId": "99",
            "ilvl": "0",
            "numFmt": "decimal",
            "lvlText": ".%1",
        },
    }

    bundle = build_phase2_slim_bundle(
        extract_dir,
        available_roles=["PARAGRAPH"],
        role_specs={"PARAGRAPH": canadian_role},
    )

    assert bundle["paragraphs"] == []
    assert bundle["deterministic_classifications"] == [
        {"paragraph_index": 0, "csi_role": "PARAGRAPH"}
    ]


def test_style_inherited_numbering_is_resolved_and_matched(tmp_path: Path):
    extract_dir = _write_document_xml(
        tmp_path,
        '<w:p><w:pPr><w:pStyle w:val="AutoChild"/></w:pPr>'
        '<w:r><w:t>Requirement</w:t></w:r></w:p>',
    )
    _write_numbering_parts(
        extract_dir,
        '<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
        '<w:style w:type="paragraph" w:styleId="AutoBase"><w:pPr><w:numPr>'
        '<w:ilvl w:val="2"/><w:numId w:val="8"/></w:numPr></w:pPr></w:style>'
        '<w:style w:type="paragraph" w:styleId="AutoChild"><w:basedOn w:val="AutoBase"/>'
        '</w:style></w:styles>',
        '<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
        '<w:abstractNum w:abstractNumId="4"><w:lvl w:ilvl="2">'
        '<w:numFmt w:val="lowerLetter"/><w:lvlText w:val="%3."/></w:lvl></w:abstractNum>'
        '<w:num w:numId="8"><w:abstractNumId w:val="4"/></w:num></w:numbering>',
    )
    role_specs = {
        "SUBSUBPARAGRAPH": {
            "numbering_provenance": "direct_numpr",
            "numbering_pattern": {
                "numId": "17", "ilvl": "2", "numFmt": "lowerLetter", "lvlText": "%3."
            },
        }
    }

    bundle = build_phase2_slim_bundle(
        extract_dir,
        available_roles=["SUBSUBPARAGRAPH"],
        role_specs=role_specs,
    )
    assert bundle["paragraphs"] == []
    assert bundle["deterministic_classifications"] == [
        {"paragraph_index": 0, "csi_role": "SUBSUBPARAGRAPH"}
    ]


def test_ambiguous_automatic_numbering_stays_unresolved_for_llm(tmp_path: Path):
    extract_dir = _write_document_xml(
        tmp_path,
        '<w:p><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="7"/>'
        '</w:numPr></w:pPr><w:r><w:t>Requirement</w:t></w:r></w:p>',
    )
    _write_numbering_parts(
        extract_dir,
        '<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"/>',
        '<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
        '<w:abstractNum w:abstractNumId="3"><w:lvl w:ilvl="0">'
        '<w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/></w:lvl></w:abstractNum>'
        '<w:num w:numId="7"><w:abstractNumId w:val="3"/></w:num></w:numbering>',
    )
    same_pattern = {
        "numbering_provenance": "style_numpr",
        "numbering_pattern": {
            "numId": "99", "ilvl": "0", "numFmt": "decimal", "lvlText": "%1."
        },
    }

    bundle = build_phase2_slim_bundle(
        extract_dir,
        available_roles=["PARAGRAPH", "SUBPARAGRAPH"],
        role_specs={"PARAGRAPH": same_pattern, "SUBPARAGRAPH": same_pattern},
    )
    assert bundle["deterministic_classifications"] == []
    assert len(bundle["paragraphs"]) == 1
    assert bundle["paragraphs"][0]["marker_type"] == "automatic_numbering"
    assert bundle["paragraphs"][0]["numbering_match_candidates"] == [
        "PARAGRAPH", "SUBPARAGRAPH"
    ]


def test_extra_target_start_override_prevents_deterministic_match(tmp_path: Path):
    extract_dir = _write_document_xml(
        tmp_path,
        '<w:p><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="7"/>'
        '</w:numPr></w:pPr><w:r><w:t>Requirement</w:t></w:r></w:p>',
    )
    _write_numbering_parts(
        extract_dir,
        '<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"/>',
        '<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
        '<w:abstractNum w:abstractNumId="3"><w:lvl w:ilvl="0">'
        '<w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/></w:lvl></w:abstractNum>'
        '<w:num w:numId="7"><w:abstractNumId w:val="3"/>'
        '<w:lvlOverride w:ilvl="0"><w:startOverride w:val="4"/></w:lvlOverride>'
        '</w:num></w:numbering>',
    )
    bundle = build_phase2_slim_bundle(
        extract_dir,
        available_roles=["SUBPARAGRAPH"],
        role_specs={
            "SUBPARAGRAPH": {
                "numbering_provenance": "style_numpr",
                "numbering_pattern": {
                    "numId": "9", "ilvl": "0", "numFmt": "decimal", "lvlText": "%1."
                },
            }
        },
    )
    assert bundle["deterministic_classifications"] == []
    assert bundle["paragraphs"][0]["numbering_match_candidates"] == []


def test_deep_typed_csi_markers_are_detected_and_preclassified(tmp_path: Path):
    extract_dir = _write_document_xml(
        tmp_path,
        '<w:p><w:r><w:t>1) Fifth level</w:t></w:r></w:p>'
        '<w:p><w:r><w:t>a) Sixth level</w:t></w:r></w:p>'
        '<w:p><w:r><w:t>(1) Seventh level</w:t></w:r></w:p>'
        '<w:p><w:r><w:t>(a) Eighth level</w:t></w:r></w:p>',
    )
    roles = [f"SUBPARAGRAPH_LEVEL_{level}" for level in range(5, 9)]

    bundle = build_phase2_slim_bundle(extract_dir, available_roles=roles)

    assert bundle["paragraphs"] == []
    assert bundle["deterministic_classifications"] == [
        {"paragraph_index": index, "csi_role": role}
        for index, role in enumerate(roles)
    ]


def test_preclassify_markers():
    paragraphs = [
        {"paragraph_index": 0, "text": "PART 1 GENERAL", "in_table": False, "marker_type": None},
        {"paragraph_index": 1, "text": "1.01 SUMMARY", "in_table": False, "marker_type": None},
        {"paragraph_index": 2, "text": "A. Scope", "in_table": False, "marker_type": "upper_alpha"},
        {"paragraph_index": 3, "text": "1. Sub item", "in_table": False, "marker_type": "number"},
        {"paragraph_index": 4, "text": "a. Lower", "in_table": False, "marker_type": "lower_alpha"},
        {"paragraph_index": 5, "text": "1) Deeper", "in_table": False, "marker_type": "deep_level_5"},
        {"paragraph_index": 6, "text": "a) Deeper", "in_table": False, "marker_type": "deep_level_6"},
        {"paragraph_index": 7, "text": "(1) Deeper", "in_table": False, "marker_type": "deep_level_7"},
        {"paragraph_index": 8, "text": "(a) Deeper", "in_table": False, "marker_type": "deep_level_8"},
    ]
    roles = [
        "PART",
        "ARTICLE",
        "PARAGRAPH",
        "SUBPARAGRAPH",
        "SUBSUBPARAGRAPH",
        "SUBPARAGRAPH_LEVEL_5",
        "SUBPARAGRAPH_LEVEL_6",
        "SUBPARAGRAPH_LEVEL_7",
        "SUBPARAGRAPH_LEVEL_8",
    ]
    out = preclassify_paragraphs(paragraphs, roles)
    assert out[0] == "PART"
    assert out[1] == "ARTICLE"
    assert out[2] == "PARAGRAPH"
    assert out[3] == "SUBPARAGRAPH"
    assert out[4] == "SUBSUBPARAGRAPH"
    assert out[5] == "SUBPARAGRAPH_LEVEL_5"
    assert out[6] == "SUBPARAGRAPH_LEVEL_6"
    assert out[7] == "SUBPARAGRAPH_LEVEL_7"
    assert out[8] == "SUBPARAGRAPH_LEVEL_8"


def test_end_of_section_is_deterministic_not_boilerplate_removed(tmp_path: Path):
    extract_dir = _write_document_xml(
        tmp_path,
        '<w:p><w:r><w:t>END OF SECTION 23 05 13</w:t></w:r></w:p>',
    )

    bundle = build_phase2_slim_bundle(extract_dir, available_roles=["END_OF_SECTION", "PART", "ARTICLE"])
    assert bundle["filter_report"]["paragraphs_removed_entirely"] == []
    assert bundle["deterministic_classifications"] == [
        {"paragraph_index": 0, "csi_role": "END_OF_SECTION"}
    ]
    assert bundle["paragraphs"] == []


def test_validation_fails_on_duplicate_and_missing_coverage():
    bundle = {
        "paragraphs": [{"paragraph_index": 10}, {"paragraph_index": 11}],
        "deterministic_classifications": [],
    }
    with pytest.raises(ValueError, match="duplicate"):
        validate_phase2_classification_contract(
            bundle,
            {"classifications": [
                {"paragraph_index": 10, "csi_role": "PART"},
                {"paragraph_index": 10, "csi_role": "ARTICLE"},
            ]},
            ["PART", "ARTICLE"],
        )

    with pytest.raises(ValueError, match="missing coverage"):
        validate_phase2_classification_contract(
            bundle,
            {"classifications": [{"paragraph_index": 10, "csi_role": "PART"}]},
            ["PART", "ARTICLE"],
        )


def test_coerce_unresolved_only_merges_deterministic():
    bundle = {
        "paragraphs": [{"paragraph_index": 11}],
        "deterministic_classifications": [{"paragraph_index": 10, "csi_role": "PART"}],
    }
    out = coerce_to_final_classifications(
        bundle,
        {"classifications": [{"paragraph_index": 11, "csi_role": "ARTICLE"}]},
        ["PART", "ARTICLE"],
    )
    assert out["classifications"] == [
        {"paragraph_index": 10, "csi_role": "PART"},
        {"paragraph_index": 11, "csi_role": "ARTICLE"},
    ]


def test_coerce_rejects_deterministic_override():
    bundle = {
        "paragraphs": [{"paragraph_index": 11}],
        "deterministic_classifications": [{"paragraph_index": 10, "csi_role": "PART"}],
    }
    with pytest.raises(ValueError, match="deterministic override"):
        validate_phase2_final_payload(
            bundle,
            {
                "classifications": [
                    {"paragraph_index": 10, "csi_role": "ARTICLE"},
                    {"paragraph_index": 11, "csi_role": "ARTICLE"},
                ]
            },
            ["PART", "ARTICLE"],
        )
