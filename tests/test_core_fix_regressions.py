from __future__ import annotations

from pathlib import Path
import xml.etree.ElementTree as ET

import pytest

from docx_decomposer import (
    build_portable_styles_xml,
    build_slim_bundle,
    build_style_xml_block,
    derive_style_def_from_paragraph,
    extract_paragraph_rpr_inner,
    paragraph_text_from_block,
)
from paragraph_rules import detect_role_signal, infer_expected_roles


W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
W14_NS = "http://schemas.microsoft.com/office/word/2010/wordml"


def test_visible_text_keeps_tabs_and_explicit_breaks_as_separators() -> None:
    paragraph = (
        "<w:p>"
        "<w:r><w:t>SECTION</w:t><w:tab/><w:t>21</w:t></w:r>"
        "<w:r><w:br/><w:t>13</w:t><w:cr/><w:t>13</w:t></w:r>"
        "</w:p>"
    )

    assert paragraph_text_from_block(paragraph) == "SECTION 21 13 13"


def test_one_bold_run_does_not_make_the_derived_style_bold() -> None:
    paragraph = (
        "<w:p>"
        "<w:r><w:rPr><w:b/></w:rPr><w:t>A.</w:t></w:r>"
        "<w:r><w:t> This long requirement is ordinary body text.</w:t></w:r>"
        "</w:p>"
    )

    derived = derive_style_def_from_paragraph(
        "CSI_Paragraph__ARCH",
        "Paragraph",
        paragraph,
        based_on="Normal",
    )
    style_xml = build_style_xml_block(derived)

    assert derived["rPr_inner"] == ""
    assert "<w:b/>" not in style_xml
    assert "<w:b>" not in style_xml
    assert "<w:b " not in style_xml


def test_different_extension_run_properties_are_not_promoted() -> None:
    paragraph = (
        "<w:p>"
        "<w:r><w:rPr><w:b/><w14:textFill><w14:solidFill>"
        '<w14:srgbClr w14:val="FF0000"/>'
        "</w14:solidFill></w14:textFill></w:rPr><w:t>First</w:t></w:r>"
        "<w:r><w:rPr><w:b/><w14:textFill><w14:solidFill>"
        '<w14:srgbClr w14:val="0000FF"/>'
        "</w14:solidFill></w14:textFill></w:rPr><w:t>Second</w:t></w:r>"
        "</w:p>"
    )

    rpr_inner = extract_paragraph_rpr_inner(paragraph)

    assert "<w:b/>" in rpr_inner
    assert "w14:textFill" not in rpr_inner
    assert "w14:srgbClr" not in rpr_inner


def test_portable_styles_preserve_extension_properties_and_declare_namespace(
    tmp_path: Path,
) -> None:
    word_dir = tmp_path / "word"
    word_dir.mkdir()
    effect = (
        '<w14:glow w14:rad="63500"><w14:srgbClr w14:val="ABCDEF"/>'
        "</w14:glow>"
        "<w14:textFill><w14:solidFill>"
        '<w14:srgbClr w14:val="123456"/>'
        "</w14:solidFill></w14:textFill>"
    )
    (word_dir / "document.xml").write_text(
        f'<w:document xmlns:w="{W_NS}" xmlns:w14="{W14_NS}"><w:body>'
        f"<w:p><w:r><w:rPr>{effect}</w:rPr><w:t>PROJECT </w:t></w:r>"
        f"<w:r><w:rPr>{effect}</w:rPr><w:t>TITLE</w:t></w:r></w:p>"
        "<w:sectPr/>"
        "</w:body></w:document>",
        encoding="utf-8",
    )
    (word_dir / "styles.xml").write_text(
        f'<w:styles xmlns:w="{W_NS}">'
        '<w:style w:type="paragraph" w:default="1" w:styleId="Normal">'
        '<w:name w:val="Normal"/></w:style>'
        "</w:styles>",
        encoding="utf-8",
    )
    instructions = {
        "create_styles": [
            {
                "styleId": "CSI_SectionTitle__ARCH",
                "name": "Section Title",
                "type": "paragraph",
                "derive_from_paragraph_index": 0,
            }
        ],
        "apply_pStyle": [
            {"paragraph_index": 0, "styleId": "CSI_SectionTitle__ARCH"}
        ],
        "ignored_paragraphs": [],
        "roles": {
            "SectionTitle": {
                "styleId": "CSI_SectionTitle__ARCH",
                "exemplar_paragraph_index": 0,
            }
        },
        "notes": [],
    }

    portable = build_portable_styles_xml(tmp_path, instructions)

    assert f'xmlns:w14="{W14_NS}"' in portable
    assert "<w14:glow" in portable
    assert "<w14:textFill>" in portable
    root = ET.fromstring(portable)
    style = root.find(f"{{{W_NS}}}style[@{{{W_NS}}}styleId='CSI_SectionTitle__ARCH']")
    assert style is not None
    rpr = style.find(f"{{{W_NS}}}rPr")
    assert rpr is not None
    assert rpr.find(f"{{{W14_NS}}}glow") is not None
    assert rpr.find(f"{{{W14_NS}}}textFill") is not None


def test_derived_style_preserves_source_pstyle_and_direct_numpr() -> None:
    paragraph = (
        "<w:p><w:pPr>"
        '<w:pStyle w:val="ArchitectNumbered"/>'
        '<w:numPr><w:ilvl w:val="2"/><w:numId w:val="17"/></w:numPr>'
        '<w:spacing w:after="120"/>'
        "</w:pPr><w:r><w:t>Requirement</w:t></w:r></w:p>"
    )

    derived = derive_style_def_from_paragraph(
        "CSI_Subparagraph__ARCH",
        "Subparagraph",
        paragraph,
        based_on="Normal",
    )
    style_xml = build_style_xml_block(derived)

    assert derived["basedOn"] == "ArchitectNumbered"
    assert "<w:pStyle" not in derived["pPr_inner"]
    assert '<w:numId w:val="17"/>' in derived["pPr_inner"]
    assert '<w:ilvl w:val="2"/>' in derived["pPr_inner"]
    assert '<w:basedOn w:val="ArchitectNumbered"/>' in style_xml
    assert "<w:numPr>" in style_xml


def test_style_inherited_numbering_enters_slim_numbering_catalog(tmp_path: Path) -> None:
    word_dir = tmp_path / "word"
    word_dir.mkdir()
    (word_dir / "document.xml").write_text(
        f'<w:document xmlns:w="{W_NS}"><w:body>'
        '<w:p><w:pPr><w:pStyle w:val="AutoNumbered"/></w:pPr>'
        '<w:r><w:t>Requirement</w:t></w:r></w:p>'
        "<w:sectPr/>"
        "</w:body></w:document>",
        encoding="utf-8",
    )
    (word_dir / "styles.xml").write_text(
        f'<w:styles xmlns:w="{W_NS}">'
        '<w:style w:type="paragraph" w:styleId="NumberedBase">'
        '<w:name w:val="Numbered Base"/><w:pPr><w:numPr>'
        '<w:ilvl w:val="1"/><w:numId w:val="7"/>'
        "</w:numPr></w:pPr></w:style>"
        '<w:style w:type="paragraph" w:styleId="AutoNumbered">'
        '<w:name w:val="Auto Numbered"/><w:basedOn w:val="NumberedBase"/>'
        "</w:style>"
        "</w:styles>",
        encoding="utf-8",
    )
    (word_dir / "numbering.xml").write_text(
        f'<w:numbering xmlns:w="{W_NS}">'
        '<w:abstractNum w:abstractNumId="3">'
        '<w:lvl w:ilvl="1"><w:numFmt w:val="decimal"/>'
        '<w:lvlText w:val="%2."/></w:lvl>'
        "</w:abstractNum>"
        '<w:num w:numId="7"><w:abstractNumId w:val="3"/></w:num>'
        "</w:numbering>",
        encoding="utf-8",
    )

    bundle = build_slim_bundle(tmp_path)

    paragraph = bundle["paragraphs"][0]
    assert paragraph["numPr"] is None
    assert paragraph["effective_numPr"] == {"numId": "7", "ilvl": "1"}
    assert bundle["style_catalog"]["AutoNumbered"]["resolved_numPr"] == {
        "numId": "7",
        "ilvl": "1",
    }
    assert bundle["numbering_catalog"]["nums"]["7"]["abstractNumId"] == "3"
    assert bundle["numbering_catalog"]["abstracts"]["3"]["levels"] == [
        {"ilvl": "1", "numFmt": "decimal", "lvlText": "%2."}
    ]


@pytest.mark.parametrize(
    "text",
    [
        "SECTION 211313 - WET-PIPE SPRINKLER SYSTEMS",
        "SECTION21-13-13: WET-PIPE SPRINKLER SYSTEMS",
        "SECTION\u00a021\u00a013\u00a013 — WET-PIPE SPRINKLER SYSTEMS",
    ],
)
def test_compact_and_combined_section_lines_supply_id_and_title_roles(text: str) -> None:
    assert detect_role_signal(text, numeric_is_strong=False, lower_is_strong=False) == "SectionID"

    expected, hits = infer_expected_roles(
        [{"paragraph_index": 0, "text": text, "skip_reason": None}]
    )

    assert {"SectionID", "SectionTitle"} <= expected
    assert hits["SectionID"] == [0]
    assert hits["SectionTitle"] == [0]


@pytest.mark.parametrize(
    "text",
    [
        "END OF SECTION",
        "END OF SECTION 211313",
        "END\u00a0OF\u00a0DIVISION 21-13-13.",
    ],
)
def test_end_of_section_variants_are_structural_signals(text: str) -> None:
    assert (
        detect_role_signal(text, numeric_is_strong=False, lower_is_strong=False)
        == "END_OF_SECTION"
    )

    expected, hits = infer_expected_roles(
        [{"paragraph_index": 4, "text": text, "skip_reason": None}]
    )
    assert "END_OF_SECTION" in expected
    assert hits["END_OF_SECTION"] == [4]
