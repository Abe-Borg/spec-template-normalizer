from __future__ import annotations

from pathlib import Path

import pytest

from docx_decomposer import (
    apply_pstyle_to_paragraph_block,
    build_portable_styles_xml,
    build_slim_bundle,
    extract_document_sectpr_blocks,
    extract_paragraph_ppr_inner,
    extract_paragraph_rpr_inner,
    extract_sectpr_block,
    iter_paragraph_xml_blocks,
    paragraph_contains_sectpr,
    paragraph_numpr_from_block,
    paragraph_ppr_hints_from_block,
    paragraph_pstyle_from_block,
    paragraph_rpr_hints_from_block,
    paragraph_text_from_block,
    ppr_without_pstyle,
    snapshot_headers_footers,
    strip_pstyle_from_paragraph,
)
from llm_classifier import classify_document


def test_nested_table_context_is_xml_aware(tmp_path: Path):
    extract_dir = tmp_path / "x"
    (extract_dir / "word").mkdir(parents=True)
    (extract_dir / "word" / "styles.xml").write_text(
        '<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"></w:styles>',
        encoding="utf-8",
    )
    (extract_dir / "word" / "document.xml").write_text(
        '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>'
        '<w:tbl><w:tr><w:tc>'
        '<w:p><w:r><w:t>outer in table</w:t></w:r></w:p>'
        '<w:tbl><w:tr><w:tc><w:p><w:r><w:t>inner table p</w:t></w:r></w:p></w:tc></w:tr></w:tbl>'
        '<w:p><w:r><w:t>still in outer table</w:t></w:r></w:p>'
        '</w:tc></w:tr></w:tbl>'
        '<w:p><w:r><w:t>outside table</w:t></w:r></w:p>'
        '<w:p><w:sectPr/></w:p>'
        '</w:body></w:document>',
        encoding="utf-8",
    )

    bundle = build_slim_bundle(extract_dir)
    assert bundle["paragraphs"][0]["in_table"] is True
    assert bundle["paragraphs"][1]["in_table"] is True
    assert bundle["paragraphs"][2]["in_table"] is True
    assert bundle["paragraphs"][3]["in_table"] is False


def test_text_box_paragraphs_keep_outer_blocks_aligned(tmp_path: Path):
    extract_dir = tmp_path / "x"
    (extract_dir / "word").mkdir(parents=True)
    (extract_dir / "word" / "styles.xml").write_text(
        '<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"></w:styles>',
        encoding="utf-8",
    )
    text_box = (
        '<w:p><w:r><w:drawing><w:txbxContent>'
        '<w:p><w:pPr><w:sectPr/></w:pPr><w:r><w:t>Text box text</w:t></w:r></w:p>'
        '</w:txbxContent></w:drawing></w:r></w:p>'
    )
    document_xml = (
        '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>'
        '<w:p><w:r><w:t>Body text</w:t></w:r></w:p>'
        '<w:tbl><w:tr><w:tc>'
        '<w:p><w:r><w:t>Outer cell</w:t></w:r></w:p>'
        '<w:tbl><w:tr><w:tc><w:p><w:r><w:t>Inner cell</w:t></w:r></w:p>'
        '</w:tc></w:tr></w:tbl>'
        '</w:tc></w:tr></w:tbl>'
        f'{text_box}<w:sectPr/>'
        '</w:body></w:document>'
    )
    (extract_dir / "word" / "document.xml").write_text(document_xml, encoding="utf-8")

    blocks = list(iter_paragraph_xml_blocks(document_xml))
    assert len(blocks) == 4
    assert blocks[-1][2] == text_box
    assert all(left[1] <= right[0] for left, right in zip(blocks, blocks[1:]))

    bundle = build_slim_bundle(extract_dir)

    assert [paragraph["text"] for paragraph in bundle["paragraphs"]] == [
        "Body text",
        "Outer cell",
        "Inner cell",
        "",
    ]
    assert [paragraph["skip_reason"] for paragraph in bundle["paragraphs"]] == [
        None,
        "in_table",
        "in_table",
        "text_box",
    ]


def test_text_box_properties_are_isolated_from_host_paragraph():
    text_box_subtree = (
        '<w:txbxContent data-sentinel="keep-byte-identical">'
        '<w:p><w:pPr><w:pStyle w:val="InnerStyle"/>'
        '<w:numPr><w:ilvl w:val="2"/><w:numId w:val="41"/></w:numPr>'
        '<w:jc w:val="center"/><w:ind w:left="720"/>'
        '<w:spacing w:before="120"/><w:sectPr/>'
        '</w:pPr><w:r><w:rPr><w:b/></w:rPr><w:t>Inner box text</w:t></w:r></w:p>'
        '</w:txbxContent>'
    )
    host = (
        '<w:p><w:r><w:t>Visible host text</w:t></w:r>'
        f'<w:r><w:drawing>{text_box_subtree}</w:drawing></w:r></w:p>'
    )

    assert paragraph_text_from_block(host) == "Visible host text"
    assert paragraph_pstyle_from_block(host) is None
    assert paragraph_numpr_from_block(host) == {"numId": None, "ilvl": None}
    assert paragraph_ppr_hints_from_block(host) == {}
    assert paragraph_rpr_hints_from_block(host) == {}
    assert extract_paragraph_ppr_inner(host) == ""
    assert extract_paragraph_rpr_inner(host) == ""
    assert paragraph_contains_sectpr(host) is False
    assert ppr_without_pstyle(host) == ""
    assert strip_pstyle_from_paragraph(host) == host

    styled = apply_pstyle_to_paragraph_block(host, "OuterStyle")

    assert paragraph_pstyle_from_block(styled) == "OuterStyle"
    assert styled.startswith(
        '<w:p><w:pPr><w:pStyle w:val="OuterStyle"/></w:pPr>'
        '<w:r><w:t>Visible host text</w:t></w:r>'
    )
    start = styled.index("<w:txbxContent")
    end = styled.index("</w:txbxContent>", start) + len("</w:txbxContent>")
    assert styled[start:end] == text_box_subtree
    assert strip_pstyle_from_paragraph(styled) == host


def test_tracked_previous_paragraph_properties_are_not_live_or_edited():
    change = (
        '<w:pPrChange w:id="7"><w:pPr>'
        '<w:pStyle w:val="HistoricalList"/><w:numPr>'
        '<w:ilvl w:val="3"/><w:numId w:val="91"/>'
        '</w:numPr></w:pPr></w:pPrChange>'
    )
    paragraph = (
        f'<w:p><w:pPr>{change}</w:pPr>'
        '<w:r><w:t>Visible text</w:t></w:r></w:p>'
    )

    assert paragraph_pstyle_from_block(paragraph) is None
    assert paragraph_numpr_from_block(paragraph) == {
        "numId": None,
        "ilvl": None,
    }
    assert "HistoricalList" not in extract_paragraph_ppr_inner(paragraph)
    assert change in ppr_without_pstyle(paragraph)
    assert strip_pstyle_from_paragraph(paragraph) == paragraph

    styled = apply_pstyle_to_paragraph_block(paragraph, "CurrentBody")

    assert change in styled
    assert styled.startswith(
        '<w:p><w:pPr><w:pStyle w:val="CurrentBody"/><w:pPrChange'
    )
    assert strip_pstyle_from_paragraph(styled) == paragraph


def test_token_oversized_document_fails_before_api_call():
    bundle = {"paragraphs": [{"paragraph_index": 0, "text": "x" * 700_000}]}
    with pytest.raises(ValueError, match="safe single-pass limit"):
        classify_document(bundle, "m", "r", "fake-key")


def test_textbox_paragraphs_are_non_overlapping_and_reported_out_of_scope(tmp_path: Path):
    extract_dir = tmp_path / "x"
    word = extract_dir / "word"
    word.mkdir(parents=True)
    (word / "styles.xml").write_text(
        '<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"/>',
        encoding="utf-8",
    )
    host = (
        '<w:p><w:r><w:t>Host text</w:t></w:r><w:r><w:drawing>'
        '<wps:txbx xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">'
        '<w:txbxContent><w:p><w:pPr><w:pStyle w:val="NestedOnly"/>'
        '<w:numPr><w:numId w:val="42"/></w:numPr></w:pPr>'
        '<w:r><w:rPr><w:b/></w:rPr><w:t>Nested text box text</w:t></w:r></w:p>'
        '</w:txbxContent></wps:txbx></w:drawing></w:r></w:p>'
    )
    document = (
        '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>'
        + host
        + '<w:p><w:r><w:t>Outside</w:t></w:r></w:p>'
        '</w:body></w:document>'
    )
    (word / "document.xml").write_text(document, encoding="utf-8")

    blocks = list(iter_paragraph_xml_blocks(document))
    bundle = build_slim_bundle(extract_dir)

    assert [block[2] for block in blocks] == [host, '<w:p><w:r><w:t>Outside</w:t></w:r></w:p>']
    assert len(bundle["paragraphs"]) == 2
    assert bundle["paragraphs"][0]["text"] == "Host text"
    assert bundle["paragraphs"][0]["skip_reason"] is None
    assert bundle["paragraphs"][0]["pStyle"] is None
    assert bundle["paragraphs"][0]["numPr"] is None
    assert bundle["paragraphs"][0]["rPr_hints"] is None
    assert bundle["paragraphs"][0]["contains_drawing"] is True
    assert bundle["paragraphs"][0]["contains_textbox"] is True
    assert bundle["paragraphs"][1]["paragraph_index"] == 1


def test_visible_section_break_paragraph_is_classifiable_and_sectpr_is_preserved(tmp_path: Path):
    extract_dir = tmp_path / "x"
    word = extract_dir / "word"
    word.mkdir(parents=True)
    (word / "styles.xml").write_text(
        '<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"/>',
        encoding="utf-8",
    )
    paragraph = (
        '<w:p><w:pPr><w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr></w:pPr>'
        '<w:r><w:t>PART 1 - GENERAL</w:t></w:r></w:p>'
    )
    (word / "document.xml").write_text(
        '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
        f'<w:body>{paragraph}</w:body></w:document>',
        encoding="utf-8",
    )

    bundle = build_slim_bundle(extract_dir)
    styled = apply_pstyle_to_paragraph_block(paragraph, "CSI_Part__ARCH")

    assert bundle["paragraphs"][0]["contains_sectPr"] is True
    assert bundle["paragraphs"][0]["skip_reason"] is None
    assert '<w:pStyle w:val="CSI_Part__ARCH"/>' in styled
    assert '<w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>' in styled


def test_visible_section_break_is_not_copied_into_derived_style(tmp_path: Path):
    extract_dir = tmp_path / "x"
    word = extract_dir / "word"
    word.mkdir(parents=True)
    (word / "styles.xml").write_text(
        '<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
        '<w:style w:type="paragraph" w:default="1" w:styleId="Normal">'
        '<w:name w:val="Normal"/></w:style></w:styles>',
        encoding="utf-8",
    )
    paragraph = (
        '<w:p><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr>'
        '<w:sectPr><w:headerReference w:type="default" r:id="rId7"/>'
        '<w:pgSz w:w="12240" w:h="15840"/></w:sectPr></w:pPr>'
        '<w:r><w:t>PART 1 - GENERAL</w:t></w:r></w:p>'
    )
    document = (
        '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
        'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
        f'<w:body>{paragraph}</w:body></w:document>'
    )
    (word / "document.xml").write_text(document, encoding="utf-8")
    instructions = {
        "create_styles": [{
            "styleId": "CSI_Part__ARCH",
            "name": "CSI Part",
            "type": "paragraph",
            "derive_from_paragraph_index": 0,
            "basedOn": "Normal",
            "role": "PART",
        }],
        "apply_pStyle": [{"paragraph_index": 0, "styleId": "CSI_Part__ARCH"}],
        "ignored_paragraphs": [],
        "roles": {
            "PART": {
                "styleId": "CSI_Part__ARCH",
                "exemplar_paragraph_index": 0,
            }
        },
        "notes": [],
    }

    portable = build_portable_styles_xml(extract_dir, instructions)

    assert '<w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr>' in portable
    assert "<w:sectPr" not in portable
    assert "rId7" not in portable
    assert (word / "document.xml").read_text(encoding="utf-8") == document


def test_visible_text_excludes_revisions_and_preserves_word_controls():
    paragraph = (
        '<w:p><w:r><w:t>Kept</w:t><w:tab/></w:r>'
        '<w:del><w:r><w:t>Deleted</w:t></w:r></w:del>'
        '<w:moveFrom><w:r><w:t>Moved away</w:t></w:r></w:moveFrom>'
        '<w:moveTo><w:r><w:t>Moved here</w:t><w:br/></w:r></w:moveTo>'
        '<w:r><w:t>A</w:t><w:noBreakHyphen/><w:t>B</w:t>'
        '<w:softHyphen/><w:t>C</w:t></w:r></w:p>'
    )

    assert paragraph_text_from_block(paragraph) == "Kept Moved here A‑BC"


def test_visible_text_supports_paired_empty_word_controls():
    paragraph = (
        '<w:p><w:r><w:t>A</w:t><w:tab></w:tab><w:t>B</w:t>'
        '<w:br> </w:br><w:t>C</w:t><w:cr></w:cr><w:t>D</w:t>'
        '<w:noBreakHyphen></w:noBreakHyphen><w:t>E</w:t>'
        '<w:softHyphen></w:softHyphen><w:t>F</w:t></w:r></w:p>'
    )

    assert paragraph_text_from_block(paragraph) == "A B C D‑EF"


def test_sectpr_stability_capture_includes_paired_and_self_closing_blocks():
    document = (
        '<w:document><w:body><w:p><w:pPr><w:sectPr/></w:pPr></w:p>'
        '<w:sectPr><w:pgSz w:w="12240"/></w:sectPr></w:body></w:document>'
    )

    assert extract_sectpr_block(document) == (
        '<w:sectPr/>\n<w:sectPr><w:pgSz w:w="12240"/></w:sectPr>'
    )


def test_document_section_capture_ignores_table_and_textbox_subtrees():
    document = (
        '<w:document><w:body>'
        '<w:p><w:pPr><w:sectPr><w:type w:val="nextPage"/></w:sectPr></w:pPr></w:p>'
        '<w:tbl><w:tr><w:tc><w:p><w:pPr><w:sectPr/></w:pPr></w:p></w:tc></w:tr></w:tbl>'
        '<w:p><w:r><w:drawing><w:txbxContent><w:p><w:pPr><w:sectPr/>'
        '</w:pPr></w:p></w:txbxContent></w:drawing></w:r></w:p>'
        '<w:sectPr/></w:body></w:document>'
    )

    assert extract_document_sectpr_blocks(document) == [
        '<w:sectPr><w:type w:val="nextPage"/></w:sectPr>',
        '<w:sectPr/>',
    ]


def test_header_footer_stability_capture_follows_custom_relationship_target(tmp_path: Path):
    (tmp_path / "word" / "_rels").mkdir(parents=True)
    (tmp_path / "word" / "layout" / "_rels").mkdir(parents=True)
    (tmp_path / "word" / "layout" / "footer-looking.xml").write_bytes(b"custom-header")
    (tmp_path / "word" / "layout" / "_rels" / "footer-looking.xml.rels").write_bytes(
        b"custom-header-rels"
    )
    (tmp_path / "word" / "_rels" / "document.xml.rels").write_text(
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rIdHeader" '
        'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" '
        'Target="layout/footer-looking.xml"/></Relationships>',
        encoding="utf-8",
    )

    hashes = snapshot_headers_footers(tmp_path)

    assert set(hashes) == {
        "word/layout/footer-looking.xml",
        "word/layout/_rels/footer-looking.xml.rels",
    }
