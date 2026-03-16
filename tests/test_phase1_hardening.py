from __future__ import annotations

from pathlib import Path

import pytest

from docx_decomposer import (
    build_slim_bundle,
    build_style_registry_dict,
    extract_docx,
    validate_instructions,
)
from llm_classifier import compute_coverage, classify_document


def _bundle_for_semantics() -> dict:
    return {
        "paragraphs": [
            {"paragraph_index": 0, "text": "SECTION 01 11 00", "contains_sectPr": False, "in_table": False, "skip_reason": None},
            {"paragraph_index": 1, "text": "PROJECT TITLE", "contains_sectPr": False, "in_table": False, "skip_reason": None},
            {"paragraph_index": 2, "text": "PART 1 - GENERAL", "contains_sectPr": False, "in_table": False, "skip_reason": None},
            {"paragraph_index": 3, "text": "1.01 SUMMARY", "contains_sectPr": False, "in_table": False, "skip_reason": None},
            {"paragraph_index": 4, "text": "A. Scope", "contains_sectPr": False, "in_table": False, "skip_reason": None},
            {"paragraph_index": 5, "text": "1. Requirements", "contains_sectPr": False, "in_table": False, "skip_reason": None},
            {"paragraph_index": 6, "text": "a. Detail", "contains_sectPr": False, "in_table": False, "skip_reason": None},
        ],
        "style_catalog": {"Normal": {}},
    }


def test_semantic_validator_rejects_single_style_collapse():
    bundle = _bundle_for_semantics()
    instructions = {
        "apply_pStyle": [
            {"paragraph_index": i, "styleId": "CSI_Paragraph__ARCH"} for i in range(7)
        ],
        "roles": {
            "SectionID": {"styleId": "CSI_Paragraph__ARCH", "exemplar_paragraph_index": 0},
            "SectionTitle": {"styleId": "CSI_Paragraph__ARCH", "exemplar_paragraph_index": 1},
            "PART": {"styleId": "CSI_Paragraph__ARCH", "exemplar_paragraph_index": 2},
            "ARTICLE": {"styleId": "CSI_Paragraph__ARCH", "exemplar_paragraph_index": 3},
            "PARAGRAPH": {"styleId": "CSI_Paragraph__ARCH", "exemplar_paragraph_index": 4},
            "SUBPARAGRAPH": {"styleId": "CSI_Paragraph__ARCH", "exemplar_paragraph_index": 5},
            "SUBSUBPARAGRAPH": {"styleId": "CSI_Paragraph__ARCH", "exemplar_paragraph_index": 6},
        },
    }
    with pytest.raises(ValueError, match="Semantic validation failed"):
        validate_instructions(instructions, slim_bundle=bundle)


def test_long_bracketed_note_remains_skippable_after_truncation(tmp_path: Path):
    extract_dir = tmp_path / "x"
    (extract_dir / "word").mkdir(parents=True)
    long_note = "[" + ("X" * 260) + "]"
    (extract_dir / "word" / "document.xml").write_text(
        '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>'
        f'<w:p><w:r><w:t>{long_note}</w:t></w:r></w:p>'
        '<w:p><w:r><w:t>A. Scope</w:t></w:r></w:p>'
        '<w:p><w:sectPr/></w:p>'
        '</w:body></w:document>',
        encoding="utf-8",
    )
    (extract_dir / "word" / "styles.xml").write_text(
        '<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"></w:styles>',
        encoding="utf-8",
    )

    bundle = build_slim_bundle(extract_dir)
    assert bundle["paragraphs"][0]["text_was_truncated"] is True
    assert bundle["paragraphs"][0]["skip_reason"] == "editor_note"

    instructions = {
        "apply_pStyle": [{"paragraph_index": 1, "styleId": "CSI_Paragraph__ARCH"}],
        "roles": {"PARAGRAPH": {"styleId": "CSI_Paragraph__ARCH", "exemplar_paragraph_index": 1}},
    }
    cov, styled, classifiable = compute_coverage(bundle, instructions)
    assert (cov, styled, classifiable) == (1.0, 1, 1)


def test_article_marker_1_01_reports_text_literal(tmp_path: Path):
    extract_dir = tmp_path / "x"
    (extract_dir / "word").mkdir(parents=True)
    (extract_dir / "word" / "styles.xml").write_text(
        '<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
        '<w:style w:type="paragraph" w:styleId="CSI_Article__ARCH"><w:name w:val="Article"/></w:style>'
        '</w:styles>',
        encoding="utf-8",
    )
    (extract_dir / "word" / "document.xml").write_text(
        '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>'
        '<w:p><w:r><w:t>1.01 SUMMARY</w:t></w:r></w:p><w:p><w:sectPr/></w:p>'
        '</w:body></w:document>',
        encoding="utf-8",
    )

    instructions = {"roles": {"ARTICLE": {"styleId": "CSI_Article__ARCH", "exemplar_paragraph_index": 0}}}
    reg = build_style_registry_dict(extract_dir, "test.docx", instructions)
    assert reg["roles"]["ARTICLE"]["numbering_provenance"] == "text_literal"


def test_style_registry_warning_uses_pre_apply_bundle(tmp_path: Path):
    extract_dir = tmp_path / "x"
    (extract_dir / "word").mkdir(parents=True)
    (extract_dir / "word" / "styles.xml").write_text(
        '<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
        '<w:style w:type="paragraph" w:styleId="CSI_Paragraph__ARCH"><w:name w:val="Paragraph"/></w:style>'
        '</w:styles>',
        encoding="utf-8",
    )
    (extract_dir / "word" / "document.xml").write_text(
        '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>'
        '<w:p><w:pPr><w:pStyle w:val="CSI_Paragraph__ARCH"/></w:pPr><w:r><w:t>A. Scope</w:t></w:r></w:p>'
        '<w:p><w:sectPr/></w:p>'
        '</w:body></w:document>',
        encoding="utf-8",
    )

    instructions = {
        "create_styles": [{"styleId": "CSI_Paragraph__ARCH"}],
        "roles": {"PARAGRAPH": {"styleId": "CSI_Paragraph__ARCH", "exemplar_paragraph_index": 0}},
    }
    pre_apply = {
        "paragraphs": [
            {"paragraph_index": 0, "text": "A. Scope", "pStyle": None, "skip_reason": None},
        ]
    }
    reg = build_style_registry_dict(extract_dir, "x.docx", instructions, pre_apply_bundle=pre_apply)
    assert "warning" not in reg["roles"]["PARAGRAPH"]


def test_extract_docx_raises_if_target_exists_without_overwrite(tmp_path: Path):
    docx_path = tmp_path / "a.docx"
    docx_path.write_bytes(b"PK\x05\x06" + b"\x00" * 18)
    extract_dir = tmp_path / "existing"
    extract_dir.mkdir()
    with pytest.raises(FileExistsError):
        extract_docx(docx_path, extract_dir)


def test_large_doc_raises_clear_not_supported_error():
    bundle = {"paragraphs": [{"paragraph_index": i, "text": "x", "skip_reason": None} for i in range(501)]}
    with pytest.raises(ValueError, match="chunked mode requires redesign"):
        classify_document(bundle, "x", "y", api_key="k")
