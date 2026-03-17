from __future__ import annotations

from pathlib import Path

import pytest

from docx_decomposer import build_slim_bundle
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


def test_501_paragraph_document_fails_with_explicit_message():
    bundle = {"paragraphs": [{"paragraph_index": i, "text": "x"} for i in range(501)]}
    with pytest.raises(ValueError, match="Unsupported document size"):
        classify_document(bundle, "m", "r", "fake-key")
