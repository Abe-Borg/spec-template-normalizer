from __future__ import annotations

from pathlib import Path

from docx_decomposer import build_slim_bundle


def test_unused_paragraph_style_is_exposed(tmp_path: Path):
    extract_dir = tmp_path / "x"
    (extract_dir / "word").mkdir(parents=True)
    (extract_dir / "word" / "document.xml").write_text(
        '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>'
        '<w:p><w:pPr><w:pStyle w:val="Normal"/></w:pPr><w:r><w:t>Hello</w:t></w:r></w:p>'
        '<w:p><w:sectPr/></w:p>'
        '</w:body></w:document>',
        encoding="utf-8",
    )
    (extract_dir / "word" / "styles.xml").write_text(
        '<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
        '<w:style w:type="paragraph" w:styleId="Normal"><w:name w:val="Normal"/></w:style>'
        '<w:style w:type="paragraph" w:styleId="Unused"><w:name w:val="Unused"/></w:style>'
        '</w:styles>',
        encoding="utf-8",
    )

    bundle = build_slim_bundle(extract_dir)
    assert "Unused" in bundle["style_catalog"]
    assert bundle["style_catalog"]["Unused"]["in_use"] is False
