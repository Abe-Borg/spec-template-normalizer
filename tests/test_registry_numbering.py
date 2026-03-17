from __future__ import annotations

from pathlib import Path

from docx_decomposer import build_style_registry_dict


def test_direct_numpr_provenance(tmp_path: Path):
    extract_dir = tmp_path / "x"
    (extract_dir / "word").mkdir(parents=True)
    (extract_dir / "word" / "styles.xml").write_text(
        '<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
        '<w:style w:type="paragraph" w:styleId="Normal"><w:name w:val="Normal"/></w:style>'
        '</w:styles>', encoding="utf-8"
    )
    (extract_dir / "word" / "numbering.xml").write_text(
        '<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
        '<w:abstractNum w:abstractNumId="10"><w:lvl w:ilvl="0"><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/></w:lvl></w:abstractNum>'
        '<w:num w:numId="9"><w:abstractNumId w:val="10"/></w:num>'
        '</w:numbering>', encoding="utf-8"
    )
    (extract_dir / "word" / "document.xml").write_text(
        '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>'
        '<w:p><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="9"/></w:numPr></w:pPr><w:r><w:t>General</w:t></w:r></w:p>'
        '<w:p><w:sectPr/></w:p></w:body></w:document>', encoding="utf-8"
    )

    reg = build_style_registry_dict(
        extract_dir,
        "x.docx",
        {"roles": {"ARTICLE": {"styleId": "Normal", "exemplar_paragraph_index": 0}}},
    )
    assert reg["version"] == 2
    assert reg["roles"]["ARTICLE"]["numbering_provenance"] == "direct_numpr"
    assert reg["roles"]["ARTICLE"]["numbering_pattern"]["numFmt"] == "decimal"
