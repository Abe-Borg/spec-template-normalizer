from pathlib import Path

import pytest

from spec_formatter.style_application.core.stability import (
    extract_sectpr_block,
    snapshot_stability,
    verify_stability,
)


PKG_REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
OFFICE_REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"


def test_stability_section_capture_handles_self_closing_and_ignores_nested_content():
    document = (
        '<w:document xmlns:w="urn:w"><w:body>'
        '<w:p><w:pPr><w:sectPr><w:type w:val="nextPage"/></w:sectPr></w:pPr></w:p>'
        '<w:tbl><w:tr><w:tc><w:p><w:pPr><w:sectPr/></w:pPr></w:p></w:tc></w:tr></w:tbl>'
        '<w:p><w:r><w:drawing><w:txbxContent><w:p><w:pPr><w:sectPr/>'
        '</w:pPr></w:p></w:txbxContent></w:drawing></w:r></w:p>'
        '<w:sectPr/></w:body></w:document>'
    )

    captured = extract_sectpr_block(document)

    assert captured.count("<w:sectPr") == 2
    assert '<w:type w:val="nextPage"/>' in captured


def test_stability_follows_custom_header_owner_relationship(tmp_path: Path):
    (tmp_path / "word" / "_rels").mkdir(parents=True)
    (tmp_path / "word" / "layout" / "_rels").mkdir(parents=True)
    (tmp_path / "word" / "document.xml").write_text(
        '<w:document xmlns:w="urn:w"><w:body><w:sectPr/></w:body></w:document>',
        encoding="utf-8",
    )
    (tmp_path / "word" / "_rels" / "document.xml.rels").write_text(
        f'<Relationships xmlns="{PKG_REL_NS}">'
        f'<Relationship Id="rIdHeader" Type="{OFFICE_REL_NS}/header" '
        'Target="layout/footer-looking.xml"/></Relationships>',
        encoding="utf-8",
    )
    owner = tmp_path / "word" / "layout" / "footer-looking.xml"
    owner.write_text('<w:hdr xmlns:w="urn:w"><w:p/></w:hdr>', encoding="utf-8")
    owner_rels = tmp_path / "word" / "layout" / "_rels" / "footer-looking.xml.rels"
    owner_rels.write_text(
        f'<Relationships xmlns="{PKG_REL_NS}"/>',
        encoding="utf-8",
    )
    snapshot = snapshot_stability(tmp_path)

    owner.write_text('<w:hdr xmlns:w="urn:w"><w:p><w:r/></w:p></w:hdr>', encoding="utf-8")

    with pytest.raises(ValueError, match="Header/footer stability check FAILED"):
        verify_stability(tmp_path, snapshot)
