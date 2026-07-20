from pathlib import Path

import pytest

from spec_formatter.style_application.core.classification import apply_phase2_classifications


DOC_XML = (
    '<?xml version="1.0" encoding="UTF-8"?>'
    '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
    '<w:body>'
    '<w:p><w:pPr><w:spacing w:after="120"/></w:pPr><w:r><w:t>A</w:t></w:r></w:p>'
    '<w:p><w:pPr><w:spacing w:after="120"/></w:pPr><w:r><w:t>B</w:t></w:r></w:p>'
    '</w:body></w:document>'
)


STYLE_WITHOUT_PPR = (
    '<?xml version="1.0" encoding="UTF-8"?>'
    '<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
    '<w:style w:type="paragraph" w:styleId="Body"><w:name w:val="Body"/></w:style>'
    '</w:styles>'
)


STYLE_WITH_PPR = (
    '<?xml version="1.0" encoding="UTF-8"?>'
    '<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
    '<w:style w:type="paragraph" w:styleId="Body"><w:name w:val="Body"/>'
    '<w:pPr><w:spacing w:after="240"/></w:pPr></w:style>'
    '</w:styles>'
)


def _seed_extract(tmp_path: Path, styles_xml: str) -> Path:
    (tmp_path / "word").mkdir(parents=True, exist_ok=True)
    (tmp_path / "word" / "document.xml").write_text(DOC_XML, encoding="utf-8")
    (tmp_path / "word" / "styles.xml").write_text(styles_xml, encoding="utf-8")
    return tmp_path


def test_invalid_index_is_fatal(tmp_path):
    extract = _seed_extract(tmp_path, STYLE_WITHOUT_PPR)
    with pytest.raises(ValueError, match="Invalid paragraph indices"):
        apply_phase2_classifications(
            extract,
            {"classifications": [{"paragraph_index": 99, "csi_role": "PARAGRAPH"}]},
            {"PARAGRAPH": "Body"},
            [],
        )


def test_preserve_direct_ppr_when_style_lacks_replacement(tmp_path):
    extract = _seed_extract(tmp_path, STYLE_WITHOUT_PPR)
    report = apply_phase2_classifications(
        extract,
        {"classifications": [{"paragraph_index": 0, "csi_role": "PARAGRAPH"}]},
        {"PARAGRAPH": "Body"},
        [],
    )
    out = (extract / "word" / "document.xml").read_text(encoding="utf-8")
    assert '<w:spacing w:after="120"/>' in out
    assert report.preserved_direct_ppr == 1
    assert report.stripped_direct_ppr == 0


def test_strip_direct_ppr_when_style_has_replacement(tmp_path):
    extract = _seed_extract(tmp_path, STYLE_WITH_PPR)
    report = apply_phase2_classifications(
        extract,
        {"classifications": [{"paragraph_index": 0, "csi_role": "PARAGRAPH"}]},
        {"PARAGRAPH": "Body"},
        [],
    )
    out = (extract / "word" / "document.xml").read_text(encoding="utf-8")
    assert out.count('<w:spacing w:after="120"/>') == 1
    assert report.stripped_direct_ppr == 1
    assert report.preserved_direct_ppr == 0


def test_visible_section_break_paragraph_is_styled_and_sectpr_is_exact(tmp_path):
    extract = _seed_extract(tmp_path, STYLE_WITH_PPR)
    doc_path = extract / "word" / "document.xml"
    sectpr = '<w:sectPr><w:pgSz w:w="12240"/><w:docGrid w:linePitch="360"/></w:sectPr>'
    doc_path.write_text(
        DOC_XML.replace(
            '<w:spacing w:after="120"/>',
            f'<w:spacing w:after="120"/>{sectpr}',
            1,
        ),
        encoding="utf-8",
    )

    report = apply_phase2_classifications(
        extract,
        {"classifications": [{"paragraph_index": 0, "csi_role": "PARAGRAPH"}]},
        {"PARAGRAPH": "Body"},
        [],
    )

    out = doc_path.read_text(encoding="utf-8")
    assert report.modified == 1
    assert '<w:pStyle w:val="Body"/>' in out.split("</w:p>", 1)[0]
    assert sectpr in out


def test_visible_drawing_host_is_styled_while_textbox_subtree_is_byte_exact(tmp_path):
    extract = _seed_extract(tmp_path, STYLE_WITH_PPR)
    doc_path = extract / "word" / "document.xml"
    textbox = (
        '<w:drawing><w:txbxContent><w:p><w:pPr>'
        '<w:pStyle w:val="NestedStyle"/><w:spacing w:after="999"/>'
        '</w:pPr><w:r><w:rPr><w:rFonts w:ascii="NestedFont"/></w:rPr>'
        '<w:t>Nested text</w:t></w:r></w:p></w:txbxContent></w:drawing>'
    )
    host = (
        '<w:p><w:pPr><w:spacing w:after="120"/></w:pPr>'
        '<w:r><w:rPr><w:rFonts w:ascii="HostFont"/></w:rPr>'
        f'<w:t>Visible host</w:t>{textbox}</w:r></w:p>'
    )
    document = (
        '<?xml version="1.0" encoding="UTF-8"?>'
        '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
        f'<w:body>{host}</w:body></w:document>'
    )
    doc_path.write_text(document, encoding="utf-8")

    report = apply_phase2_classifications(
        extract,
        {"classifications": [{"paragraph_index": 0, "csi_role": "PARAGRAPH"}]},
        {"PARAGRAPH": "Body"},
        [],
    )

    out = doc_path.read_text(encoding="utf-8")
    assert report.modified == 1
    assert textbox in out
    outer_before_textbox = out.split(textbox, 1)[0]
    assert '<w:pStyle w:val="Body"/>' in outer_before_textbox
    assert "HostFont" not in outer_before_textbox
    assert 'w:after="120"' not in outer_before_textbox


def test_direct_numbering_contract_applies_imported_numid_and_level(tmp_path):
    extract = _seed_extract(tmp_path, STYLE_WITHOUT_PPR)
    apply_phase2_classifications(
        extract,
        {"classifications": [{"paragraph_index": 0, "csi_role": "PARAGRAPH"}]},
        {"PARAGRAPH": "Body"},
        [],
        role_specs={
            "PARAGRAPH": {
                "style_id": "Body",
                "numbering_provenance": "direct_numpr",
                "numbering_pattern": {"numId": "7", "ilvl": "2"},
            }
        },
        role_numpr_remap={
            "PARAGRAPH": {"old_numId": 7, "new_numId": 42, "ilvl": 2}
        },
    )
    out = (extract / "word" / "document.xml").read_text(encoding="utf-8")
    assert '<w:numId w:val="42"/>' in out
    assert '<w:ilvl w:val="2"/>' in out


def test_none_numbering_contract_removes_target_direct_numbering(tmp_path):
    extract = _seed_extract(tmp_path, STYLE_WITHOUT_PPR)
    doc_path = extract / "word" / "document.xml"
    doc_path.write_text(
        DOC_XML.replace(
            '<w:spacing w:after="120"/>',
            '<w:spacing w:after="120"/>'
            '<w:numPr><w:ilvl w:val="0"/><w:numId w:val="5"/></w:numPr>',
            1,
        ),
        encoding="utf-8",
    )
    apply_phase2_classifications(
        extract,
        {"classifications": [{"paragraph_index": 0, "csi_role": "PARAGRAPH"}]},
        {"PARAGRAPH": "Body"},
        [],
        role_specs={
            "PARAGRAPH": {
                "style_id": "Body",
                "numbering_provenance": "none",
            }
        },
    )
    out = doc_path.read_text(encoding="utf-8")
    first_paragraph = out.split("</w:p>", 1)[0]
    assert "<w:numPr" not in first_paragraph
    assert '<w:spacing w:after="120"/>' in first_paragraph


def test_style_numbering_contract_does_not_leave_target_numpr_override(tmp_path):
    styles = (
        '<?xml version="1.0" encoding="UTF-8"?>'
        '<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
        '<w:style w:type="paragraph" w:styleId="ListBase"><w:pPr><w:numPr>'
        '<w:ilvl w:val="1"/><w:numId w:val="42"/></w:numPr></w:pPr></w:style>'
        '<w:style w:type="paragraph" w:styleId="Body"><w:basedOn w:val="ListBase"/></w:style>'
        '</w:styles>'
    )
    extract = _seed_extract(tmp_path, styles)
    doc_path = extract / "word" / "document.xml"
    doc_path.write_text(
        DOC_XML.replace(
            '<w:spacing w:after="120"/>',
            '<w:numPr><w:ilvl w:val="0"/><w:numId w:val="5"/></w:numPr>',
            1,
        ),
        encoding="utf-8",
    )
    apply_phase2_classifications(
        extract,
        {"classifications": [{"paragraph_index": 0, "csi_role": "PARAGRAPH"}]},
        {"PARAGRAPH": "Body"},
        [],
        role_specs={
            "PARAGRAPH": {
                "style_id": "Body",
                "numbering_provenance": "style_numpr",
                "numbering_pattern": {"numId": "42", "ilvl": "1"},
            }
        },
    )
    first_paragraph = doc_path.read_text(encoding="utf-8").split("</w:p>", 1)[0]
    assert '<w:pStyle w:val="Body"/>' in first_paragraph
    assert "<w:numPr" not in first_paragraph


def test_text_literal_contract_fails_closed_for_automatic_target_numbering(tmp_path):
    extract = _seed_extract(tmp_path, STYLE_WITHOUT_PPR)
    doc_path = extract / "word" / "document.xml"
    doc_path.write_text(
        DOC_XML.replace(
            '<w:spacing w:after="120"/>',
            '<w:numPr><w:ilvl w:val="0"/><w:numId w:val="5"/></w:numPr>',
            1,
        ),
        encoding="utf-8",
    )
    with pytest.raises(ValueError, match="requires literal-text numbering"):
        apply_phase2_classifications(
            extract,
            {"classifications": [{"paragraph_index": 0, "csi_role": "PARAGRAPH"}]},
            {"PARAGRAPH": "Body"},
            [],
            role_specs={
                "PARAGRAPH": {
                    "style_id": "Body",
                    "numbering_provenance": "text_literal",
                }
            },
        )
