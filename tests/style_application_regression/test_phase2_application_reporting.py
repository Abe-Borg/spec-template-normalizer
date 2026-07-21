from pathlib import Path

import pytest

from spec_formatter.style_application.core.classification import apply_phase2_classifications
from spec_formatter.style_application.core.classification import (
    _build_numbering_catalog,
    _effective_numbering_semantics,
)
from spec_formatter.style_application.core.xml_helpers import (
    iter_paragraph_xml_blocks,
    paragraph_text_from_block,
)


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


FORMAT_ONLY_MATRIX_SOURCE_STYLES = (
    '<?xml version="1.0" encoding="UTF-8"?>'
    '<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
    '<w:style w:type="paragraph" w:styleId="Normal"/>'
    '<w:style w:type="paragraph" w:styleId="TargetAutoBase"><w:pPr><w:numPr>'
    '<w:ilvl w:val="2"/><w:numId w:val="5"/>'
    '</w:numPr></w:pPr></w:style>'
    '<w:style w:type="paragraph" w:styleId="TargetAutoChild">'
    '<w:basedOn w:val="TargetAutoBase"/></w:style>'
    '</w:styles>'
)


FORMAT_ONLY_MATRIX_NUMBERING = (
    '<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
    '<w:abstractNum w:abstractNumId="3"><w:multiLevelType w:val="multilevel"/>'
    '<w:lvl w:ilvl="2"><w:start w:val="3"/><w:numFmt w:val="decimal"/>'
    '<w:lvlText w:val="%3)"/><w:lvlRestart w:val="1"/>'
    '<w:suff w:val="space"/></w:lvl></w:abstractNum>'
    '<w:num w:numId="5"><w:abstractNumId w:val="3"/>'
    '<w:lvlOverride w:ilvl="2"><w:startOverride w:val="7"/>'
    '</w:lvlOverride></w:num></w:numbering>'
)


def _seed_extract(tmp_path: Path, styles_xml: str) -> Path:
    (tmp_path / "word").mkdir(parents=True, exist_ok=True)
    (tmp_path / "word" / "document.xml").write_text(DOC_XML, encoding="utf-8")
    (tmp_path / "word" / "styles.xml").write_text(styles_xml, encoding="utf-8")
    return tmp_path


@pytest.mark.parametrize(
    "architect_provenance",
    ["style_numpr", "direct_numpr", "text_literal", "none"],
)
@pytest.mark.parametrize(
    "target_numbering",
    ["direct_automatic", "style_inherited_automatic", "typed_marker", "unnumbered"],
)
def test_format_only_numbering_authority_matrix(
    tmp_path,
    target_numbering,
    architect_provenance,
):
    """Every architect provenance preserves every target numbering state."""

    body_numpr = (
        '<w:pPr><w:numPr><w:ilvl w:val="4"/><w:numId w:val="91"/>'
        '</w:numPr><w:spacing w:before="240"/></w:pPr>'
        if architect_provenance == "style_numpr"
        else '<w:pPr><w:spacing w:before="240"/></w:pPr>'
    )
    current_styles = FORMAT_ONLY_MATRIX_SOURCE_STYLES.replace(
        "</w:styles>",
        '<w:style w:type="paragraph" w:styleId="Body">'
        f'{body_numpr}<w:rPr><w:lang w:val="en-CA"/></w:rPr>'
        '</w:style></w:styles>',
    )

    if target_numbering == "direct_automatic":
        text = "Sanitized direct automatic item"
        paragraph = (
            '<w:p><w:pPr><w:pStyle w:val="Normal"/><w:numPr>'
            '<w:ilvl w:val="2"/><w:numId w:val="5"/>'
            '</w:numPr></w:pPr><w:r><w:t>'
            f'{text}</w:t></w:r></w:p>'
        )
    elif target_numbering == "style_inherited_automatic":
        text = "Sanitized inherited automatic item"
        paragraph = (
            '<w:p><w:pPr><w:pStyle w:val="TargetAutoChild"/></w:pPr>'
            f'<w:r><w:t>{text}</w:t></w:r></w:p>'
        )
    elif target_numbering == "typed_marker":
        text = "A. Sanitized typed marker"
        paragraph = f'<w:p><w:r><w:t>{text}</w:t></w:r></w:p>'
    else:
        text = "Sanitized unnumbered item"
        paragraph = f'<w:p><w:r><w:t>{text}</w:t></w:r></w:p>'

    document = (
        '<?xml version="1.0" encoding="UTF-8"?>'
        '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
        f'<w:body>{paragraph}</w:body></w:document>'
    )
    word_dir = tmp_path / "word"
    word_dir.mkdir(parents=True)
    doc_path = word_dir / "document.xml"
    styles_path = word_dir / "styles.xml"
    numbering_path = word_dir / "numbering.xml"
    doc_path.write_text(document, encoding="utf-8")
    styles_path.write_text(current_styles, encoding="utf-8")
    numbering_path.write_text(FORMAT_ONLY_MATRIX_NUMBERING, encoding="utf-8")

    source_paragraph = next(iter_paragraph_xml_blocks(document))[2]
    numbering_catalog = _build_numbering_catalog(FORMAT_ONLY_MATRIX_NUMBERING)
    before_semantics = _effective_numbering_semantics(
        source_paragraph,
        FORMAT_ONLY_MATRIX_SOURCE_STYLES,
        numbering_catalog,
    )
    role_spec = {
        "style_id": "Body",
        "numbering_provenance": architect_provenance,
    }
    if architect_provenance in {"style_numpr", "direct_numpr"}:
        role_spec["numbering_pattern"] = {"numId": "91", "ilvl": "4"}

    report = apply_phase2_classifications(
        tmp_path,
        {"classifications": [{"paragraph_index": 0, "csi_role": "PARAGRAPH"}]},
        {"PARAGRAPH": "Body"},
        [],
        role_specs={"PARAGRAPH": role_spec},
        source_styles_xml=FORMAT_ONLY_MATRIX_SOURCE_STYLES,
        source_numbering_xml=FORMAT_ONLY_MATRIX_NUMBERING,
        conversion_mode="format_only",
    )

    output_document = doc_path.read_text(encoding="utf-8")
    output_paragraph = next(iter_paragraph_xml_blocks(output_document))[2]
    after_semantics = _effective_numbering_semantics(
        output_paragraph,
        styles_path.read_text(encoding="utf-8"),
        numbering_catalog,
    )

    assert paragraph_text_from_block(output_paragraph) == text
    assert before_semantics == after_semantics
    assert numbering_path.read_text(encoding="utf-8") == FORMAT_ONLY_MATRIX_NUMBERING
    assert '<w:pStyle w:val="Body"/>' in output_paragraph
    if "automatic" in target_numbering:
        assert after_semantics is not None
        assert after_semantics["numPr"] == {"numId": "5", "ilvl": "2"}
    else:
        assert after_semantics is None
    assert report.numbering_checks["effective_numbering_preserved"] is True
    assert report.numbering_checks["body_text_preserved"] is True


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
    # Body supplies paragraph spacing but no run font, so the target's direct
    # font remains authoritative.
    assert "HostFont" in outer_before_textbox
    assert 'w:after="120"' not in outer_before_textbox


def test_direct_run_properties_are_removed_only_when_effective_style_replaces_them(
    tmp_path,
):
    styles = (
        '<?xml version="1.0" encoding="UTF-8"?>'
        '<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
        '<w:style w:type="paragraph" w:styleId="Base"><w:rPr>'
        '<w:rFonts w:ascii="Architect"/><w:lang w:val="en-CA"/>'
        '</w:rPr></w:style>'
        '<w:style w:type="paragraph" w:styleId="Body">'
        '<w:basedOn w:val="Base"/><w:rPr><w:b/></w:rPr>'
        '</w:style></w:styles>'
    )
    extract = _seed_extract(tmp_path, styles)
    doc_path = extract / "word" / "document.xml"
    doc_path.write_text(
        DOC_XML.replace(
            '<w:r><w:t>A</w:t></w:r>',
            '<w:r><w:rPr><w:rFonts w:ascii="Target"/>'
            '<w:sz w:val="19"/><w:szCs w:val="21"/>'
            '<w:lang w:val="fr-CA"/><w:b w:val="0"/>'
            '<w:color w:val="123456"/></w:rPr><w:t>A</w:t></w:r>',
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

    first_paragraph = doc_path.read_text(encoding="utf-8").split("</w:p>", 1)[0]
    assert "<w:rFonts" not in first_paragraph
    assert "<w:lang" not in first_paragraph
    assert "<w:b " not in first_paragraph
    assert "<w:b/>" not in first_paragraph
    assert '<w:sz w:val="19"/>' in first_paragraph
    assert '<w:szCs w:val="21"/>' in first_paragraph
    assert '<w:color w:val="123456"/>' in first_paragraph
    assert report.stripped_run_fonts == 1


def test_direct_fonts_sizes_and_language_survive_when_style_replaces_none(tmp_path):
    extract = _seed_extract(tmp_path, STYLE_WITHOUT_PPR)
    doc_path = extract / "word" / "document.xml"
    direct_rpr = (
        '<w:rPr><w:rFonts w:ascii="Target"/><w:sz w:val="19"/>'
        '<w:szCs w:val="21"/><w:lang w:val="fr-CA"/></w:rPr>'
    )
    doc_path.write_text(
        DOC_XML.replace(
            '<w:r><w:t>A</w:t></w:r>',
            f'<w:r>{direct_rpr}<w:t>A</w:t></w:r>',
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

    first_paragraph = doc_path.read_text(encoding="utf-8").split("</w:p>", 1)[0]
    assert direct_rpr in first_paragraph
    assert report.stripped_run_fonts == 0


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
        conversion_mode="csi_to_canadian",
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
        conversion_mode="csi_to_canadian",
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
        conversion_mode="csi_to_canadian",
    )
    first_paragraph = doc_path.read_text(encoding="utf-8").split("</w:p>", 1)[0]
    assert '<w:pStyle w:val="Body"/>' in first_paragraph
    assert "<w:numPr" not in first_paragraph


def test_text_literal_contract_preserves_direct_automatic_target_numbering(tmp_path):
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
    log = []
    report = apply_phase2_classifications(
        extract,
        {"classifications": [{"paragraph_index": 0, "csi_role": "PARAGRAPH"}]},
        {"PARAGRAPH": "Body"},
        log,
        role_specs={
            "PARAGRAPH": {
                "style_id": "Body",
                "numbering_provenance": "text_literal",
            }
        },
    )

    first_paragraph = doc_path.read_text(encoding="utf-8").split("</w:p>", 1)[0]
    assert '<w:pStyle w:val="Body"/>' in first_paragraph
    assert '<w:numId w:val="5"/>' in first_paragraph
    assert report.preserved_automatic_numbering == 1
    assert any("Preserved source Word numbering" in line for line in log)


def test_text_literal_contract_materializes_inherited_target_numbering(tmp_path):
    source_styles = (
        '<?xml version="1.0" encoding="UTF-8"?>'
        '<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
        '<w:style w:type="paragraph" w:styleId="SourceList"><w:pPr><w:numPr>'
        '<w:ilvl w:val="0"/><w:numId w:val="5"/></w:numPr></w:pPr></w:style>'
        '</w:styles>'
    )
    imported_styles = (
        '<?xml version="1.0" encoding="UTF-8"?>'
        '<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
        '<w:style w:type="paragraph" w:styleId="SourceList">'
        '<w:name w:val="Replaced Source List"/></w:style>'
        '<w:style w:type="paragraph" w:styleId="Body"><w:pPr>'
        '<w:spacing w:after="240"/></w:pPr></w:style>'
        '</w:styles>'
    )
    extract = _seed_extract(tmp_path, imported_styles)
    doc_path = extract / "word" / "document.xml"
    doc_path.write_text(
        DOC_XML.replace(
            '<w:spacing w:after="120"/>',
            '<w:pStyle w:val="SourceList"/><w:spacing w:after="120"/>',
            1,
        ),
        encoding="utf-8",
    )

    report = apply_phase2_classifications(
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
        source_styles_xml=source_styles,
    )

    first_paragraph = doc_path.read_text(encoding="utf-8").split("</w:p>", 1)[0]
    assert '<w:pStyle w:val="Body"/>' in first_paragraph
    assert '<w:numId w:val="5"/>' in first_paragraph
    assert '<w:ilvl w:val="0"/>' in first_paragraph
    assert 'w:after="120"' not in first_paragraph
    assert report.preserved_automatic_numbering == 1


def test_text_literal_contract_materializes_merged_partial_numpr(tmp_path):
    source_styles = (
        '<?xml version="1.0" encoding="UTF-8"?>'
        '<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
        '<w:style w:type="paragraph" w:styleId="SourceList"><w:pPr><w:numPr>'
        '<w:ilvl w:val="0"/><w:numId w:val="5"/></w:numPr></w:pPr></w:style>'
        '<w:style w:type="paragraph" w:styleId="Body"/>'
        '</w:styles>'
    )
    extract = _seed_extract(tmp_path, source_styles)
    doc_path = extract / "word" / "document.xml"
    doc_path.write_text(
        DOC_XML.replace(
            '<w:spacing w:after="120"/>',
            '<w:pStyle w:val="SourceList"/><w:numPr>'
            '<w:ilvl w:val="2"/></w:numPr>',
            1,
        ),
        encoding="utf-8",
    )

    report = apply_phase2_classifications(
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
        source_styles_xml=source_styles,
    )

    first_paragraph = doc_path.read_text(encoding="utf-8").split("</w:p>", 1)[0]
    assert '<w:numId w:val="5"/>' in first_paragraph
    assert '<w:ilvl w:val="2"/>' in first_paragraph
    assert report.preserved_automatic_numbering == 1


def test_text_literal_deep_marker_removes_doubled_automatic_numbering(tmp_path):
    extract = _seed_extract(tmp_path, STYLE_WITHOUT_PPR)
    doc_path = extract / "word" / "document.xml"
    doc_path.write_text(
        DOC_XML.replace(
            '<w:spacing w:after="120"/>',
            '<w:numPr><w:ilvl w:val="0"/><w:numId w:val="5"/></w:numPr>',
            1,
        ).replace('<w:t>A</w:t>', '<w:t>1) Deep item</w:t>', 1),
        encoding="utf-8",
    )

    report = apply_phase2_classifications(
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
        conversion_mode="csi_to_canadian",
    )

    first_paragraph = doc_path.read_text(encoding="utf-8").split("</w:p>", 1)[0]
    assert "1) Deep item" in first_paragraph
    assert "<w:numPr" not in first_paragraph
    assert report.preserved_automatic_numbering == 0


@pytest.mark.parametrize(
    "architect_provenance",
    ["style_numpr", "direct_numpr", "text_literal", "none"],
)
def test_format_only_preserves_target_automatic_numbering_for_all_architect_provenance(
    tmp_path,
    architect_provenance,
):
    styles = (
        '<?xml version="1.0" encoding="UTF-8"?>'
        '<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
        '<w:style w:type="paragraph" w:styleId="Body"><w:pPr>'
        '<w:spacing w:after="240"/><w:numPr><w:ilvl w:val="1"/>'
        '<w:numId w:val="42"/></w:numPr></w:pPr></w:style>'
        '</w:styles>'
    )
    extract = _seed_extract(tmp_path, styles)
    doc_path = extract / "word" / "document.xml"
    doc_path.write_text(
        DOC_XML.replace(
            '<w:spacing w:after="120"/>',
            '<w:spacing w:after="120"/><w:numPr>'
            '<w:ilvl w:val="2"/><w:numId w:val="5"/></w:numPr>',
            1,
        ),
        encoding="utf-8",
    )

    report = apply_phase2_classifications(
        extract,
        {"classifications": [{"paragraph_index": 0, "csi_role": "PARAGRAPH"}]},
        {"PARAGRAPH": "Body"},
        [],
        role_specs={
            "PARAGRAPH": {
                "style_id": "Body",
                "numbering_provenance": architect_provenance,
            }
        },
        role_numpr_remap={
            "PARAGRAPH": {"old_numId": 42, "new_numId": 99, "ilvl": 1},
        },
    )

    first_paragraph = doc_path.read_text(encoding="utf-8").split("</w:p>", 1)[0]
    assert '<w:numId w:val="5"/>' in first_paragraph
    assert '<w:ilvl w:val="2"/>' in first_paragraph
    assert 'w:numId w:val="42"' not in first_paragraph
    assert 'w:numId w:val="99"' not in first_paragraph
    assert report.preserved_automatic_numbering == 1
    assert report.numbering_checks == {
        "policy": "format_only",
        "paragraphs_checked": 2,
        "automatic_numbered_before": 1,
        "effective_numbering_preserved": True,
        "body_text_preserved": True,
    }


def test_format_only_suppresses_inherited_architect_numbering_on_unnumbered_target(
    tmp_path,
):
    styles = (
        '<?xml version="1.0" encoding="UTF-8"?>'
        '<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
        '<w:style w:type="paragraph" w:styleId="NumberedBase"><w:pPr>'
        '<w:numPr><w:ilvl w:val="1"/><w:numId w:val="42"/></w:numPr>'
        '</w:pPr></w:style>'
        '<w:style w:type="paragraph" w:styleId="Body">'
        '<w:basedOn w:val="NumberedBase"/></w:style>'
        '</w:styles>'
    )
    extract = _seed_extract(tmp_path, styles)
    doc_path = extract / "word" / "document.xml"

    report = apply_phase2_classifications(
        extract,
        {"classifications": [{"paragraph_index": 0, "csi_role": "PARAGRAPH"}]},
        {"PARAGRAPH": "Body"},
        [],
    )

    first_paragraph = doc_path.read_text(encoding="utf-8").split("</w:p>", 1)[0]
    assert '<w:numId w:val="0"/>' in first_paragraph
    assert report.suppressed_architect_numbering == 1


def test_format_only_uses_effective_style_ppr_and_preserves_protected_properties(
    tmp_path,
):
    styles = (
        '<?xml version="1.0" encoding="UTF-8"?>'
        '<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
        '<w:style w:type="paragraph" w:styleId="Base"><w:pPr>'
        '<w:tabs><w:tab w:val="left" w:pos="720"/></w:tabs>'
        '<w:keepNext/><w:shd w:fill="EEEEEE"/>'
        '</w:pPr></w:style>'
        '<w:style w:type="paragraph" w:styleId="Body">'
        '<w:basedOn w:val="Base"/><w:pPr><w:spacing w:after="240"/>'
        '<w:numPr><w:ilvl w:val="1"/><w:numId w:val="42"/></w:numPr>'
        '</w:pPr></w:style>'
        '</w:styles>'
    )
    extract = _seed_extract(tmp_path, styles)
    doc_path = extract / "word" / "document.xml"
    doc_path.write_text(
        DOC_XML.replace(
            '<w:spacing w:after="120"/>',
            '<w:cnfStyle w:val="1"/><w:keepNext/><w:tabs>'
            '<w:tab w:val="right" w:pos="1440"/></w:tabs>'
            '<w:shd w:fill="FFFFFF"/><w:spacing w:after="120"/>'
            '<w:pPrChange w:id="7"><w:pPr>'
            '<w:pStyle w:val="HistoricalList"/><w:spacing w:after="1"/>'
            '<w:numPr><w:ilvl w:val="4"/><w:numId w:val="91"/></w:numPr>'
            '</w:pPr></w:pPrChange><w:sectPr><w:type w:val="continuous"/>'
            '</w:sectPr>',
            1,
        ),
        encoding="utf-8",
    )

    apply_phase2_classifications(
        extract,
        {"classifications": [{"paragraph_index": 0, "csi_role": "PARAGRAPH"}]},
        {"PARAGRAPH": "Body"},
        [],
    )

    first_paragraph = doc_path.read_text(encoding="utf-8").split("</w:p>", 1)[0]
    assert '<w:tabs>' not in first_paragraph
    assert '<w:keepNext' not in first_paragraph
    assert '<w:shd' not in first_paragraph
    assert 'w:after="120"' not in first_paragraph
    assert '<w:cnfStyle w:val="1"/>' in first_paragraph
    assert '<w:pPrChange w:id="7">' in first_paragraph
    assert '<w:pStyle w:val="HistoricalList"/>' in first_paragraph
    assert '<w:numId w:val="91"/>' in first_paragraph
    assert '<w:sectPr><w:type w:val="continuous"/></w:sectPr>' in first_paragraph
    assert '<w:numId w:val="0"/>' in first_paragraph


def test_explicitly_ignored_paragraph_is_byte_exact(tmp_path):
    extract = _seed_extract(tmp_path, STYLE_WITH_PPR)
    doc_path = extract / "word" / "document.xml"
    ignored_before = doc_path.read_text(encoding="utf-8").split("<w:p>")[2].split(
        "</w:p>", 1
    )[0]

    report = apply_phase2_classifications(
        extract,
        {
            "classifications": [
                {"paragraph_index": 0, "csi_role": "PARAGRAPH"},
            ],
            "ignored_paragraphs": [
                {"paragraph_index": 1, "reason": "non_csi_content"},
            ],
        },
        {"PARAGRAPH": "Body"},
        [],
    )

    ignored_after = doc_path.read_text(encoding="utf-8").split("<w:p>")[2].split(
        "</w:p>", 1
    )[0]
    assert ignored_after == ignored_before
    assert report.ignored == 1
