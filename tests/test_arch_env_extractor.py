"""
Regression tests for arch_env_extractor.py

These tests cover the XML block extraction functions that populate
arch_template_registry.json. They are designed to:
  - FAIL on the current broken extractor (paired tags truncated to bare openers)
  - PASS only when paired tags are extracted in full

Coverage:
  1. self-closing tag extraction
  2. paired tag extraction
  3. repeated paired tag extraction
  4. adjacent and nested style blocks
  5. docDefaults capture with both rPrDefault and pPrDefault
  6. style capture including name, basedOn, next, link, pPr, rPr
  7. compat_xml capture as a full block
  8. sectPr capture as a full block
"""

from __future__ import annotations

import xml.etree.ElementTree as ET

import pytest

from arch_env_extractor import (
    _extract_first_block,
    _extract_all_blocks,
    extract_doc_defaults,
    extract_style_defs,
    extract_settings,
    extract_page_layout,
)


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

NS_MAP = (
    ' xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
    ' xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"'
    ' xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"'
)


def assert_parses_as_xml(xml_fragment: str, msg: str = "") -> None:
    """Assert that an XML fragment parses when wrapped in a namespace root."""
    try:
        ET.fromstring(f"<root{NS_MAP}>{xml_fragment}</root>")
    except ET.ParseError as exc:
        pytest.fail(f"XML fragment is not parseable{': ' + msg if msg else ''}: {exc}\nFragment: {xml_fragment!r}")


# ═══════════════════════════════════════════════════════════════════════════
# Unit tests — _extract_first_block
# ═══════════════════════════════════════════════════════════════════════════


class TestExtractBlock:
    """Unit tests for _extract_first_block()."""

    def test_self_closing_tag(self):
        """Self-closing tags (with />) must be extracted correctly."""
        xml = '<w:styles><w:docGrid w:linePitch="360"/></w:styles>'
        result = _extract_first_block(xml, "docGrid")
        assert result == '<w:docGrid w:linePitch="360"/>'
        assert_parses_as_xml(result)

    def test_paired_tag_extraction(self):
        """Paired tags must be extracted in full, not truncated to the opener."""
        xml = '<w:styles><w:pPr><w:jc w:val="center"/></w:pPr></w:styles>'
        result = _extract_first_block(xml, "pPr")
        assert result is not None
        # Must contain the closing tag — not just "<w:pPr>"
        assert "</w:pPr>" in result
        assert '<w:jc w:val="center"/>' in result
        assert_parses_as_xml(result)

    def test_paired_tag_with_attributes(self):
        """Paired tags whose opener has attributes must be captured in full."""
        xml = (
            '<w:styles>'
            '<w:style w:type="paragraph" w:styleId="Heading1">'
            '<w:name w:val="heading 1"/>'
            '<w:pPr><w:spacing w:before="240"/></w:pPr>'
            '</w:style>'
            '</w:styles>'
        )
        result = _extract_first_block(xml, "style")
        assert result is not None
        assert "</w:style>" in result
        assert '<w:name w:val="heading 1"/>' in result
        assert "<w:pPr>" in result
        assert_parses_as_xml(result)

    def test_empty_paired_tag(self):
        """An empty paired tag like <w:pPr></w:pPr> must not be confused with self-closing."""
        xml = '<w:styles><w:pPr></w:pPr></w:styles>'
        result = _extract_first_block(xml, "pPr")
        assert result is not None
        assert "</w:pPr>" in result

    def test_no_match_returns_none(self):
        """Non-existent tag returns None."""
        xml = '<w:styles><w:pPr><w:jc w:val="center"/></w:pPr></w:styles>'
        result = _extract_first_block(xml, "rPr")
        assert result is None

    def test_non_default_namespace_prefix(self):
        """The ns_prefix parameter must work for non-w namespaces (e.g. 'a')."""
        xml = (
            '<a:theme>'
            '<a:themeElements>'
            '<a:fontScheme name="Office"/>'
            '</a:themeElements>'
            '</a:theme>'
        )
        result = _extract_first_block(xml, "themeElements", ns_prefix="a")
        assert result is not None
        assert "</a:themeElements>" in result
        assert '<a:fontScheme name="Office"/>' in result


# ═══════════════════════════════════════════════════════════════════════════
# Unit tests — _extract_all_blocks
# ═══════════════════════════════════════════════════════════════════════════


class TestExtractAllBlocks:
    """Unit tests for _extract_all_blocks()."""

    def test_extract_all_repeated_paired_tags(self):
        """Multiple paired tags of the same type must all be captured in full."""
        xml = (
            '<w:styles>'
            '<w:style w:type="paragraph" w:styleId="Normal">'
            '<w:name w:val="Normal"/>'
            '<w:pPr><w:spacing w:after="200"/></w:pPr>'
            '</w:style>'
            '<w:style w:type="paragraph" w:styleId="Heading1">'
            '<w:name w:val="heading 1"/>'
            '<w:pPr><w:spacing w:before="240"/></w:pPr>'
            '</w:style>'
            '</w:styles>'
        )
        results = _extract_all_blocks(xml, "style")
        assert len(results) == 2
        for block in results:
            assert "</w:style>" in block
            assert "<w:name" in block
            assert "<w:pPr>" in block
            assert "</w:pPr>" in block
            assert_parses_as_xml(block)

    def test_extract_all_self_closing_tags(self):
        """Multiple self-closing tags of the same type are all found."""
        xml = (
            '<w:sectPr>'
            '<w:pgSz w:w="12240" w:h="15840"/>'
            '</w:sectPr>'
        )
        results = _extract_all_blocks(xml, "pgSz")
        assert len(results) == 1
        assert 'w:w="12240"' in results[0]

    def test_extract_all_adjacent_paired_tags(self):
        """Adjacent blocks of the same tag type are extracted individually."""
        xml = (
            '<w:body>'
            '<w:sectPr>'
            '<w:pgSz w:w="12240" w:h="15840"/>'
            '<w:docGrid w:linePitch="360"/>'
            '</w:sectPr>'
            '<w:sectPr>'
            '<w:pgSz w:w="15840" w:h="12240" w:orient="landscape"/>'
            '</w:sectPr>'
            '</w:body>'
        )
        results = _extract_all_blocks(xml, "sectPr")
        assert len(results) == 2
        assert 'linePitch="360"' in results[0]
        assert 'orient="landscape"' in results[1]
        for block in results:
            assert "</w:sectPr>" in block
            assert_parses_as_xml(block)


# ═══════════════════════════════════════════════════════════════════════════
# Integration tests — extract_doc_defaults
# ═══════════════════════════════════════════════════════════════════════════


class TestExtractDocDefaults:
    """Integration tests for extract_doc_defaults()."""

    STYLES_XML_BOTH = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
        '<w:docDefaults>'
        '<w:rPrDefault>'
        '<w:rPr>'
        '<w:rFonts w:asciiTheme="minorHAnsi" w:hAnsiTheme="minorHAnsi"/>'
        '<w:sz w:val="22"/>'
        '</w:rPr>'
        '</w:rPrDefault>'
        '<w:pPrDefault>'
        '<w:pPr>'
        '<w:spacing w:after="160" w:line="259" w:lineRule="auto"/>'
        '</w:pPr>'
        '</w:pPrDefault>'
        '</w:docDefaults>'
        '</w:styles>'
    )

    def test_doc_defaults_with_both_defaults(self):
        """Both rPrDefault/rPr and pPrDefault/pPr must be captured."""
        result = extract_doc_defaults(self.STYLES_XML_BOTH)

        rpr = result["default_run_props"]["rPr"]
        ppr = result["default_paragraph_props"]["pPr"]

        assert rpr is not None, "rPr must not be None"
        assert ppr is not None, "pPr must not be None"

        # Verify contents are complete — not just an opener tag
        assert "rFonts" in rpr
        assert "sz" in rpr
        assert "</w:rPr>" in rpr
        assert_parses_as_xml(rpr, "rPr")

        assert "spacing" in ppr
        assert "</w:pPr>" in ppr
        assert_parses_as_xml(ppr, "pPr")

    def test_doc_defaults_rpr_only(self):
        """When only rPrDefault is present, rPr is extracted and pPr is None."""
        styles_xml = (
            '<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
            '<w:docDefaults>'
            '<w:rPrDefault>'
            '<w:rPr>'
            '<w:rFonts w:asciiTheme="minorHAnsi"/>'
            '<w:sz w:val="24"/>'
            '</w:rPr>'
            '</w:rPrDefault>'
            '</w:docDefaults>'
            '</w:styles>'
        )
        result = extract_doc_defaults(styles_xml)
        rpr = result["default_run_props"]["rPr"]
        ppr = result["default_paragraph_props"]["pPr"]

        assert rpr is not None
        assert "rFonts" in rpr
        assert "</w:rPr>" in rpr
        assert_parses_as_xml(rpr, "rPr")

        assert ppr is None


# ═══════════════════════════════════════════════════════════════════════════
# Integration tests — extract_style_defs
# ═══════════════════════════════════════════════════════════════════════════


class TestExtractStyleDefs:
    """Integration tests for extract_style_defs()."""

    STYLES_XML = (
        '<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
        '<w:style w:type="paragraph" w:styleId="Heading1">'
        '<w:name w:val="heading 1"/>'
        '<w:basedOn w:val="Normal"/>'
        '<w:next w:val="Normal"/>'
        '<w:link w:val="Heading1Char"/>'
        '<w:uiPriority w:val="9"/>'
        '<w:qFormat/>'
        '<w:pPr>'
        '<w:keepNext/>'
        '<w:spacing w:before="240" w:after="0"/>'
        '<w:outlineLvl w:val="0"/>'
        '</w:pPr>'
        '<w:rPr>'
        '<w:rFonts w:asciiTheme="majorHAnsi"/>'
        '<w:b/>'
        '<w:sz w:val="32"/>'
        '</w:rPr>'
        '</w:style>'
        '</w:styles>'
    )

    def test_style_defs_captures_all_fields(self):
        """Style extraction must capture name, basedOn, next, link, pPr, rPr."""
        result = extract_style_defs(self.STYLES_XML)
        assert len(result) == 1

        style = result[0]
        assert style["style_id"] == "Heading1"
        assert style["name"] == "heading 1"
        assert style["based_on"] == "Normal"
        assert style["next"] == "Normal"
        assert style["link"] == "Heading1Char"
        assert style["ui_priority"] == 9
        assert style["qformat"] is True

        # pPr must be a full block, not a bare opener
        ppr = style["pPr"]
        assert ppr is not None
        assert "keepNext" in ppr
        assert "outlineLvl" in ppr
        assert "</w:pPr>" in ppr
        assert_parses_as_xml(ppr, "pPr")

        # rPr must be a full block, not a bare opener
        rpr = style["rPr"]
        assert rpr is not None
        assert "rFonts" in rpr
        assert "<w:b/>" in rpr
        assert "sz" in rpr
        assert "</w:rPr>" in rpr
        assert_parses_as_xml(rpr, "rPr")

    def test_style_defs_multiple_styles(self):
        """Multiple styles are individually extracted with correct IDs."""
        styles_xml = (
            '<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
            '<w:style w:type="paragraph" w:styleId="Normal">'
            '<w:name w:val="Normal"/>'
            '<w:pPr><w:spacing w:after="200"/></w:pPr>'
            '</w:style>'
            '<w:style w:type="paragraph" w:styleId="Heading1">'
            '<w:name w:val="heading 1"/>'
            '<w:rPr><w:b/></w:rPr>'
            '</w:style>'
            '</w:styles>'
        )
        result = extract_style_defs(styles_xml)
        assert len(result) == 2

        ids = {s["style_id"] for s in result}
        assert ids == {"Normal", "Heading1"}

        for style in result:
            if style["style_id"] == "Normal":
                assert style["pPr"] is not None
                assert "spacing" in style["pPr"]
                assert "</w:pPr>" in style["pPr"]
            elif style["style_id"] == "Heading1":
                assert style["rPr"] is not None
                assert "<w:b/>" in style["rPr"]
                assert "</w:rPr>" in style["rPr"]


# ═══════════════════════════════════════════════════════════════════════════
# Integration tests — extract_settings (compat_xml)
# ═══════════════════════════════════════════════════════════════════════════


class TestExtractSettings:
    """Integration tests for extract_settings() — compat_xml capture."""

    SETTINGS_XML = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
        '<w:compat>'
        '<w:compatSetting w:name="compatibilityMode"'
        ' w:uri="http://schemas.microsoft.com/office/word" w:val="15"/>'
        '<w:doNotExpandShiftReturn/>'
        '</w:compat>'
        '</w:settings>'
    )

    def test_compat_xml_full_block(self, tmp_path):
        """compat_xml must be captured as a complete block with children."""
        word_dir = tmp_path / "word"
        word_dir.mkdir()
        (word_dir / "settings.xml").write_text(self.SETTINGS_XML, encoding="utf-8")

        result = extract_settings(tmp_path)
        compat_xml = result["compat"]["compat_xml"]

        assert compat_xml is not None
        # Must contain the closing tag — not just "<w:compat>"
        assert "</w:compat>" in compat_xml
        assert "compatSetting" in compat_xml
        assert "doNotExpandShiftReturn" in compat_xml
        assert_parses_as_xml(compat_xml, "compat_xml")

        # Important flags should be detected
        assert "doNotExpandShiftReturn" in result["compat"]["important_flags"]
        assert "compatSetting:*" in result["compat"]["important_flags"]


# ═══════════════════════════════════════════════════════════════════════════
# Integration tests — extract_page_layout (sectPr)
# ═══════════════════════════════════════════════════════════════════════════


class TestExtractPageLayout:
    """Integration tests for extract_page_layout() — sectPr capture."""

    DOCUMENT_XML = (
        '<w:body xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
        ' xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
        '<w:sectPr>'
        '<w:pgSz w:w="12240" w:h="15840"/>'
        '<w:pgMar w:top="1440" w:right="1440" w:bottom="1440"'
        ' w:left="1440" w:header="720" w:footer="720" w:gutter="0"/>'
        '<w:cols w:space="720"/>'
        '<w:docGrid w:linePitch="360"/>'
        '</w:sectPr>'
        '</w:body>'
    )

    def test_sectpr_full_block_capture(self, tmp_path):
        """sectPr must be captured in full with all sub-elements extractable."""
        # Create minimal rels file
        word_dir = tmp_path / "word"
        rels_dir = word_dir / "_rels"
        rels_dir.mkdir(parents=True)
        (rels_dir / "document.xml.rels").write_text(
            '<?xml version="1.0" encoding="UTF-8"?>'
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            '</Relationships>',
            encoding="utf-8",
        )

        result = extract_page_layout(self.DOCUMENT_XML, tmp_path)

        assert len(result["section_chain"]) == 1
        section = result["section_chain"][0]

        # Page size must be parsed from the full sectPr block
        assert section["page_size"]["w"] == 12240
        assert section["page_size"]["h"] == 15840

        # Margins must be parsed
        assert section["page_margins"]["top"] == 1440
        assert section["page_margins"]["left"] == 1440

        # docGrid must be extracted as a child of the full sectPr
        assert section["doc_grid"] is not None
        assert "linePitch" in section["doc_grid"]
        assert_parses_as_xml(section["doc_grid"], "docGrid")

        # The raw sectPr stored in the registry must be a complete block
        assert "</w:sectPr>" in section["sectPr"]
        assert_parses_as_xml(section["sectPr"], "sectPr")

        # default_section should be set
        assert result["default_section"] is not None
