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

import base64
import xml.etree.ElementTree as ET

import pytest

from arch_env_extractor import (
    _extract_first_block,
    _extract_all_blocks,
    extract_doc_defaults,
    extract_style_defs,
    extract_latent_styles,
    extract_settings,
    extract_page_layout,
    extract_numbering,
    extract_theme,
    extract_fonts,
    extract_headers_footers,
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


# ═══════════════════════════════════════════════════════════════════════════
# Integration tests — extract_numbering
# ═══════════════════════════════════════════════════════════════════════════


class TestExtractNumbering:
    """Integration tests for extract_numbering() — abstractNum/num capture."""

    NUMBERING_XML = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
        '<w:abstractNum w:abstractNumId="0">'
        '<w:nsid w:val="AABB1122"/>'
        '<w:multiLevelType w:val="hybridMultilevel"/>'
        '<w:lvl w:ilvl="0">'
        '<w:start w:val="1"/>'
        '<w:numFmt w:val="decimal"/>'
        '<w:lvlText w:val="%1."/>'
        '<w:lvlJc w:val="left"/>'
        '<w:pPr><w:ind w:left="720" w:hanging="360"/></w:pPr>'
        '</w:lvl>'
        '</w:abstractNum>'
        '<w:abstractNum w:abstractNumId="1">'
        '<w:nsid w:val="CCDD3344"/>'
        '<w:multiLevelType w:val="singleLevel"/>'
        '<w:lvl w:ilvl="0">'
        '<w:start w:val="1"/>'
        '<w:numFmt w:val="bullet"/>'
        '<w:lvlText w:val="&#61623;"/>'
        '<w:lvlJc w:val="left"/>'
        '</w:lvl>'
        '</w:abstractNum>'
        '<w:num w:numId="1">'
        '<w:abstractNumId w:val="0"/>'
        '</w:num>'
        '<w:num w:numId="2">'
        '<w:abstractNumId w:val="1"/>'
        '</w:num>'
        '</w:numbering>'
    )

    def test_numbering_extracts_abstract_nums(self, tmp_path):
        """All abstractNum blocks are captured in full with correct IDs."""
        word_dir = tmp_path / "word"
        word_dir.mkdir()
        (word_dir / "numbering.xml").write_text(self.NUMBERING_XML, encoding="utf-8")

        result = extract_numbering(tmp_path)

        assert len(result["abstract_nums"]) == 2
        ids = {a["abstractNumId"] for a in result["abstract_nums"]}
        assert ids == {0, 1}

        for entry in result["abstract_nums"]:
            assert "</w:abstractNum>" in entry["xml"]
            assert "<w:lvl" in entry["xml"]
            assert_parses_as_xml(entry["xml"], f"abstractNum {entry['abstractNumId']}")

    def test_numbering_extracts_nums_with_abstract_ref(self, tmp_path):
        """All num blocks are captured with correct numId and abstractNumId mapping."""
        word_dir = tmp_path / "word"
        word_dir.mkdir()
        (word_dir / "numbering.xml").write_text(self.NUMBERING_XML, encoding="utf-8")

        result = extract_numbering(tmp_path)

        assert len(result["nums"]) == 2
        by_id = {n["numId"]: n for n in result["nums"]}
        assert by_id[1]["abstractNumId"] == 0
        assert by_id[2]["abstractNumId"] == 1

        for entry in result["nums"]:
            assert "</w:num>" in entry["xml"]
            assert_parses_as_xml(entry["xml"], f"num {entry['numId']}")

    def test_numbering_missing_file_returns_empty(self, tmp_path):
        """When numbering.xml is absent, result has None/empty fields."""
        word_dir = tmp_path / "word"
        word_dir.mkdir()

        result = extract_numbering(tmp_path)

        assert result["numbering_xml"] is None
        assert result["abstract_nums"] == []
        assert result["nums"] == []

    def test_numbering_full_xml_preserved(self, tmp_path):
        """The full numbering_xml is stored as a parseable string."""
        word_dir = tmp_path / "word"
        word_dir.mkdir()
        (word_dir / "numbering.xml").write_text(self.NUMBERING_XML, encoding="utf-8")

        result = extract_numbering(tmp_path)

        assert result["numbering_xml"] is not None
        assert "</w:numbering>" in result["numbering_xml"]


# ═══════════════════════════════════════════════════════════════════════════
# Integration tests — extract_latent_styles
# ═══════════════════════════════════════════════════════════════════════════


class TestExtractLatentStyles:
    """Integration tests for extract_latent_styles()."""

    STYLES_XML_WITH_LATENT = (
        '<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
        '<w:latentStyles w:defLockedState="0" w:defUIPriority="99"'
        ' w:defSemiHidden="0" w:defUnhideWhenUsed="0" w:defQFormat="0" w:count="376">'
        '<w:lsdException w:name="Normal" w:uiPriority="0" w:qFormat="1"/>'
        '<w:lsdException w:name="heading 1" w:uiPriority="9" w:qFormat="1"/>'
        '</w:latentStyles>'
        '</w:styles>'
    )

    def test_latent_styles_captures_full_block(self):
        """latentStyles block must include closing tag and all lsdException children."""
        result = extract_latent_styles(self.STYLES_XML_WITH_LATENT)

        xml = result["latentStyles_xml"]
        assert xml is not None
        assert "</w:latentStyles>" in xml
        assert "lsdException" in xml
        assert 'w:name="Normal"' in xml
        assert_parses_as_xml(xml, "latentStyles")

    def test_latent_styles_missing_returns_none(self):
        """When no latentStyles block exists, returns None."""
        styles_xml = (
            '<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
            '<w:style w:type="paragraph" w:styleId="Normal">'
            '<w:name w:val="Normal"/>'
            '</w:style>'
            '</w:styles>'
        )
        result = extract_latent_styles(styles_xml)
        assert result["latentStyles_xml"] is None


# ═══════════════════════════════════════════════════════════════════════════
# Integration tests — extract_theme
# ═══════════════════════════════════════════════════════════════════════════


class TestExtractTheme:
    """Integration tests for extract_theme()."""

    THEME_XML = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"'
        ' name="Office Theme">'
        '<a:themeElements>'
        '<a:fontScheme name="Office">'
        '<a:majorFont><a:latin typeface="Calibri Light"/></a:majorFont>'
        '<a:minorFont><a:latin typeface="Calibri"/></a:minorFont>'
        '</a:fontScheme>'
        '</a:themeElements>'
        '</a:theme>'
    )

    def test_theme_captures_full_xml(self, tmp_path):
        """Theme XML must be stored in full."""
        theme_dir = tmp_path / "word" / "theme"
        theme_dir.mkdir(parents=True)
        (theme_dir / "theme1.xml").write_text(self.THEME_XML, encoding="utf-8")

        result = extract_theme(tmp_path)

        assert result["theme1_xml"] is not None
        assert "fontScheme" in result["theme1_xml"]
        assert "majorFont" in result["theme1_xml"]
        assert "minorFont" in result["theme1_xml"]

    def test_theme_missing_returns_none(self, tmp_path):
        """When theme file is absent, theme1_xml is None."""
        (tmp_path / "word").mkdir()

        result = extract_theme(tmp_path)
        assert result["theme1_xml"] is None


# ═══════════════════════════════════════════════════════════════════════════
# Integration tests — extract_fonts
# ═══════════════════════════════════════════════════════════════════════════


class TestExtractFonts:
    """Integration tests for extract_fonts()."""

    FONT_TABLE_XML = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<w:fonts xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
        '<w:font w:name="Calibri">'
        '<w:panose1 w:val="020F0502020204030204"/>'
        '<w:charset w:val="00"/>'
        '<w:family w:val="swiss"/>'
        '<w:pitch w:val="variable"/>'
        '</w:font>'
        '</w:fonts>'
    )

    def test_fonts_captures_full_xml(self, tmp_path):
        """Font table XML must be stored in full with all child elements."""
        word_dir = tmp_path / "word"
        word_dir.mkdir()
        (word_dir / "fontTable.xml").write_text(self.FONT_TABLE_XML, encoding="utf-8")

        result = extract_fonts(tmp_path)

        assert result["font_table_xml"] is not None
        assert "</w:font>" in result["font_table_xml"]
        assert "panose1" in result["font_table_xml"]
        assert "charset" in result["font_table_xml"]

    def test_fonts_missing_returns_none(self, tmp_path):
        """When fontTable.xml is absent, font_table_xml is None."""
        (tmp_path / "word").mkdir()

        result = extract_fonts(tmp_path)
        assert result["font_table_xml"] is None


class TestExtractHeadersFooters:
    def test_extracts_media_and_rels_xml(self, tmp_path):
        word_dir = tmp_path / "word"
        rels_dir = word_dir / "_rels"
        media_dir = word_dir / "media"
        rels_dir.mkdir(parents=True)
        media_dir.mkdir(parents=True)

        (word_dir / "header1.xml").write_text(
            '<w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"/>',
            encoding="utf-8",
        )
        (rels_dir / "document.xml.rels").write_text(
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            '<Relationship Id="rId10" Target="header1.xml" />'
            '</Relationships>',
            encoding="utf-8",
        )
        (rels_dir / "header1.xml.rels").write_text(
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            '<Relationship Id="rId1" '
            'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" '
            'Target="media/image1.png" />'
            '</Relationships>',
            encoding="utf-8",
        )
        (media_dir / "image1.png").write_bytes(b"PNGDATA")

        extracted = extract_headers_footers(tmp_path)

        header = extracted["headers"][0]
        assert header["part_name"] == "word/header1.xml"
        assert header["rel_id"] == "rId10"
        assert isinstance(header["rels_xml"], str)
        assert len(header["media"]) == 1
        assert header["media"][0]["target"] == "media/image1.png"
        assert header["media"][0]["content_type"] == "image/png"
        assert base64.b64decode(header["media"][0]["data_base64"]) == b"PNGDATA"
        assert extracted["header_footer_media"] == ["media/image1.png"]
