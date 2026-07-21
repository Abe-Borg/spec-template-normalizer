"""Tests for core.xml_helpers — paragraph-level XML manipulation."""

import pytest
from spec_formatter.style_application.core.xml_helpers import (
    apply_pstyle_to_paragraph_block,
    strip_direct_run_properties,
    strip_run_font_formatting,
    iter_paragraph_xml_blocks,
    paragraph_text_from_block,
    paragraph_contains_sectpr,
    paragraph_pstyle_from_block,
    paragraph_numpr_from_block,
    strip_conflicting_direct_ppr,
)


# ── apply_pstyle_to_paragraph_block ──────────────────────────────────────────

class TestApplyPstyle:
    def test_replace_existing_pstyle(self):
        p = '<w:p><w:pPr><w:pStyle w:val="OldStyle"/></w:pPr><w:r><w:t>Hello</w:t></w:r></w:p>'
        result = apply_pstyle_to_paragraph_block(p, "NewStyle")
        assert 'w:val="NewStyle"' in result
        assert 'w:val="OldStyle"' not in result

    def test_self_closing_ppr(self):
        p = '<w:p><w:pPr/><w:r><w:t>Hello</w:t></w:r></w:p>'
        result = apply_pstyle_to_paragraph_block(p, "MyStyle")
        assert '<w:pPr><w:pStyle w:val="MyStyle"/></w:pPr>' in result

    def test_open_ppr_no_pstyle(self):
        p = '<w:p><w:pPr><w:jc w:val="center"/></w:pPr><w:r><w:t>Hello</w:t></w:r></w:p>'
        result = apply_pstyle_to_paragraph_block(p, "MyStyle")
        assert '<w:pStyle w:val="MyStyle"/>' in result
        assert '<w:jc w:val="center"/>' in result

    def test_no_ppr_at_all(self):
        p = '<w:p><w:r><w:t>Hello</w:t></w:r></w:p>'
        result = apply_pstyle_to_paragraph_block(p, "MyStyle")
        assert '<w:pPr><w:pStyle w:val="MyStyle"/></w:pPr>' in result

    def test_historical_pstyle_is_preserved_and_new_live_style_is_inserted(self):
        change = (
            '<w:pPrChange w:id="7"><w:pPr>'
            '<w:pStyle w:val="HistoricalList"/>'
            '</w:pPr></w:pPrChange>'
        )
        p = f'<w:p><w:pPr>{change}</w:pPr><w:r><w:t>Hello</w:t></w:r></w:p>'

        result = apply_pstyle_to_paragraph_block(p, "MyStyle")

        assert change in result
        assert result.startswith(
            '<w:p><w:pPr><w:pStyle w:val="MyStyle"/>'
            '<w:pPrChange'
        )

    def test_visible_sectpr_paragraph_is_styled_without_changing_sectpr(self):
        sectpr = '<w:sectPr><w:pgSz/></w:sectPr>'
        p = f'<w:p><w:pPr>{sectpr}</w:pPr><w:r><w:t>Visible</w:t></w:r></w:p>'
        result = apply_pstyle_to_paragraph_block(p, "MyStyle")
        assert '<w:pStyle w:val="MyStyle"/>' in result
        assert sectpr in result


# ── strip_run_font_formatting ────────────────────────────────────────────────

class TestStripRunFontFormatting:
    def test_strip_rfonts_sz_szcs(self):
        p = (
            '<w:p><w:r><w:rPr>'
            '<w:rFonts w:ascii="Arial" w:hAnsi="Arial"/>'
            '<w:sz w:val="20"/>'
            '<w:szCs w:val="20"/>'
            '</w:rPr><w:t>Hello</w:t></w:r></w:p>'
        )
        result = strip_run_font_formatting(p)
        assert '<w:rFonts' not in result
        assert '<w:sz' not in result
        assert '<w:szCs' not in result
        assert '<w:t>Hello</w:t>' in result

    def test_preserve_bold(self):
        p = (
            '<w:p><w:r><w:rPr>'
            '<w:rFonts w:ascii="Arial"/>'
            '<w:b/>'
            '</w:rPr><w:t>Bold</w:t></w:r></w:p>'
        )
        result = strip_run_font_formatting(p)
        assert '<w:rFonts' not in result
        assert '<w:b/>' in result

    def test_no_font_formatting_unchanged(self):
        p = '<w:p><w:r><w:rPr><w:b/><w:i/></w:rPr><w:t>BI</w:t></w:r></w:p>'
        result = strip_run_font_formatting(p)
        assert result == p

    def test_selected_properties_include_language_without_touching_size(self):
        p = (
            '<w:p><w:r><w:rPr>'
            '<w:rFonts w:ascii="Target"/><w:sz w:val="20"/>'
            '<w:szCs w:val="20"/><w:lang w:val="fr-CA"/><w:b/>'
            '</w:rPr><w:t>Text</w:t></w:r></w:p>'
        )

        result = strip_direct_run_properties(p, {"rFonts", "lang"})

        assert "<w:rFonts" not in result
        assert "<w:lang" not in result
        assert '<w:sz w:val="20"/>' in result
        assert '<w:szCs w:val="20"/>' in result
        assert "<w:b/>" in result

    def test_historical_run_properties_are_byte_stable(self):
        historical = (
            '<w:rPrChange w:id="4"><w:rPr><w:rFonts w:ascii="Historical"/>'
            '<w:lang w:val="de-DE"/></w:rPr></w:rPrChange>'
        )
        p = (
            '<w:p><w:r><w:rPr><w:rFonts w:ascii="Live"/>'
            f'<w:lang w:val="fr-CA"/>{historical}</w:rPr><w:t>Text</w:t></w:r></w:p>'
        )

        result = strip_direct_run_properties(p, {"rFonts", "lang"})

        assert historical in result
        live_prefix = result.split("<w:rPrChange", 1)[0]
        assert "<w:rFonts" not in live_prefix
        assert "<w:lang" not in live_prefix

    def test_selected_property_removal_does_not_touch_nested_descendant(self):
        nested = '<w14:textFill><w:color w:val="Nested"/></w14:textFill>'
        p = (
            '<w:p><w:r><w:rPr><w:color w:val="Direct"/>'
            f'{nested}</w:rPr><w:t>Text</w:t></w:r></w:p>'
        )

        result = strip_direct_run_properties(p, {"color"})

        assert '<w:color w:val="Direct"/>' not in result
        assert nested in result


class TestStripConflictingDirectPpr:
    def test_removes_jc_ind_spacing_and_numpr(self):
        p = (
            '<w:p><w:pPr>'
            '<w:numPr><w:numId w:val="4"/><w:ilvl w:val="1"/></w:numPr>'
            '<w:jc w:val="center"/><w:ind w:left="720"/><w:spacing w:before="120"/>'
            '</w:pPr><w:r><w:t>X</w:t></w:r></w:p>'
        )
        result = strip_conflicting_direct_ppr(p)
        assert '<w:jc' not in result
        assert '<w:ind' not in result
        assert '<w:spacing' not in result
        assert '<w:numPr' not in result

    def test_sectpr_preserved_while_conflicting_property_is_removed(self):
        p = '<w:p><w:pPr><w:sectPr/><w:jc w:val="center"/></w:pPr></w:p>'
        result = strip_conflicting_direct_ppr(p)
        assert '<w:sectPr/>' in result
        assert '<w:jc' not in result

    def test_multiple_runs(self):
        p = (
            '<w:p>'
            '<w:r><w:rPr><w:rFonts w:ascii="Arial"/><w:sz w:val="20"/></w:rPr><w:t>A</w:t></w:r>'
            '<w:r><w:rPr><w:rFonts w:ascii="Times"/><w:b/></w:rPr><w:t>B</w:t></w:r>'
            '</w:p>'
        )
        result = strip_run_font_formatting(p)
        assert '<w:rFonts' not in result
        assert '<w:sz' not in result
        assert '<w:b/>' in result
        assert '<w:t>A</w:t>' in result
        assert '<w:t>B</w:t>' in result

    def test_sectpr_preserved_while_visible_run_font_is_stripped(self):
        p = '<w:p><w:pPr><w:sectPr/></w:pPr><w:r><w:rPr><w:rFonts w:ascii="Arial"/></w:rPr><w:t>X</w:t></w:r></w:p>'
        result = strip_run_font_formatting(p)
        assert '<w:sectPr/>' in result
        assert '<w:rFonts' not in result
        assert '<w:t>X</w:t>' in result


# ── iter_paragraph_xml_blocks ────────────────────────────────────────────────

class TestIterParagraphXmlBlocks:
    def test_basic_document(self):
        doc = (
            '<?xml version="1.0"?>'
            '<w:document><w:body>'
            '<w:p><w:r><w:t>First</w:t></w:r></w:p>'
            '<w:p><w:r><w:t>Second</w:t></w:r></w:p>'
            '<w:p><w:r><w:t>Third</w:t></w:r></w:p>'
            '</w:body></w:document>'
        )
        blocks = list(iter_paragraph_xml_blocks(doc))
        assert len(blocks) == 3

    def test_correct_positions(self):
        doc = '<w:body><w:p><w:t>A</w:t></w:p></w:body>'
        blocks = list(iter_paragraph_xml_blocks(doc))
        assert len(blocks) == 1
        start, end, p_xml = blocks[0]
        assert doc[start:end] == p_xml
        assert '<w:t>A</w:t>' in p_xml

    def test_paragraph_count(self):
        paras = ''.join(f'<w:p><w:r><w:t>P{i}</w:t></w:r></w:p>' for i in range(10))
        doc = f'<w:body>{paras}</w:body>'
        blocks = list(iter_paragraph_xml_blocks(doc))
        assert len(blocks) == 10

    def test_self_closing_paragraph(self):
        doc = '<w:body><w:p/><w:p><w:r><w:t>A</w:t></w:r></w:p></w:body>'
        blocks = list(iter_paragraph_xml_blocks(doc))
        assert [block for _start, _end, block in blocks] == [
            '<w:p/>',
            '<w:p><w:r><w:t>A</w:t></w:r></w:p>',
        ]

    def test_self_closing_paragraph_scanner_is_quote_aware(self):
        doc = '<w:body><w:p data-note=">"/></w:body>'

        blocks = list(iter_paragraph_xml_blocks(doc))

        assert len(blocks) == 1
        start, end, block = blocks[0]
        assert block == '<w:p data-note=">"/>'
        assert doc[start:end] == block

    def test_nested_textbox_paragraph_does_not_truncate_host(self):
        host = (
            '<w:p><w:r><w:t>Host</w:t><w:drawing><w:txbxContent>'
            '<w:p><w:r><w:t>Nested</w:t></w:r></w:p>'
            '</w:txbxContent></w:drawing></w:r></w:p>'
        )
        doc = f'<w:body>{host}<w:p><w:r><w:t>After</w:t></w:r></w:p></w:body>'
        blocks = list(iter_paragraph_xml_blocks(doc))
        assert len(blocks) == 2
        assert blocks[0][2] == host
        assert paragraph_text_from_block(blocks[0][2]) == "Host"
        assert paragraph_text_from_block(blocks[1][2]) == "After"


# ── paragraph_text_from_block ────────────────────────────────────────────────

class TestParagraphTextFromBlock:
    def test_basic_text(self):
        p = '<w:p><w:r><w:t>Hello World</w:t></w:r></w:p>'
        assert paragraph_text_from_block(p) == "Hello World"

    def test_multiple_runs(self):
        p = '<w:p><w:r><w:t>Hello </w:t></w:r><w:r><w:t>World</w:t></w:r></w:p>'
        assert paragraph_text_from_block(p) == "Hello World"

    def test_empty_paragraph(self):
        p = '<w:p><w:pPr><w:pStyle w:val="Normal"/></w:pPr></w:p>'
        assert paragraph_text_from_block(p) == ""

    def test_html_entities(self):
        p = '<w:p><w:r><w:t>A &amp; B</w:t></w:r></w:p>'
        assert paragraph_text_from_block(p) == "A & B"

    def test_visible_text_controls_revisions_and_special_hyphens(self):
        p = (
            '<w:p><w:r><w:t>SECTION</w:t><w:tab/><w:t>21</w:t>'
            '<w:br/><w:t>13</w:t><w:cr/><w:t>13</w:t>'
            '<w:noBreakHyphen/><w:t>A</w:t><w:softHyphen/><w:t>B</w:t></w:r>'
            '<w:del><w:r><w:t>DELETED</w:t></w:r></w:del>'
            '<w:moveFrom><w:r><w:t>MOVED</w:t></w:r></w:moveFrom>'
            '<w:moveTo><w:r><w:t> Kept</w:t></w:r></w:moveTo></w:p>'
        )
        assert paragraph_text_from_block(p) == "SECTION 21 13 13\u2011AB Kept"

    def test_paired_empty_visible_controls_match_self_closing_forms(self):
        p = (
            '<w:p><w:r><w:t>A</w:t><w:tab></w:tab><w:t>B</w:t>'
            '<w:br></w:br><w:t>C</w:t><w:noBreakHyphen></w:noBreakHyphen>'
            '<w:t>D</w:t><w:softHyphen></w:softHyphen><w:t>E</w:t></w:r></w:p>'
        )

        assert paragraph_text_from_block(p) == "A B C\u2011DE"

    @pytest.mark.parametrize(
        "textbox",
        [
            '<w:txbxContent><w:p><w:r><w:t>Hidden</w:t></w:r></w:p></w:txbxContent>',
            '<v:textbox><w:txbxContent><w:p><w:r><w:t>Hidden</w:t></w:r></w:p></w:txbxContent></v:textbox>',
            '<wps:txbx><w:txbxContent><w:p><w:r><w:t>Hidden</w:t></w:r></w:p></w:txbxContent></wps:txbx>',
        ],
    )
    def test_bare_and_legacy_textbox_subtrees_are_excluded(self, textbox):
        paragraph = f'<w:p><w:r><w:t>Host</w:t>{textbox}</w:r></w:p>'
        assert paragraph_text_from_block(paragraph) == "Host"


# ── paragraph_contains_sectpr ────────────────────────────────────────────────

class TestParagraphContainsSectpr:
    def test_with_sectpr(self):
        p = '<w:p><w:pPr><w:sectPr><w:pgSz/></w:sectPr></w:pPr></w:p>'
        assert paragraph_contains_sectpr(p) is True

    def test_without_sectpr(self):
        p = '<w:p><w:r><w:t>Normal</w:t></w:r></w:p>'
        assert paragraph_contains_sectpr(p) is False

    def test_sectpr_text_inside_comment_is_not_structural(self):
        p = '<w:p><!-- <w:sectPr/> --><w:r><w:t>Normal</w:t></w:r></w:p>'
        assert paragraph_contains_sectpr(p) is False


# ── paragraph_pstyle_from_block ──────────────────────────────────────────────

class TestParagraphPstyleFromBlock:
    def test_with_pstyle(self):
        p = '<w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:r><w:t>H</w:t></w:r></w:p>'
        assert paragraph_pstyle_from_block(p) == "Heading1"

    def test_without_pstyle(self):
        p = '<w:p><w:r><w:t>Normal</w:t></w:r></w:p>'
        assert paragraph_pstyle_from_block(p) is None

    def test_tracked_previous_pstyle_is_not_live(self):
        p = (
            '<w:p><w:pPr><w:pPrChange w:id="7"><w:pPr>'
            '<w:pStyle w:val="HistoricalList"/>'
            '</w:pPr></w:pPrChange></w:pPr><w:r><w:t>Normal</w:t></w:r></w:p>'
        )
        assert paragraph_pstyle_from_block(p) is None

    def test_live_pstyle_wins_over_tracked_previous_pstyle(self):
        p = (
            '<w:p><w:pPr><w:pStyle w:val="CurrentBody"/>'
            '<w:pPrChange w:id="7"><w:pPr>'
            '<w:pStyle w:val="HistoricalList"/>'
            '</w:pPr></w:pPrChange></w:pPr></w:p>'
        )
        assert paragraph_pstyle_from_block(p) == "CurrentBody"


# ── paragraph_numpr_from_block ───────────────────────────────────────────────

class TestParagraphNumprFromBlock:
    def test_with_numpr(self):
        p = '<w:p><w:pPr><w:numPr><w:numId w:val="5"/><w:ilvl w:val="2"/></w:numPr></w:pPr></w:p>'
        result = paragraph_numpr_from_block(p)
        assert result["numId"] == "5"
        assert result["ilvl"] == "2"

    def test_without_numpr(self):
        p = '<w:p><w:r><w:t>No list</w:t></w:r></w:p>'
        result = paragraph_numpr_from_block(p)
        assert result["numId"] is None
        assert result["ilvl"] is None

    def test_tracked_previous_numpr_is_not_live(self):
        p = (
            '<w:p><w:pPr><w:pPrChange w:id="7"><w:pPr><w:numPr>'
            '<w:ilvl w:val="4"/><w:numId w:val="91"/>'
            '</w:numPr></w:pPr></w:pPrChange></w:pPr></w:p>'
        )
        assert paragraph_numpr_from_block(p) == {"numId": None, "ilvl": None}

    def test_live_numpr_wins_over_tracked_previous_numpr(self):
        p = (
            '<w:p><w:pPr><w:numPr><w:ilvl w:val="1"/>'
            '<w:numId w:val="5"/></w:numPr><w:pPrChange w:id="7"><w:pPr>'
            '<w:numPr><w:ilvl w:val="4"/><w:numId w:val="91"/></w:numPr>'
            '</w:pPr></w:pPrChange></w:pPr></w:p>'
        )
        assert paragraph_numpr_from_block(p) == {"numId": "5", "ilvl": "1"}
