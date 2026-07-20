"""Tests for build_arch_styles_xml_from_registry hardening."""

import xml.etree.ElementTree as ET

import pytest

from spec_formatter.style_application.core.registry import build_arch_styles_xml_from_registry


def _make_registry(*style_defs):
    """Build a minimal registry dict containing the given style_defs."""
    return {
        "styles": {"style_defs": list(style_defs)},
        "doc_defaults": {},
    }


def _sd(style_id, **overrides):
    """Shortcut to build a style_def dict."""
    d = {"style_id": style_id}
    d.update(overrides)
    return d


# ── XML well-formedness ──────────────────────────────────────────────


class TestWellFormedness:
    def test_basic_style_parses(self):
        reg = _make_registry(_sd("Heading1", name="Heading 1"))
        xml = build_arch_styles_xml_from_registry(reg)
        ET.fromstring(xml)  # must not raise

    def test_empty_style_defs_parses(self):
        reg = _make_registry()
        xml = build_arch_styles_xml_from_registry(reg)
        ET.fromstring(xml)

    def test_multiple_styles_parse(self):
        reg = _make_registry(
            _sd("Normal", name="Normal"),
            _sd("Heading1", name="Heading 1", based_on="Normal"),
        )
        xml = build_arch_styles_xml_from_registry(reg)
        ET.fromstring(xml)


# ── XML attribute escaping ───────────────────────────────────────────


class TestXmlEscaping:
    def test_ampersand_in_name(self):
        reg = _make_registry(_sd("S1", name="Foo & Bar"))
        xml = build_arch_styles_xml_from_registry(reg)
        assert "&amp;" in xml
        ET.fromstring(xml)

    def test_quotes_in_name(self):
        reg = _make_registry(_sd("S1", name='Style "Quoted"'))
        xml = build_arch_styles_xml_from_registry(reg)
        assert "&quot;" in xml
        ET.fromstring(xml)

    def test_angle_brackets_in_name(self):
        reg = _make_registry(_sd("S1", name="A < B > C"))
        xml = build_arch_styles_xml_from_registry(reg)
        assert "&lt;" in xml
        assert "&gt;" in xml
        ET.fromstring(xml)

    def test_special_chars_in_style_id(self):
        reg = _make_registry(_sd("ID&1", name="Normal"))
        xml = build_arch_styles_xml_from_registry(reg)
        assert "&amp;" in xml
        ET.fromstring(xml)

    def test_special_chars_in_based_on(self):
        reg = _make_registry(_sd("S1", based_on="Base&Style"))
        xml = build_arch_styles_xml_from_registry(reg)
        assert "&amp;" in xml
        ET.fromstring(xml)

    def test_special_chars_in_next(self):
        reg = _make_registry(_sd("S1", next="Next<Style"))
        xml = build_arch_styles_xml_from_registry(reg)
        assert "&lt;" in xml
        ET.fromstring(xml)

    def test_special_chars_in_link(self):
        reg = _make_registry(_sd("S1", link='Link"Style'))
        xml = build_arch_styles_xml_from_registry(reg)
        assert "&quot;" in xml
        ET.fromstring(xml)

    def test_apostrophe_in_name(self):
        """Apostrophe is valid in XML attributes — verify well-formedness."""
        reg = _make_registry(_sd("S1", name="Foo's Style"))
        xml = build_arch_styles_xml_from_registry(reg)
        ET.fromstring(xml)  # must not raise


# ── Name fallback ────────────────────────────────────────────────────


class TestNameFallback:
    def test_none_name_falls_back_to_style_id(self):
        reg = _make_registry(_sd("MyStyleId", name=None))
        xml = build_arch_styles_xml_from_registry(reg)
        assert 'w:val="None"' not in xml
        assert 'w:val="MyStyleId"' in xml
        ET.fromstring(xml)

    def test_missing_name_falls_back_to_style_id(self):
        reg = _make_registry(_sd("FallbackId"))
        xml = build_arch_styles_xml_from_registry(reg)
        assert 'w:val="FallbackId"' in xml
        ET.fromstring(xml)

    def test_empty_string_name_falls_back_to_style_id(self):
        reg = _make_registry(_sd("EmptyName", name=""))
        xml = build_arch_styles_xml_from_registry(reg)
        assert 'w:val="EmptyName"' in xml
        ET.fromstring(xml)

    def test_whitespace_only_name_kept_as_is(self):
        """Whitespace-only name is truthy so it is NOT replaced by style_id."""
        reg = _make_registry(_sd("WsId", name="   "))
        xml = build_arch_styles_xml_from_registry(reg)
        # The name value should be the whitespace, not the style_id
        assert 'w:val="WsId"' not in xml or 'w:styleId="WsId"' in xml
        ET.fromstring(xml)  # still well-formed


# ── Raw XML fragments preserved verbatim ─────────────────────────────


class TestRawFragments:
    def test_ppr_not_escaped(self):
        ppr = '<w:pPr><w:jc w:val="center"/></w:pPr>'
        reg = _make_registry(_sd("S1", pPr=ppr))
        xml = build_arch_styles_xml_from_registry(reg)
        assert ppr in xml
        ET.fromstring(xml)

    def test_rpr_not_escaped(self):
        rpr = '<w:rPr><w:b/><w:sz w:val="24"/></w:rPr>'
        reg = _make_registry(_sd("S1", rPr=rpr))
        xml = build_arch_styles_xml_from_registry(reg)
        assert rpr in xml
        ET.fromstring(xml)

    def test_tblpr_not_escaped(self):
        tbl = '<w:tblPr><w:tblW w:w="5000" w:type="pct"/></w:tblPr>'
        reg = _make_registry(_sd("S1", tblPr=tbl))
        xml = build_arch_styles_xml_from_registry(reg)
        assert tbl in xml
        ET.fromstring(xml)


# ── Validation catches bad XML fragments ─────────────────────────────


class TestValidation:
    def test_malformed_ppr_raises(self):
        bad_ppr = "<w:pPr><w:jc>"  # unclosed tag
        reg = _make_registry(_sd("S1", pPr=bad_ppr))
        with pytest.raises(ValueError, match="well-formedness"):
            build_arch_styles_xml_from_registry(reg)


# ── Skips empty style_id ─────────────────────────────────────────────


class TestDependencyRefs:
    """Verify basedOn, link, and next references appear in output XML."""

    def test_all_three_refs_present(self):
        reg = _make_registry(
            _sd("Normal"),
            _sd("Heading1", based_on="Normal", link="Heading1Char", next="BodyText"),
            _sd("Heading1Char", type="character"),
            _sd("BodyText", based_on="Normal"),
        )
        xml = build_arch_styles_xml_from_registry(reg)
        assert '<w:basedOn w:val="Normal"/>' in xml
        assert '<w:link w:val="Heading1Char"/>' in xml
        assert '<w:next w:val="BodyText"/>' in xml
        ET.fromstring(xml)

    def test_based_on_chain_all_present(self):
        """A -> B -> C chain: all three styles appear."""
        reg = _make_registry(
            _sd("StyleC"),
            _sd("StyleB", based_on="StyleC"),
            _sd("StyleA", based_on="StyleB"),
        )
        xml = build_arch_styles_xml_from_registry(reg)
        assert 'w:styleId="StyleA"' in xml
        assert 'w:styleId="StyleB"' in xml
        assert 'w:styleId="StyleC"' in xml
        ET.fromstring(xml)


class TestEdgeCases:
    def test_empty_style_id_skipped(self):
        reg = _make_registry(_sd(""), _sd("Valid"))
        xml = build_arch_styles_xml_from_registry(reg)
        assert 'w:styleId=""' not in xml
        assert 'w:styleId="Valid"' in xml
        ET.fromstring(xml)

    def test_character_type(self):
        reg = _make_registry(_sd("CharStyle", type="character", name="Char"))
        xml = build_arch_styles_xml_from_registry(reg)
        assert 'w:type="character"' in xml
        ET.fromstring(xml)

    def test_table_type(self):
        reg = _make_registry(_sd("TblStyle", type="table", name="Table"))
        xml = build_arch_styles_xml_from_registry(reg)
        assert 'w:type="table"' in xml
        ET.fromstring(xml)

    def test_ui_priority_rendered(self):
        reg = _make_registry(_sd("S1", ui_priority=99))
        xml = build_arch_styles_xml_from_registry(reg)
        assert '<w:uiPriority w:val="99"/>' in xml
        ET.fromstring(xml)
