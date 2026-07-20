"""Tests for core.style_import — style extraction and materialization."""

import pytest
from spec_formatter.style_application.core.style_import import (
    ensure_explicit_numpr_from_current_style,
    materialize_arch_style_block,
    _extract_style_block,
    _extract_basedOn,
    _find_style_numpr_in_chain,
    _collect_style_deps_from_arch,
)


# ── Sample styles.xml fragments ─────────────────────────────────────────────

STYLES_WITH_NUMPR = '''
<w:styles>
  <w:style w:type="paragraph" w:styleId="ListBullet">
    <w:name w:val="List Bullet"/>
    <w:pPr>
      <w:numPr><w:numId w:val="1"/><w:ilvl w:val="0"/></w:numPr>
    </w:pPr>
  </w:style>
  <w:style w:type="paragraph" w:styleId="ListBullet2">
    <w:name w:val="List Bullet 2"/>
    <w:basedOn w:val="ListBullet"/>
  </w:style>
  <w:style w:type="paragraph" w:styleId="ListBullet3">
    <w:name w:val="List Bullet 3"/>
    <w:basedOn w:val="ListBullet2"/>
  </w:style>
  <w:style w:type="paragraph" w:styleId="Normal">
    <w:name w:val="Normal"/>
  </w:style>
</w:styles>
'''

STYLES_FOR_MATERIALIZE = '''
<w:docDefaults>
  <w:rPrDefault>
    <w:rPr>
      <w:rFonts w:ascii="Calibri" w:hAnsi="Calibri"/>
      <w:sz w:val="22"/>
      <w:szCs w:val="22"/>
      <w:lang w:val="en-US"/>
    </w:rPr>
  </w:rPrDefault>
  <w:pPrDefault>
    <w:pPr>
      <w:spacing w:after="160"/>
    </w:pPr>
  </w:pPrDefault>
</w:docDefaults>
<w:styles>
  <w:style w:type="paragraph" w:styleId="CSI-Part">
    <w:name w:val="CSI-Part"/>
    <w:rPr>
      <w:b/>
      <w:rFonts w:ascii="Arial" w:hAnsi="Arial"/>
    </w:rPr>
  </w:style>
  <w:style w:type="paragraph" w:styleId="CSI-Article">
    <w:name w:val="CSI-Article"/>
    <w:basedOn w:val="CSI-Part"/>
  </w:style>
  <w:style w:type="paragraph" w:styleId="NoRpr">
    <w:name w:val="NoRpr"/>
  </w:style>
  <w:style w:type="paragraph" w:styleId="CompleteRpr">
    <w:name w:val="CompleteRpr"/>
    <w:rPr>
      <w:rFonts w:ascii="Times" w:hAnsi="Times"/>
      <w:sz w:val="24"/>
      <w:szCs w:val="24"/>
      <w:lang w:val="en-GB"/>
    </w:rPr>
  </w:style>
</w:styles>
'''


# ── ensure_explicit_numpr_from_current_style ─────────────────────────────────

class TestEnsureExplicitNumpr:
    def test_already_has_numpr(self):
        p = '<w:p><w:pPr><w:pStyle w:val="ListBullet"/><w:numPr><w:numId w:val="1"/></w:numPr></w:pPr><w:r><w:t>A</w:t></w:r></w:p>'
        result = ensure_explicit_numpr_from_current_style(p, STYLES_WITH_NUMPR)
        assert result == p  # Unchanged

    def test_no_pstyle(self):
        p = '<w:p><w:r><w:t>Plain</w:t></w:r></w:p>'
        result = ensure_explicit_numpr_from_current_style(p, STYLES_WITH_NUMPR)
        assert result == p  # Unchanged

    def test_style_has_numpr_depth_1(self):
        p = '<w:p><w:pPr><w:pStyle w:val="ListBullet"/></w:pPr><w:r><w:t>Item</w:t></w:r></w:p>'
        result = ensure_explicit_numpr_from_current_style(p, STYLES_WITH_NUMPR)
        assert '<w:numPr>' in result
        assert '<w:numId w:val="1"/>' in result

    def test_style_has_numpr_depth_2(self):
        """ListBullet2 -> basedOn ListBullet which has numPr."""
        p = '<w:p><w:pPr><w:pStyle w:val="ListBullet2"/></w:pPr><w:r><w:t>Item</w:t></w:r></w:p>'
        result = ensure_explicit_numpr_from_current_style(p, STYLES_WITH_NUMPR)
        assert '<w:numPr>' in result

    def test_style_has_numpr_depth_3(self):
        """ListBullet3 -> ListBullet2 -> ListBullet which has numPr."""
        p = '<w:p><w:pPr><w:pStyle w:val="ListBullet3"/></w:pPr><w:r><w:t>Item</w:t></w:r></w:p>'
        result = ensure_explicit_numpr_from_current_style(p, STYLES_WITH_NUMPR)
        assert '<w:numPr>' in result

    def test_style_no_numpr(self):
        p = '<w:p><w:pPr><w:pStyle w:val="Normal"/></w:pPr><w:r><w:t>Text</w:t></w:r></w:p>'
        result = ensure_explicit_numpr_from_current_style(p, STYLES_WITH_NUMPR)
        assert '<w:numPr>' not in result

    def test_sectpr_unchanged(self):
        p = '<w:p><w:pPr><w:pStyle w:val="ListBullet"/><w:sectPr/></w:pPr></w:p>'
        result = ensure_explicit_numpr_from_current_style(p, STYLES_WITH_NUMPR)
        assert result == p


# ── materialize_arch_style_block ─────────────────────────────────────────────

class TestMaterializeArchStyleBlock:
    def test_no_rpr_gets_effective_rpr(self):
        """Style with no rPr should get effective rPr injected from docDefaults."""
        style = '<w:style w:type="paragraph" w:styleId="NoRpr"><w:name w:val="NoRpr"/></w:style>'
        result = materialize_arch_style_block(style, "NoRpr", STYLES_FOR_MATERIALIZE)
        assert '<w:rPr>' in result
        assert '<w:rFonts' in result
        assert '<w:sz' in result

    def test_partial_rpr_gets_missing_filled(self):
        """CSI-Part has rFonts but no sz/szCs/lang — those should be filled from docDefaults."""
        style = _extract_style_block(STYLES_FOR_MATERIALIZE, "CSI-Part")
        assert style is not None
        result = materialize_arch_style_block(style, "CSI-Part", STYLES_FOR_MATERIALIZE)
        assert '<w:sz' in result
        assert '<w:szCs' in result
        assert '<w:lang' in result
        # Original rFonts should still be there
        assert 'w:ascii="Arial"' in result

    def test_complete_rpr_unchanged(self):
        """CompleteRpr already has all FORCE tags — should not be modified."""
        style = _extract_style_block(STYLES_FOR_MATERIALIZE, "CompleteRpr")
        assert style is not None
        result = materialize_arch_style_block(style, "CompleteRpr", STYLES_FOR_MATERIALIZE)
        # All original values preserved
        assert 'w:ascii="Times"' in result
        assert 'w:val="24"' in result
        assert 'w:val="en-GB"' in result


# ── _collect_style_deps_from_arch ──────────────────────────────────────────

STYLES_WITH_LINK_NEXT = '''
<w:styles>
  <w:style w:type="paragraph" w:styleId="Heading1">
    <w:name w:val="Heading 1"/>
    <w:basedOn w:val="Normal"/>
    <w:link w:val="Heading1Char"/>
    <w:next w:val="BodyText"/>
  </w:style>
  <w:style w:type="character" w:styleId="Heading1Char">
    <w:name w:val="Heading 1 Char"/>
    <w:link w:val="Heading1"/>
  </w:style>
  <w:style w:type="paragraph" w:styleId="Normal">
    <w:name w:val="Normal"/>
  </w:style>
  <w:style w:type="paragraph" w:styleId="BodyText">
    <w:name w:val="Body Text"/>
    <w:basedOn w:val="Normal"/>
  </w:style>
</w:styles>
'''


class TestCollectStyleDeps:
    def test_follows_basedOn(self):
        seen = set()
        _collect_style_deps_from_arch(STYLES_WITH_LINK_NEXT, "Heading1", seen)
        assert "Normal" in seen

    def test_follows_link(self):
        seen = set()
        _collect_style_deps_from_arch(STYLES_WITH_LINK_NEXT, "Heading1", seen)
        assert "Heading1Char" in seen

    def test_follows_next(self):
        seen = set()
        _collect_style_deps_from_arch(STYLES_WITH_LINK_NEXT, "Heading1", seen)
        assert "BodyText" in seen

    def test_transitive_deps(self):
        """BodyText basedOn Normal — reached transitively via next."""
        seen = set()
        _collect_style_deps_from_arch(STYLES_WITH_LINK_NEXT, "Heading1", seen)
        assert seen == {"Heading1", "Normal", "Heading1Char", "BodyText"}

    def test_cycle_protection(self):
        """Heading1 <-> Heading1Char via mutual link — must not loop."""
        seen = set()
        _collect_style_deps_from_arch(STYLES_WITH_LINK_NEXT, "Heading1", seen)
        # Completing without infinite recursion is the test
        assert "Heading1" in seen
        assert "Heading1Char" in seen

    def test_missing_target_skipped(self):
        styles = '''
<w:styles>
  <w:style w:type="paragraph" w:styleId="Orphan">
    <w:name w:val="Orphan"/>
    <w:next w:val="DoesNotExist"/>
  </w:style>
</w:styles>
'''
        seen = set()
        _collect_style_deps_from_arch(styles, "Orphan", seen)
        # DoesNotExist is visited (added to seen) but has no block — no crash
        assert "Orphan" in seen
        assert "DoesNotExist" in seen

    def test_only_based_on(self):
        """Style with only basedOn — only basedOn target collected."""
        styles = '''
<w:styles>
  <w:style w:type="paragraph" w:styleId="Child">
    <w:name w:val="Child"/>
    <w:basedOn w:val="Parent"/>
  </w:style>
  <w:style w:type="paragraph" w:styleId="Parent">
    <w:name w:val="Parent"/>
  </w:style>
</w:styles>
'''
        seen = set()
        _collect_style_deps_from_arch(styles, "Child", seen)
        assert seen == {"Child", "Parent"}

    def test_only_next(self):
        """Style with only next — only next target collected."""
        styles = '''
<w:styles>
  <w:style w:type="paragraph" w:styleId="StyleA">
    <w:name w:val="StyleA"/>
    <w:next w:val="StyleB"/>
  </w:style>
  <w:style w:type="paragraph" w:styleId="StyleB">
    <w:name w:val="StyleB"/>
  </w:style>
</w:styles>
'''
        seen = set()
        _collect_style_deps_from_arch(styles, "StyleA", seen)
        assert seen == {"StyleA", "StyleB"}

    def test_deep_based_on_chain(self):
        """A -> B -> C -> D: all 4 collected via transitive basedOn."""
        styles = '''
<w:styles>
  <w:style w:type="paragraph" w:styleId="A">
    <w:name w:val="A"/>
    <w:basedOn w:val="B"/>
  </w:style>
  <w:style w:type="paragraph" w:styleId="B">
    <w:name w:val="B"/>
    <w:basedOn w:val="C"/>
  </w:style>
  <w:style w:type="paragraph" w:styleId="C">
    <w:name w:val="C"/>
    <w:basedOn w:val="D"/>
  </w:style>
  <w:style w:type="paragraph" w:styleId="D">
    <w:name w:val="D"/>
  </w:style>
</w:styles>
'''
        seen = set()
        _collect_style_deps_from_arch(styles, "A", seen)
        assert seen == {"A", "B", "C", "D"}

    def test_nonexistent_start_style(self):
        """Starting from a nonexistent style — seen contains only that ID."""
        styles = '<w:styles></w:styles>'
        seen = set()
        _collect_style_deps_from_arch(styles, "Ghost", seen)
        assert seen == {"Ghost"}


# ── Additional materialize_arch_style_block tests ──────────────────────────

class TestMaterializeArchStyleBlockExtended:
    def test_derived_style_gets_parent_rfonts(self):
        """CSI-Article (basedOn CSI-Part) should get Arial from parent chain."""
        style = _extract_style_block(STYLES_FOR_MATERIALIZE, "CSI-Article")
        assert style is not None
        result = materialize_arch_style_block(style, "CSI-Article", STYLES_FOR_MATERIALIZE)
        # Should have rPr with fonts from CSI-Part (Arial)
        assert '<w:rPr>' in result
        assert 'w:ascii="Arial"' in result
        # Should also get sz/szCs from docDefaults
        assert '<w:sz' in result
        assert '<w:szCs' in result

    def test_paragraph_no_ppr_gets_effective_ppr(self):
        """Paragraph style with no pPr gets effective pPr from chain."""
        # CSI-Article has no pPr and basedOn CSI-Part which also has no pPr,
        # so it should get pPr from docDefaults (spacing after="160")
        style = _extract_style_block(STYLES_FOR_MATERIALIZE, "CSI-Article")
        assert style is not None
        result = materialize_arch_style_block(style, "CSI-Article", STYLES_FOR_MATERIALIZE)
        assert '<w:pPr>' in result
        assert '<w:spacing' in result


# ── TestBuiltinStyleSkipping ───────────────────────────────────────────────

class TestBuiltinStyleSkipping:
    """
    Tests that Word built-in styles (Normal, DefaultParagraphFont, etc.)
    are silently skipped during import when absent from both architect
    and target, rather than raising ValueError.
    """

    ARCH_STYLES_XML_NORMAL_MISSING = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
        '<w:docDefaults>'
        '<w:rPrDefault><w:rPr><w:sz w:val="24"/></w:rPr></w:rPrDefault>'
        '<w:pPrDefault><w:pPr/></w:pPrDefault>'
        '</w:docDefaults>'
        '<w:style w:type="paragraph" w:styleId="CSI_Part__ARCH">'
        '<w:name w:val="CSI Part"/>'
        '<w:basedOn w:val="Normal"/>'
        '<w:pPr><w:spacing w:before="240"/></w:pPr>'
        '<w:rPr><w:b/></w:rPr>'
        '</w:style>'
        '</w:styles>'
    )

    TARGET_STYLES_XML_NO_NORMAL = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
        '<w:docDefaults>'
        '<w:rPrDefault><w:rPr><w:sz w:val="22"/></w:rPr></w:rPrDefault>'
        '<w:pPrDefault><w:pPr/></w:pPrDefault>'
        '</w:docDefaults>'
        '</w:styles>'
    )

    def test_normal_dependency_skipped_when_missing_from_both(self, tmp_path):
        """
        When a CSI style has basedOn="Normal" but neither architect nor target
        explicitly define "Normal", import should succeed.
        """
        from spec_formatter.style_application.core.style_import import import_arch_styles_into_target

        word_dir = tmp_path / "word"
        word_dir.mkdir()
        (word_dir / "styles.xml").write_text(
            self.TARGET_STYLES_XML_NO_NORMAL, encoding="utf-8"
        )

        log = []
        import_arch_styles_into_target(
            target_extract_dir=tmp_path,
            arch_styles_xml=self.ARCH_STYLES_XML_NORMAL_MISSING,
            needed_style_ids=["CSI_Part__ARCH"],
            log=log,
        )

        result_xml = (word_dir / "styles.xml").read_text(encoding="utf-8")
        assert 'w:styleId="CSI_Part__ARCH"' in result_xml
        assert any("Normal" in msg and "built-in" in msg.lower() for msg in log), \
            f"Expected log entry about skipping built-in Normal, got: {log}"

    def test_normal_used_from_target_when_explicit(self, tmp_path):
        """
        When the target explicitly defines "Normal", the existing check
        catches it and no built-in skip is needed.
        """
        from spec_formatter.style_application.core.style_import import import_arch_styles_into_target

        target_xml = (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
            '<w:docDefaults>'
            '<w:rPrDefault><w:rPr><w:sz w:val="22"/></w:rPr></w:rPrDefault>'
            '<w:pPrDefault><w:pPr/></w:pPrDefault>'
            '</w:docDefaults>'
            '<w:style w:type="paragraph" w:styleId="Normal">'
            '<w:name w:val="Normal"/>'
            '<w:pPr><w:spacing w:after="200"/></w:pPr>'
            '</w:style>'
            '</w:styles>'
        )

        word_dir = tmp_path / "word"
        word_dir.mkdir()
        (word_dir / "styles.xml").write_text(target_xml, encoding="utf-8")

        log = []
        import_arch_styles_into_target(
            target_extract_dir=tmp_path,
            arch_styles_xml=self.ARCH_STYLES_XML_NORMAL_MISSING,
            needed_style_ids=["CSI_Part__ARCH"],
            log=log,
        )

        result_xml = (word_dir / "styles.xml").read_text(encoding="utf-8")
        assert 'w:styleId="CSI_Part__ARCH"' in result_xml
        assert not any("built-in" in msg.lower() for msg in log), \
            f"Should not skip Normal when it exists in target, got: {log}"

    def test_nonbuiltin_dependency_still_fails_when_missing(self, tmp_path):
        """
        A non-built-in style missing from both architect and target
        must still raise ValueError.
        """
        from spec_formatter.style_application.core.style_import import import_arch_styles_into_target

        arch_xml = (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
            '<w:docDefaults>'
            '<w:rPrDefault><w:rPr/></w:rPrDefault>'
            '<w:pPrDefault><w:pPr/></w:pPrDefault>'
            '</w:docDefaults>'
            '<w:style w:type="paragraph" w:styleId="CSI_Part__ARCH">'
            '<w:name w:val="CSI Part"/>'
            '<w:basedOn w:val="MyCustomBase"/>'
            '</w:style>'
            '</w:styles>'
        )

        word_dir = tmp_path / "word"
        word_dir.mkdir()
        (word_dir / "styles.xml").write_text(
            self.TARGET_STYLES_XML_NO_NORMAL, encoding="utf-8"
        )

        with pytest.raises(ValueError, match="MyCustomBase"):
            import_arch_styles_into_target(
                target_extract_dir=tmp_path,
                arch_styles_xml=arch_xml,
                needed_style_ids=["CSI_Part__ARCH"],
                log=[],
            )

    def test_default_paragraph_font_also_skipped(self, tmp_path):
        """
        DefaultParagraphFont is another common implicit built-in.
        Verify it is also skipped when referenced via link.
        """
        from spec_formatter.style_application.core.style_import import import_arch_styles_into_target

        arch_xml = (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
            '<w:docDefaults>'
            '<w:rPrDefault><w:rPr/></w:rPrDefault>'
            '<w:pPrDefault><w:pPr/></w:pPrDefault>'
            '</w:docDefaults>'
            '<w:style w:type="paragraph" w:styleId="CSI_Article__ARCH">'
            '<w:name w:val="CSI Article"/>'
            '<w:link w:val="DefaultParagraphFont"/>'
            '<w:pPr><w:spacing w:before="200"/></w:pPr>'
            '</w:style>'
            '</w:styles>'
        )

        word_dir = tmp_path / "word"
        word_dir.mkdir()
        (word_dir / "styles.xml").write_text(
            self.TARGET_STYLES_XML_NO_NORMAL, encoding="utf-8"
        )

        log = []
        import_arch_styles_into_target(
            target_extract_dir=tmp_path,
            arch_styles_xml=arch_xml,
            needed_style_ids=["CSI_Article__ARCH"],
            log=log,
        )

        result_xml = (word_dir / "styles.xml").read_text(encoding="utf-8")
        assert 'w:styleId="CSI_Article__ARCH"' in result_xml

def test_conflicting_existing_style_is_not_silently_skipped(tmp_path):
    from spec_formatter.style_application.core.style_import import import_arch_styles_into_target

    word_dir = tmp_path / "word"
    word_dir.mkdir()
    (word_dir / "styles.xml").write_text(
        '<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
        '<w:style w:type="paragraph" w:styleId="CSI"><w:rPr><w:b/></w:rPr></w:style>'
        '</w:styles>',
        encoding="utf-8",
    )
    arch = (
        '<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
        '<w:style w:type="paragraph" w:styleId="CSI"><w:rPr><w:i/></w:rPr></w:style>'
        '</w:styles>'
    )
    log = []
    import_arch_styles_into_target(tmp_path, arch, ["CSI"], log)
    out = (word_dir / "styles.xml").read_text(encoding="utf-8")
    assert "<w:i/>" in out and "<w:b/>" not in out
    assert any("Replaced conflicting" in m for m in log)


def test_equivalent_existing_style_is_left_in_place(tmp_path):
    from spec_formatter.style_application.core.style_import import import_arch_styles_into_target

    style_block = '<w:style w:type="paragraph" w:styleId="CSI"><w:rPr><w:b/></w:rPr></w:style>'
    word_dir = tmp_path / "word"
    word_dir.mkdir()
    (word_dir / "styles.xml").write_text(
        f'<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">{style_block}</w:styles>',
        encoding="utf-8",
    )
    arch = f'<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">{style_block}</w:styles>'
    log = []
    import_arch_styles_into_target(tmp_path, arch, ["CSI"], log)
    assert any("already matches" in m for m in log)
