"""Tests for arch_env_applier.py — settings and font table hardening."""

import pytest
from pathlib import Path

from spec_formatter.style_application.arch_env_applier import apply_settings, apply_font_table


# ---------------------------------------------------------------------------
# Helpers to build minimal extract directories
# ---------------------------------------------------------------------------

_CONTENT_TYPES_XML = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
    '</Types>'
)

_RELS_XML = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
    '  <Relationship Id="rId1" '
    'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" '
    'Target="document.xml"/>'
    '</Relationships>'
)

_SETTINGS_XML = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    '<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
    '</w:settings>'
)

_SETTINGS_WITH_COMPAT_XML = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    '<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
    '  <w:compat><w:useFELayout/></w:compat>'
    '</w:settings>'
)

_FONT_TABLE_XML = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    '<w:fonts xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
    '  <w:font w:name="Calibri"><w:panose1 w:val="020F0502020204030204"/></w:font>'
    '</w:fonts>'
)


def _setup_extract_dir(tmp_path, content_types=True, rels=True):
    """Create a minimal extract directory with optional plumbing files."""
    (tmp_path / "[Content_Types].xml").write_text(
        _CONTENT_TYPES_XML if content_types else "", encoding="utf-8"
    )
    rels_dir = tmp_path / "word" / "_rels"
    rels_dir.mkdir(parents=True, exist_ok=True)
    if rels:
        (rels_dir / "document.xml.rels").write_text(_RELS_XML, encoding="utf-8")
    (tmp_path / "word").mkdir(exist_ok=True)
    return tmp_path


# ---------------------------------------------------------------------------
# Part A — apply_settings() tests
# ---------------------------------------------------------------------------

class TestApplySettingsRejectsMalformed:
    """Malformed compat_xml must be rejected without mutating anything."""

    def test_missing_close_tag(self, tmp_path):
        extract = _setup_extract_dir(tmp_path)
        settings_path = extract / "word" / "settings.xml"
        settings_path.write_text(_SETTINGS_XML, encoding="utf-8")

        registry = {
            "settings": {
                "compat": {"compat_xml": "<w:compat><w:useFELayout/>"}  # no close
            }
        }
        log = []
        apply_settings(extract, registry, log)

        assert any("WARNING" in m and "Skipping compat" in m for m in log)
        # settings.xml must be untouched
        assert settings_path.read_text(encoding="utf-8") == _SETTINGS_XML

    def test_missing_open_tag(self, tmp_path):
        extract = _setup_extract_dir(tmp_path)
        settings_path = extract / "word" / "settings.xml"
        settings_path.write_text(_SETTINGS_XML, encoding="utf-8")

        registry = {
            "settings": {
                "compat": {"compat_xml": "</w:compat>"}
            }
        }
        log = []
        apply_settings(extract, registry, log)

        assert any("WARNING" in m and "Skipping compat" in m for m in log)
        assert settings_path.read_text(encoding="utf-8") == _SETTINGS_XML

    def test_empty_string_compat_skips(self, tmp_path):
        """Empty string compat_xml is falsy — skipped before validation."""
        extract = _setup_extract_dir(tmp_path)
        settings_path = extract / "word" / "settings.xml"
        settings_path.write_text(_SETTINGS_XML, encoding="utf-8")

        registry = {"settings": {"compat": {"compat_xml": ""}}}
        log = []
        apply_settings(extract, registry, log)

        assert any("No compat" in m or "skipping" in m.lower() for m in log)
        assert settings_path.read_text(encoding="utf-8") == _SETTINGS_XML

    def test_none_compat_skips(self, tmp_path):
        """None compat_xml is falsy — skipped before validation."""
        extract = _setup_extract_dir(tmp_path)
        settings_path = extract / "word" / "settings.xml"
        settings_path.write_text(_SETTINGS_XML, encoding="utf-8")

        registry = {"settings": {"compat": {"compat_xml": None}}}
        log = []
        apply_settings(extract, registry, log)

        assert any("No compat" in m or "skipping" in m.lower() for m in log)
        assert settings_path.read_text(encoding="utf-8") == _SETTINGS_XML


class TestApplySettingsCreatesWhenMissing:
    """When target has no settings.xml, create one and wire plumbing."""

    def test_creates_settings_with_compat(self, tmp_path):
        extract = _setup_extract_dir(tmp_path)
        compat = "<w:compat><w:useFELayout/></w:compat>"
        registry = {"settings": {"compat": {"compat_xml": compat}}}

        log = []
        apply_settings(extract, registry, log)

        settings_path = extract / "word" / "settings.xml"
        assert settings_path.exists()
        content = settings_path.read_text(encoding="utf-8")
        assert "<w:useFELayout/>" in content
        assert "<w:settings" in content
        assert "</w:settings>" in content

    def test_wires_content_types(self, tmp_path):
        extract = _setup_extract_dir(tmp_path)
        compat = "<w:compat><w:useFELayout/></w:compat>"
        registry = {"settings": {"compat": {"compat_xml": compat}}}

        log = []
        apply_settings(extract, registry, log)

        ct = (extract / "[Content_Types].xml").read_text(encoding="utf-8")
        assert 'PartName="/word/settings.xml"' in ct
        assert "wordprocessingml.settings+xml" in ct

    def test_wires_rels(self, tmp_path):
        extract = _setup_extract_dir(tmp_path)
        compat = "<w:compat><w:useFELayout/></w:compat>"
        registry = {"settings": {"compat": {"compat_xml": compat}}}

        log = []
        apply_settings(extract, registry, log)

        rels = (extract / "word" / "_rels" / "document.xml.rels").read_text(encoding="utf-8")
        assert 'Target="settings.xml"' in rels
        assert "relationships/settings" in rels


class TestApplySettingsReplacesExisting:
    """Existing compat block should be replaced."""

    def test_replaces_compat(self, tmp_path):
        extract = _setup_extract_dir(tmp_path)
        settings_path = extract / "word" / "settings.xml"
        settings_path.write_text(_SETTINGS_WITH_COMPAT_XML, encoding="utf-8")

        new_compat = "<w:compat><w:compatSetting w:name=\"test\" w:val=\"1\"/></w:compat>"
        registry = {"settings": {"compat": {"compat_xml": new_compat}}}

        log = []
        apply_settings(extract, registry, log)

        content = settings_path.read_text(encoding="utf-8")
        assert "w:compatSetting" in content
        assert "<w:useFELayout/>" not in content
        assert any("Replaced" in m for m in log)


class TestApplySettingsInserts:
    """When settings.xml exists but has no compat, insert it."""

    def test_inserts_compat(self, tmp_path):
        extract = _setup_extract_dir(tmp_path)
        settings_path = extract / "word" / "settings.xml"
        settings_path.write_text(_SETTINGS_XML, encoding="utf-8")

        compat = "<w:compat><w:useFELayout/></w:compat>"
        registry = {"settings": {"compat": {"compat_xml": compat}}}

        log = []
        apply_settings(extract, registry, log)

        content = settings_path.read_text(encoding="utf-8")
        assert "<w:useFELayout/>" in content
        assert "</w:settings>" in content
        assert any("Inserted" in m for m in log)


class TestApplySettingsIdempotent:
    """Calling twice should not duplicate plumbing entries."""

    def test_idempotent_content_types_and_rels(self, tmp_path):
        extract = _setup_extract_dir(tmp_path)
        compat = "<w:compat><w:useFELayout/></w:compat>"
        registry = {"settings": {"compat": {"compat_xml": compat}}}

        apply_settings(extract, registry, [])
        apply_settings(extract, registry, [])

        ct = (extract / "[Content_Types].xml").read_text(encoding="utf-8")
        assert ct.count('PartName="/word/settings.xml"') == 1

        rels = (extract / "word" / "_rels" / "document.xml.rels").read_text(encoding="utf-8")
        assert rels.count('Target="settings.xml"') == 1


class TestApplySettingsRawArtifactIsProvenanceOnly:
    def test_raw_settings_never_overwrites_target_semantics_or_relationships(self, tmp_path):
        extract = _setup_extract_dir(tmp_path)
        settings_path = extract / "word" / "settings.xml"
        settings_path.write_text(
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
            'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
            '<w:docId w:val="AAAA"/>'
            '<w:trackRevisions/>'
            '<w:documentProtection w:edit="readOnly"/>'
            '<w:attachedTemplate r:id="rIdTargetTemplate"/>'
            '<w:rsids><w:rsidRoot w:val="BBBB"/></w:rsids>'
            '<w:compat><w:useFELayout/></w:compat>'
            '</w:settings>',
            encoding="utf-8",
        )
        registry_dir = tmp_path / "registry"
        registry_dir.mkdir()
        (registry_dir / "arch_settings_raw.xml").write_text(
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
            'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
            '<w:docId w:val="ARCH"/>'
            '<w:mailMerge><w:dataSource r:id="rIdMissing"/></w:mailMerge>'
            '<w:rsids><w:rsidRoot w:val="ARCHRSID"/></w:rsids>'
            '</w:settings>',
            encoding="utf-8",
        )

        log = []
        registry = {
            "settings": {
                "compat": {
                    "compat_xml": '<w:compat><w:compatSetting w:name="compatibilityMode" w:val="15"/></w:compat>'
                }
            }
        }
        apply_settings(extract, registry, log, registry_dir=registry_dir)
        out = settings_path.read_text(encoding="utf-8")
        assert '<w:docId w:val="AAAA"/>' in out
        assert "trackRevisions" in out
        assert "documentProtection" in out
        assert 'r:id="rIdTargetTemplate"' in out
        assert "BBBB" in out
        assert "compatibilityMode" in out
        assert "ARCHRSID" not in out
        assert "rIdMissing" not in out


# ---------------------------------------------------------------------------
# Part B — apply_font_table() tests
# ---------------------------------------------------------------------------

class TestApplyFontTableCreatesWithPlumbing:
    """New fontTable.xml should be wired into content types and rels."""

    def test_creates_font_table(self, tmp_path):
        extract = _setup_extract_dir(tmp_path)
        arch_font_xml = (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<w:fonts xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
            '  <w:font w:name="Arial"><w:panose1 w:val="020B0604020202020204"/></w:font>'
            '</w:fonts>'
        )
        registry = {"fonts": {"font_table_xml": arch_font_xml}}

        log = []
        apply_font_table(extract, registry, log)

        font_path = extract / "word" / "fontTable.xml"
        assert font_path.exists()
        assert "Arial" in font_path.read_text(encoding="utf-8")

    def test_wires_content_types(self, tmp_path):
        extract = _setup_extract_dir(tmp_path)
        arch_font_xml = (
            '<w:fonts xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
            '  <w:font w:name="Arial"><w:panose1 w:val="020B0604020202020204"/></w:font>'
            '</w:fonts>'
        )
        registry = {"fonts": {"font_table_xml": arch_font_xml}}

        log = []
        apply_font_table(extract, registry, log)

        ct = (extract / "[Content_Types].xml").read_text(encoding="utf-8")
        assert 'PartName="/word/fontTable.xml"' in ct
        assert "wordprocessingml.fontTable+xml" in ct

    def test_wires_rels(self, tmp_path):
        extract = _setup_extract_dir(tmp_path)
        arch_font_xml = (
            '<w:fonts xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
            '  <w:font w:name="Arial"><w:panose1 w:val="020B0604020202020204"/></w:font>'
            '</w:fonts>'
        )
        registry = {"fonts": {"font_table_xml": arch_font_xml}}

        log = []
        apply_font_table(extract, registry, log)

        rels = (extract / "word" / "_rels" / "document.xml.rels").read_text(encoding="utf-8")
        assert 'Target="fontTable.xml"' in rels
        assert "relationships/fontTable" in rels

    def test_content_type_mime(self, tmp_path):
        """Content type must use the correct OOXML MIME type."""
        extract = _setup_extract_dir(tmp_path)
        arch_font_xml = (
            '<w:fonts xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
            '  <w:font w:name="Arial"><w:panose1 w:val="020B0604020202020204"/></w:font>'
            '</w:fonts>'
        )
        registry = {"fonts": {"font_table_xml": arch_font_xml}}
        apply_font_table(extract, registry, [])

        ct = (extract / "[Content_Types].xml").read_text(encoding="utf-8")
        assert "application/vnd.openxmlformats-officedocument.wordprocessingml.fontTable+xml" in ct

    def test_rels_relationship_type_uri(self, tmp_path):
        """Rels must use the correct OOXML relationship type URI."""
        extract = _setup_extract_dir(tmp_path)
        arch_font_xml = (
            '<w:fonts xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
            '  <w:font w:name="Arial"><w:panose1 w:val="020B0604020202020204"/></w:font>'
            '</w:fonts>'
        )
        registry = {"fonts": {"font_table_xml": arch_font_xml}}
        apply_font_table(extract, registry, [])

        rels = (extract / "word" / "_rels" / "document.xml.rels").read_text(encoding="utf-8")
        assert "http://schemas.openxmlformats.org/officeDocument/2006/relationships/fontTable" in rels


class TestApplyFontTableMerge:
    """When fontTable exists, only missing fonts are added."""

    def test_adds_missing_fonts(self, tmp_path):
        extract = _setup_extract_dir(tmp_path)
        font_path = extract / "word" / "fontTable.xml"
        font_path.write_text(_FONT_TABLE_XML, encoding="utf-8")

        arch_font_xml = (
            '<w:fonts xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
            '  <w:font w:name="Calibri"><w:panose1 w:val="020F0502020204030204"/></w:font>'
            '  <w:font w:name="Arial"><w:panose1 w:val="020B0604020202020204"/></w:font>'
            '</w:fonts>'
        )
        registry = {"fonts": {"font_table_xml": arch_font_xml}}

        log = []
        apply_font_table(extract, registry, log)

        content = font_path.read_text(encoding="utf-8")
        assert "Arial" in content
        assert content.count('w:name="Calibri"') == 1  # not duplicated
        assert any("1 font" in m for m in log)

    def test_multiple_existing_fonts_only_missing_added(self, tmp_path):
        """Target with multiple fonts — only truly missing ones added."""
        extract = _setup_extract_dir(tmp_path)
        font_path = extract / "word" / "fontTable.xml"
        multi_font_xml = (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<w:fonts xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
            '  <w:font w:name="Calibri"><w:panose1 w:val="020F0502020204030204"/></w:font>'
            '  <w:font w:name="Times New Roman"><w:panose1 w:val="02020603050405020304"/></w:font>'
            '</w:fonts>'
        )
        font_path.write_text(multi_font_xml, encoding="utf-8")

        arch_font_xml = (
            '<w:fonts xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
            '  <w:font w:name="Calibri"><w:panose1 w:val="020F0502020204030204"/></w:font>'
            '  <w:font w:name="Times New Roman"><w:panose1 w:val="02020603050405020304"/></w:font>'
            '  <w:font w:name="Arial"><w:panose1 w:val="020B0604020202020204"/></w:font>'
            '</w:fonts>'
        )
        registry = {"fonts": {"font_table_xml": arch_font_xml}}

        log = []
        apply_font_table(extract, registry, log)

        content = font_path.read_text(encoding="utf-8")
        assert "Arial" in content
        assert content.count('w:name="Calibri"') == 1
        assert content.count('w:name="Times New Roman"') == 1

    def test_skips_when_all_present(self, tmp_path):
        extract = _setup_extract_dir(tmp_path)
        font_path = extract / "word" / "fontTable.xml"
        font_path.write_text(_FONT_TABLE_XML, encoding="utf-8")

        arch_font_xml = (
            '<w:fonts xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
            '  <w:font w:name="Calibri"><w:panose1 w:val="020F0502020204030204"/></w:font>'
            '</w:fonts>'
        )
        registry = {"fonts": {"font_table_xml": arch_font_xml}}

        log = []
        apply_font_table(extract, registry, log)

        assert any("already present" in m for m in log)


class TestApplyFontTableIdempotent:
    """Calling twice should not duplicate plumbing entries."""

    def test_idempotent_plumbing(self, tmp_path):
        extract = _setup_extract_dir(tmp_path)
        arch_font_xml = (
            '<w:fonts xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
            '  <w:font w:name="Arial"><w:panose1 w:val="020B0604020202020204"/></w:font>'
            '</w:fonts>'
        )
        registry = {"fonts": {"font_table_xml": arch_font_xml}}

        apply_font_table(extract, registry, [])
        apply_font_table(extract, registry, [])

        ct = (extract / "[Content_Types].xml").read_text(encoding="utf-8")
        assert ct.count('PartName="/word/fontTable.xml"') == 1

        rels = (extract / "word" / "_rels" / "document.xml.rels").read_text(encoding="utf-8")
        assert rels.count('Target="fontTable.xml"') == 1


class TestApplyFontTableValidation:
    """Post-mutation validation catches malformed results."""

    def test_valid_result_no_warning(self, tmp_path):
        extract = _setup_extract_dir(tmp_path)
        arch_font_xml = (
            '<w:fonts xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
            '  <w:font w:name="Arial"><w:panose1 w:val="020B0604020202020204"/></w:font>'
            '</w:fonts>'
        )
        registry = {"fonts": {"font_table_xml": arch_font_xml}}

        log = []
        apply_font_table(extract, registry, log)

        assert not any("WARNING" in m and "malformed" in m for m in log)
