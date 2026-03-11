"""
Parseability validation tests for arch_template_registry.json generation.

Generates a complete arch_template_registry from a synthetic extracted DOCX
directory and verifies that every XML string value in the resulting JSON is
parseable by xml.etree.ElementTree.

Acceptance criteria:
  - Tests FAIL on the current broken extractor (truncated paired tags produce
    unparseable XML fragments like bare "<w:pPr>" openers).
  - Tests PASS only when paired tags are no longer truncated.
"""

from __future__ import annotations

import xml.etree.ElementTree as ET
from pathlib import Path
from typing import Any, List, Tuple

import pytest

from arch_env_extractor import extract_arch_template_registry


# ---------------------------------------------------------------------------
# Namespace declarations needed to parse OOXML fragments
# ---------------------------------------------------------------------------

NS_DECLS = (
    'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
    'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" '
    'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" '
    'xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" '
    'xmlns:o="urn:schemas-microsoft-com:office:office" '
    'xmlns:v="urn:schemas-microsoft-com:vml" '
    'xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math" '
    'xmlns:wpc="http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas" '
    'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"'
)


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def collect_xml_strings(obj: Any, path: str = "") -> List[Tuple[str, str]]:
    """Yield (json_path, string_value) for all string values that look like XML."""
    if isinstance(obj, dict):
        for k, v in obj.items():
            yield from collect_xml_strings(v, f"{path}.{k}")
    elif isinstance(obj, list):
        for i, v in enumerate(obj):
            yield from collect_xml_strings(v, f"{path}[{i}]")
    elif isinstance(obj, str) and obj.strip().startswith("<"):
        yield (path, obj)


def parse_xml_fragment(xml_str: str) -> None:
    """Parse an XML fragment by wrapping it in a root element with namespace declarations.

    Raises ET.ParseError if the fragment is not well-formed.
    """
    # Some fragments are full XML documents (with <?xml ...?> declaration).
    # Strip the declaration before wrapping.
    stripped = xml_str.strip()
    if stripped.startswith("<?xml"):
        # Remove the XML declaration line
        idx = stripped.index("?>")
        stripped = stripped[idx + 2:].strip()

    ET.fromstring(f"<root {NS_DECLS}>{stripped}</root>")


# ---------------------------------------------------------------------------
# Synthetic DOCX extraction directory fixture
# ---------------------------------------------------------------------------

STYLES_XML = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    '<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
    '<w:docDefaults>'
    '<w:rPrDefault>'
    '<w:rPr>'
    '<w:rFonts w:asciiTheme="minorHAnsi" w:hAnsiTheme="minorHAnsi"'
    ' w:eastAsiaTheme="minorHAnsi" w:cstheme="minorBidi"/>'
    '<w:sz w:val="22"/>'
    '<w:szCs w:val="22"/>'
    '</w:rPr>'
    '</w:rPrDefault>'
    '<w:pPrDefault>'
    '<w:pPr>'
    '<w:spacing w:after="160" w:line="259" w:lineRule="auto"/>'
    '</w:pPr>'
    '</w:pPrDefault>'
    '</w:docDefaults>'
    '<w:latentStyles w:defLockedState="0" w:defUIPriority="99"'
    ' w:defSemiHidden="0" w:defUnhideWhenUsed="0" w:defQFormat="0" w:count="376">'
    '<w:lsdException w:name="Normal" w:uiPriority="0" w:qFormat="1"/>'
    '<w:lsdException w:name="heading 1" w:uiPriority="9" w:qFormat="1"/>'
    '</w:latentStyles>'
    '<w:style w:type="paragraph" w:default="1" w:styleId="Normal">'
    '<w:name w:val="Normal"/>'
    '<w:qFormat/>'
    '<w:pPr>'
    '<w:spacing w:after="200" w:line="276" w:lineRule="auto"/>'
    '</w:pPr>'
    '<w:rPr>'
    '<w:rFonts w:ascii="Calibri" w:hAnsi="Calibri"/>'
    '<w:sz w:val="22"/>'
    '</w:rPr>'
    '</w:style>'
    '<w:style w:type="paragraph" w:styleId="Heading1">'
    '<w:name w:val="heading 1"/>'
    '<w:basedOn w:val="Normal"/>'
    '<w:next w:val="Normal"/>'
    '<w:link w:val="Heading1Char"/>'
    '<w:uiPriority w:val="9"/>'
    '<w:qFormat/>'
    '<w:pPr>'
    '<w:keepNext/>'
    '<w:keepLines/>'
    '<w:spacing w:before="240" w:after="0"/>'
    '<w:outlineLvl w:val="0"/>'
    '</w:pPr>'
    '<w:rPr>'
    '<w:rFonts w:asciiTheme="majorHAnsi" w:hAnsiTheme="majorHAnsi"/>'
    '<w:b/>'
    '<w:color w:val="2F5496" w:themeColor="accent1" w:themeShade="BF"/>'
    '<w:sz w:val="32"/>'
    '</w:rPr>'
    '</w:style>'
    '</w:styles>'
)

DOCUMENT_XML = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
    ' xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
    '<w:body>'
    '<w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr>'
    '<w:r><w:t>Test Document</w:t></w:r></w:p>'
    '<w:p><w:r><w:t>Body text paragraph.</w:t></w:r></w:p>'
    '<w:sectPr>'
    '<w:pgSz w:w="12240" w:h="15840"/>'
    '<w:pgMar w:top="1440" w:right="1440" w:bottom="1440"'
    ' w:left="1440" w:header="720" w:footer="720" w:gutter="0"/>'
    '<w:cols w:space="720"/>'
    '<w:docGrid w:linePitch="360"/>'
    '</w:sectPr>'
    '</w:body>'
    '</w:document>'
)

SETTINGS_XML = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    '<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
    '<w:zoom w:percent="100"/>'
    '<w:compat>'
    '<w:compatSetting w:name="compatibilityMode"'
    ' w:uri="http://schemas.microsoft.com/office/word" w:val="15"/>'
    '<w:compatSetting w:name="overrideTableStyleFontSizeAndJustification"'
    ' w:uri="http://schemas.microsoft.com/office/word" w:val="1"/>'
    '</w:compat>'
    '</w:settings>'
)

NUMBERING_XML = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    '<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
    '<w:abstractNum w:abstractNumId="0">'
    '<w:nsid w:val="12345678"/>'
    '<w:multiLevelType w:val="hybridMultilevel"/>'
    '<w:lvl w:ilvl="0">'
    '<w:start w:val="1"/>'
    '<w:numFmt w:val="decimal"/>'
    '<w:lvlText w:val="%1."/>'
    '<w:lvlJc w:val="left"/>'
    '<w:pPr><w:ind w:left="720" w:hanging="360"/></w:pPr>'
    '</w:lvl>'
    '</w:abstractNum>'
    '<w:num w:numId="1">'
    '<w:abstractNumId w:val="0"/>'
    '</w:num>'
    '</w:numbering>'
)

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

THEME_XML = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    '<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Office Theme">'
    '<a:themeElements>'
    '<a:fontScheme name="Office">'
    '<a:majorFont>'
    '<a:latin typeface="Calibri Light"/>'
    '</a:majorFont>'
    '<a:minorFont>'
    '<a:latin typeface="Calibri"/>'
    '</a:minorFont>'
    '</a:fontScheme>'
    '</a:themeElements>'
    '</a:theme>'
)

RELS_XML = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
    '<Relationship Id="rId1"'
    ' Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles"'
    ' Target="styles.xml"/>'
    '<Relationship Id="rId2"'
    ' Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings"'
    ' Target="settings.xml"/>'
    '</Relationships>'
)


@pytest.fixture
def minimal_extract_dir(tmp_path: Path) -> Path:
    """Build a synthetic extracted DOCX directory with all required parts."""
    word_dir = tmp_path / "word"
    word_dir.mkdir()
    (word_dir / "styles.xml").write_text(STYLES_XML, encoding="utf-8")
    (word_dir / "document.xml").write_text(DOCUMENT_XML, encoding="utf-8")
    (word_dir / "settings.xml").write_text(SETTINGS_XML, encoding="utf-8")
    (word_dir / "numbering.xml").write_text(NUMBERING_XML, encoding="utf-8")
    (word_dir / "fontTable.xml").write_text(FONT_TABLE_XML, encoding="utf-8")

    theme_dir = word_dir / "theme"
    theme_dir.mkdir()
    (theme_dir / "theme1.xml").write_text(THEME_XML, encoding="utf-8")

    rels_dir = word_dir / "_rels"
    rels_dir.mkdir()
    (rels_dir / "document.xml.rels").write_text(RELS_XML, encoding="utf-8")

    return tmp_path


# ---------------------------------------------------------------------------
# Parseability test
# ---------------------------------------------------------------------------


class TestArchTemplateRegistryValidation:
    """Verify that all XML fragments in a generated registry are parseable."""

    def test_full_registry_xml_fragments_parseable(self, minimal_extract_dir: Path):
        """Every XML string value in the generated registry must parse as valid XML."""
        registry = extract_arch_template_registry(minimal_extract_dir)

        failures: list[tuple[str, str, str]] = []
        xml_strings = list(collect_xml_strings(registry))

        # Sanity: we should find a meaningful number of XML fragments
        assert len(xml_strings) > 5, (
            f"Expected many XML fragments in registry, found only {len(xml_strings)}"
        )

        for json_path, xml_value in xml_strings:
            try:
                parse_xml_fragment(xml_value)
            except ET.ParseError as exc:
                failures.append((json_path, str(exc), xml_value[:200]))

        if failures:
            msg_parts = [f"\n{len(failures)} XML fragment(s) failed to parse:\n"]
            for path, error, snippet in failures:
                msg_parts.append(f"  {path}:\n    error: {error}\n    snippet: {snippet!r}\n")
            pytest.fail("".join(msg_parts))
