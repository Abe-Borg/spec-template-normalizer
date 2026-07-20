from spec_formatter.style_application.core.classification import _normalize_paragraph_for_contract
from spec_formatter.style_application.core.xml_helpers import strip_conflicting_direct_ppr


def test_inline_numpr_stripped_when_arch_style_has_numpr():
    """Inline numPr must be stripped so architect's style-level numPr takes effect."""
    p_xml = '<w:p><w:pPr><w:pStyle w:val="OldStyle"/><w:numPr><w:ilvl w:val="0"/><w:numId w:val="5"/></w:numPr></w:pPr><w:r><w:t>A. Content</w:t></w:r></w:p>'
    result = strip_conflicting_direct_ppr(p_xml)
    assert "<w:numPr" not in result
    assert "<w:numId" not in result
    assert "<w:ilvl" not in result
    assert "<w:pStyle" in result


def test_no_numpr_passes_through_cleanly():
    """Paragraph without numPr should pass through unchanged (minus other overrides)."""
    p_xml = '<w:p><w:pPr><w:pStyle w:val="OldStyle"/><w:jc w:val="center"/></w:pPr><w:r><w:t>Content</w:t></w:r></w:p>'
    result = strip_conflicting_direct_ppr(p_xml)
    assert "<w:numPr" not in result
    assert "<w:jc" not in result


def test_normalize_strips_numpr_symmetrically():
    """Both before and after should normalize identically when only numPr differs."""
    before = '<w:p><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="5"/></w:numPr></w:pPr><w:r><w:t>Text</w:t></w:r></w:p>'
    after = '<w:p><w:pPr><w:pStyle w:val="CSI_Paragraph__ARCH"/></w:pPr><w:r><w:t>Text</w:t></w:r></w:p>'
    assert _normalize_paragraph_for_contract(before) == _normalize_paragraph_for_contract(after)


def test_self_closing_numpr_stripped():
    """Edge case: some DOCX generators emit self-closing numPr."""
    p_xml = '<w:p><w:pPr><w:numPr/></w:pPr><w:r><w:t>Text</w:t></w:r></w:p>'
    result = strip_conflicting_direct_ppr(p_xml)
    assert "<w:numPr" not in result
