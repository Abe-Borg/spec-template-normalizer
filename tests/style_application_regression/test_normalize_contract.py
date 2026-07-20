"""Tests for _normalize_paragraph_for_contract in core.classification."""

from spec_formatter.style_application.core.classification import _normalize_paragraph_for_contract


class TestNormalizeParagraphForContract:
    """Verify that contract normalization strips allowed elements and
    cleans up empty wrapper blocks so before/after comparisons don't
    false-fail."""

    def test_strips_pstyle(self):
        p = '<w:p><w:pPr><w:pStyle w:val="Normal"/><w:jc w:val="center"/></w:pPr><w:r><w:t>Hi</w:t></w:r></w:p>'
        result = _normalize_paragraph_for_contract(p)
        assert "w:pStyle" not in result
        assert 'w:jc w:val="center"' not in result

    def test_removes_empty_ppr_after_strip(self):
        """pPr that only contained a pStyle should vanish entirely."""
        p = '<w:p><w:pPr><w:pStyle w:val="Normal"/></w:pPr><w:r><w:t>Hi</w:t></w:r></w:p>'
        result = _normalize_paragraph_for_contract(p)
        assert "<w:pPr" not in result

    def test_removes_self_closing_ppr(self):
        p = '<w:p><w:pPr/><w:r><w:t>Hi</w:t></w:r></w:p>'
        result = _normalize_paragraph_for_contract(p)
        assert "<w:pPr" not in result

    def test_no_ppr_unchanged(self):
        """Paragraph without pPr stays the same."""
        p = '<w:p><w:r><w:t>Hello</w:t></w:r></w:p>'
        result = _normalize_paragraph_for_contract(p)
        assert result == p

    def test_pstyle_only_matches_no_ppr(self):
        """A paragraph whose pPr held only pStyle should normalize
        identically to one that never had pPr at all."""
        with_ppr = '<w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:r><w:t>X</w:t></w:r></w:p>'
        without_ppr = '<w:p><w:r><w:t>X</w:t></w:r></w:p>'
        assert _normalize_paragraph_for_contract(with_ppr) == _normalize_paragraph_for_contract(without_ppr)

    def test_removes_empty_ppr_with_whitespace(self):
        """pPr with only whitespace inside after stripping should be removed."""
        p = '<w:p><w:pPr>  \n  </w:pPr><w:r><w:t>Hi</w:t></w:r></w:p>'
        result = _normalize_paragraph_for_contract(p)
        assert "<w:pPr" not in result

    def test_strips_numpr(self):
        p = '<w:p><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr></w:pPr><w:r><w:t>A</w:t></w:r></w:p>'
        result = _normalize_paragraph_for_contract(p)
        assert "w:numPr" not in result
        assert "<w:pPr" not in result  # empty pPr cleaned up too

    def test_strips_run_font_formatting(self):
        p = '<w:p><w:r><w:rPr><w:rFonts w:ascii="Arial"/><w:sz w:val="24"/><w:szCs w:val="24"/></w:rPr><w:t>Hi</w:t></w:r></w:p>'
        result = _normalize_paragraph_for_contract(p)
        assert "w:rFonts" not in result
        assert "w:sz" not in result
        assert "<w:rPr" not in result  # empty rPr cleaned up too

    def test_mixed_ppr_preserves_remaining(self):
        """pPr with only allowed-change tags should normalize to no pPr."""
        p = (
            '<w:p><w:pPr>'
            '<w:pStyle w:val="Normal"/>'
            '<w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr>'
            '<w:jc w:val="center"/>'
            '</w:pPr><w:r><w:t>Hi</w:t></w:r></w:p>'
        )
        result = _normalize_paragraph_for_contract(p)
        assert "w:pStyle" not in result
        assert "w:numPr" not in result
        assert 'w:jc w:val="center"' not in result
        assert "<w:pPr" not in result

    def test_both_empty_ppr_and_rpr_cleaned(self):
        """pPr with only pStyle + rPr with only fonts — both shells removed."""
        p = (
            '<w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr>'
            '<w:r><w:rPr><w:rFonts w:ascii="Arial"/><w:sz w:val="24"/><w:szCs w:val="24"/></w:rPr>'
            '<w:t>X</w:t></w:r></w:p>'
        )
        result = _normalize_paragraph_for_contract(p)
        assert "<w:pPr" not in result
        assert "<w:rPr" not in result
        assert "<w:t>X</w:t>" in result

    def test_self_closing_rpr_cleaned(self):
        """Self-closing <w:rPr/> should be removed."""
        p = '<w:p><w:r><w:rPr/><w:t>Hi</w:t></w:r></w:p>'
        result = _normalize_paragraph_for_contract(p)
        assert "<w:rPr" not in result
