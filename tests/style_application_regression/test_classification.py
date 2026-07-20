"""Tests for core.classification boilerplate filtering and deterministic classification."""

from spec_formatter.style_application.core.classification import (
    _deterministic_role_for_paragraph,
    _resolve_role,
    strip_boilerplate_with_report,
)


class TestEndOfSectionBoilerplateBehavior:
    """END OF SECTION lines are no longer boilerplate-stripped."""

    def test_plain_end_of_section_passes_through(self):
        cleaned, tags = strip_boilerplate_with_report("END OF SECTION")
        assert cleaned == "END OF SECTION"
        assert tags == []

    def test_end_of_section_with_number_passes_through(self):
        cleaned, tags = strip_boilerplate_with_report("END OF SECTION 211300")
        assert cleaned == "END OF SECTION 211300"
        assert tags == []

    def test_end_of_section_with_spaced_number_passes_through(self):
        cleaned, tags = strip_boilerplate_with_report("END OF SECTION 23 05 13")
        assert cleaned == "END OF SECTION 23 05 13"
        assert tags == []

    def test_end_of_section_case_variants_pass_through(self):
        cleaned, tags = strip_boilerplate_with_report("End Of Section")
        assert cleaned == "End Of Section"
        assert tags == []

    def test_end_of_section_trims_whitespace_only(self):
        cleaned, tags = strip_boilerplate_with_report("  END OF SECTION  ")
        assert cleaned == "END OF SECTION"
        assert tags == []


class TestBoilerplateRegressionChecks:
    def test_real_content_unchanged(self):
        text = "A. Provide valves as specified."
        cleaned, tags = strip_boilerplate_with_report(text)
        assert cleaned == text
        assert tags == []

    def test_article_heading_unchanged(self):
        text = "1.01 SUMMARY"
        cleaned, tags = strip_boilerplate_with_report(text)
        assert cleaned == text
        assert tags == []

    def test_part_heading_unchanged(self):
        text = "PART 1 GENERAL"
        cleaned, tags = strip_boilerplate_with_report(text)
        assert cleaned == text
        assert tags == []

    def test_end_of_section_embedded_sentence_not_modified(self):
        text = "THE END OF SECTION DESCRIBES THE SCOPE"
        cleaned, tags = strip_boilerplate_with_report(text)
        assert cleaned == text
        assert tags == []

    def test_empty_string_no_tags(self):
        cleaned, tags = strip_boilerplate_with_report("")
        assert cleaned == ""
        assert tags == []


class TestEndOfSectionDeterministicClassification:
    def test_deterministic_end_of_section_variants(self):
        for text in [
            "END OF SECTION",
            "End of Section",
            "END OF SECTION 23 05 13",
            "  END OF SECTION  ",
        ]:
            paragraph = {"text": text, "in_table": False, "marker_type": None}
            assert _deterministic_role_for_paragraph(paragraph) == "END_OF_SECTION"

    def test_resolve_role_when_available(self):
        resolved = _resolve_role("END_OF_SECTION", ["END_OF_SECTION", "PART", "ARTICLE"])
        assert resolved == "END_OF_SECTION"

    def test_resolve_role_when_unavailable(self):
        resolved = _resolve_role("END_OF_SECTION", ["PART", "ARTICLE"])
        assert resolved is None
