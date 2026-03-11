"""Tests for phase1_validator — contract validation before writing registries."""

from __future__ import annotations

import pytest

from phase1_validator import (
    validate_template_registry,
    validate_style_registry,
    validate_cross_registry,
    validate_phase1_contracts,
)


# ---------------------------------------------------------------------------
# Fixtures — minimal valid registries
# ---------------------------------------------------------------------------

def _minimal_template_registry():
    """Return a minimal valid template registry dict."""
    return {
        "meta": {
            "schema_version": "1.0.0",
            "source_docx": {"filename": "test.docx", "sha256": "abc123", "extracted_utc": "2025-01-01T00:00:00Z"},
        },
        "package_inventory": {"has_styles": True, "has_theme": True},
        "doc_defaults": {
            "default_run_props": {"rPr": '<w:rPr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:sz w:val="22"/></w:rPr>'},
            "default_paragraph_props": {"pPr": None},
        },
        "styles": {
            "style_defs": [
                {
                    "style_id": "Normal",
                    "name": "Normal",
                    "type": "paragraph",
                    "based_on": None,
                    "pPr": '<w:pPr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:spacing w:after="200"/></w:pPr>',
                    "rPr": '<w:rPr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:sz w:val="22"/></w:rPr>',
                    "tblPr": None,
                    "trPr": None,
                    "tcPr": None,
                },
                {
                    "style_id": "CSI_Part__ARCH",
                    "name": "CSI Part (Architect Template)",
                    "type": "paragraph",
                    "based_on": "Normal",
                    "pPr": None,
                    "rPr": None,
                    "tblPr": None,
                    "trPr": None,
                    "tcPr": None,
                },
                {
                    "style_id": "CSI_Article__ARCH",
                    "name": None,
                    "type": "paragraph",
                    "based_on": None,
                    "pPr": None,
                    "rPr": None,
                    "tblPr": None,
                    "trPr": None,
                    "tcPr": None,
                },
                {
                    "style_id": "CSI_Paragraph__ARCH",
                    "name": None,
                    "type": "paragraph",
                    "based_on": None,
                    "pPr": None,
                    "rPr": None,
                    "tblPr": None,
                    "trPr": None,
                    "tcPr": None,
                },
                {
                    "style_id": "CSI_Subparagraph__ARCH",
                    "name": None,
                    "type": "paragraph",
                    "based_on": None,
                    "pPr": None,
                    "rPr": None,
                    "tblPr": None,
                    "trPr": None,
                    "tcPr": None,
                },
                {
                    "style_id": "CSI_Subsubparagraph__ARCH",
                    "name": None,
                    "type": "paragraph",
                    "based_on": None,
                    "pPr": None,
                    "rPr": None,
                    "tblPr": None,
                    "trPr": None,
                    "tcPr": None,
                },
                {
                    "style_id": "CSI_SectionTitle__ARCH",
                    "name": None,
                    "type": "paragraph",
                    "based_on": None,
                    "pPr": None,
                    "rPr": None,
                    "tblPr": None,
                    "trPr": None,
                    "tcPr": None,
                },
            ],
            "latent_styles": {"latentStyles_xml": None},
            "table_styles": [],
        },
        "theme": {"theme1_xml": None},
        "settings": {"settings_xml": None, "compat": {"compat_xml": None, "important_flags": []}},
        "page_layout": {
            "default_section": {
                "section_index": 0,
                "sectPr": '<w:sectPr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>',
            },
            "section_chain": [],
        },
        "headers_footers": {"headers": [], "footers": []},
        "numbering": {"numbering_xml": None, "abstract_nums": [], "nums": []},
        "fonts": {"font_table_xml": None},
        "custom_xml": {"relationships": [], "other_parts_passthrough": []},
        "capture_policy": {"store_raw_xml_blocks": True},
    }


def _minimal_style_registry():
    """Return a minimal valid style registry dict."""
    return {
        "version": 1,
        "source_docx": "test.docx",
        "roles": {
            "SectionTitle": {"style_id": "CSI_SectionTitle__ARCH", "exemplar_paragraph_index": 0},
            "PART": {"style_id": "CSI_Part__ARCH", "exemplar_paragraph_index": 1},
            "ARTICLE": {"style_id": "CSI_Article__ARCH", "exemplar_paragraph_index": 2},
            "PARAGRAPH": {"style_id": "CSI_Paragraph__ARCH", "exemplar_paragraph_index": 3},
            "SUBPARAGRAPH": {"style_id": "CSI_Subparagraph__ARCH", "exemplar_paragraph_index": 4},
            "SUBSUBPARAGRAPH": {"style_id": "CSI_Subsubparagraph__ARCH", "exemplar_paragraph_index": 5},
        },
    }


# ---------------------------------------------------------------------------
# Template registry validation tests
# ---------------------------------------------------------------------------

class TestValidateTemplateRegistry:

    def test_valid_passes(self):
        validate_template_registry(_minimal_template_registry())

    def test_missing_top_level_key(self):
        reg = _minimal_template_registry()
        del reg["styles"]
        with pytest.raises(ValueError, match="missing required keys"):
            validate_template_registry(reg)

    def test_not_a_dict(self):
        with pytest.raises(ValueError, match="must be a JSON object"):
            validate_template_registry([])

    def test_style_defs_not_list(self):
        reg = _minimal_template_registry()
        reg["styles"]["style_defs"] = "not a list"
        with pytest.raises(ValueError, match="style_defs must be a list"):
            validate_template_registry(reg)

    def test_empty_style_id(self):
        reg = _minimal_template_registry()
        reg["styles"]["style_defs"][0]["style_id"] = ""
        with pytest.raises(ValueError, match="empty or missing style_id"):
            validate_template_registry(reg)

    def test_duplicate_style_id(self):
        reg = _minimal_template_registry()
        reg["styles"]["style_defs"].append({
            "style_id": "Normal",
            "name": "Normal Dup",
            "type": "paragraph",
            "based_on": None,
            "pPr": None, "rPr": None, "tblPr": None, "trPr": None, "tcPr": None,
        })
        with pytest.raises(ValueError, match="Duplicate style_id.*Normal"):
            validate_template_registry(reg)

    def test_malformed_xml_fragment(self):
        reg = _minimal_template_registry()
        reg["styles"]["style_defs"][0]["pPr"] = "<w:pPr><w:broken>"
        with pytest.raises(ValueError, match="Malformed XML fragment"):
            validate_template_registry(reg)

    def test_none_xml_fragments_skipped(self):
        """None values for XML fields should not cause errors."""
        reg = _minimal_template_registry()
        # All XML fields are already None except a few — this should pass
        validate_template_registry(reg)

    def test_malformed_sectpr(self):
        reg = _minimal_template_registry()
        reg["page_layout"]["default_section"]["sectPr"] = "<w:sectPr><unclosed"
        with pytest.raises(ValueError, match="Malformed XML fragment.*sectPr"):
            validate_template_registry(reg)

    def test_malformed_numbering_block(self):
        reg = _minimal_template_registry()
        reg["numbering"]["abstract_nums"] = [
            {"abstractNumId": 0, "xml": "<w:abstractNum><w:broken>"}
        ]
        with pytest.raises(ValueError, match="Malformed XML fragment.*numbering"):
            validate_template_registry(reg)

    def test_malformed_theme(self):
        reg = _minimal_template_registry()
        reg["theme"]["theme1_xml"] = "<a:theme><unclosed"
        with pytest.raises(ValueError, match="Malformed XML fragment.*theme"):
            validate_template_registry(reg)

    def test_malformed_compat(self):
        reg = _minimal_template_registry()
        reg["settings"]["compat"]["compat_xml"] = "<w:compat><oops"
        with pytest.raises(ValueError, match="Malformed XML fragment.*compat"):
            validate_template_registry(reg)

    def test_malformed_font_table(self):
        reg = _minimal_template_registry()
        reg["fonts"]["font_table_xml"] = "<w:fonts><bad"
        with pytest.raises(ValueError, match="Malformed XML fragment.*font"):
            validate_template_registry(reg)

    def test_malformed_latent_styles(self):
        reg = _minimal_template_registry()
        reg["styles"]["latent_styles"]["latentStyles_xml"] = "<w:latentStyles><bad"
        with pytest.raises(ValueError, match="Malformed XML fragment.*latent"):
            validate_template_registry(reg)

    def test_valid_xml_with_xml_declaration(self):
        """XML fragments with <?xml ...?> declarations should parse fine."""
        reg = _minimal_template_registry()
        reg["theme"]["theme1_xml"] = (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Test">'
            '<a:themeElements/></a:theme>'
        )
        validate_template_registry(reg)


# ---------------------------------------------------------------------------
# Style registry validation tests
# ---------------------------------------------------------------------------

class TestValidateStyleRegistry:

    def test_valid_passes(self):
        validate_style_registry(_minimal_style_registry())

    def test_wrong_version(self):
        reg = _minimal_style_registry()
        reg["version"] = 2
        with pytest.raises(ValueError, match="version must be 1"):
            validate_style_registry(reg)

    def test_empty_source_docx(self):
        reg = _minimal_style_registry()
        reg["source_docx"] = ""
        with pytest.raises(ValueError, match="source_docx must be a non-empty string"):
            validate_style_registry(reg)

    def test_missing_required_role(self):
        reg = _minimal_style_registry()
        del reg["roles"]["PART"]
        with pytest.raises(ValueError, match="roles missing.*PART"):
            validate_style_registry(reg)

    def test_empty_style_id_in_role(self):
        reg = _minimal_style_registry()
        reg["roles"]["PART"]["style_id"] = ""
        with pytest.raises(ValueError, match="style_id must be a non-empty string"):
            validate_style_registry(reg)

    def test_negative_exemplar_index(self):
        reg = _minimal_style_registry()
        reg["roles"]["PART"]["exemplar_paragraph_index"] = -1
        with pytest.raises(ValueError, match="exemplar_paragraph_index must be a non-negative int"):
            validate_style_registry(reg)

    def test_optional_section_id(self):
        """SectionID is optional — its absence should not fail."""
        reg = _minimal_style_registry()
        assert "SectionID" not in reg["roles"]
        validate_style_registry(reg)

    def test_style_name_not_required(self):
        """style_name is not required on roles — only style_id is mandatory."""
        reg = _minimal_style_registry()
        # No style_name on any role — should pass
        validate_style_registry(reg)


# ---------------------------------------------------------------------------
# Cross-registry validation tests
# ---------------------------------------------------------------------------

class TestValidateCrossRegistry:

    def test_valid_passes(self):
        validate_cross_registry(_minimal_style_registry(), _minimal_template_registry())

    def test_missing_style_id_in_template(self):
        style_reg = _minimal_style_registry()
        tmpl_reg = _minimal_template_registry()
        # Remove CSI_Part__ARCH from template
        tmpl_reg["styles"]["style_defs"] = [
            s for s in tmpl_reg["styles"]["style_defs"]
            if s["style_id"] != "CSI_Part__ARCH"
        ]
        with pytest.raises(ValueError, match="CSI_Part__ARCH"):
            validate_cross_registry(style_reg, tmpl_reg)

    def test_multiple_missing_ids(self):
        style_reg = _minimal_style_registry()
        tmpl_reg = _minimal_template_registry()
        # Keep only Normal in template
        tmpl_reg["styles"]["style_defs"] = [
            s for s in tmpl_reg["styles"]["style_defs"]
            if s["style_id"] == "Normal"
        ]
        with pytest.raises(ValueError, match="not found in template registry"):
            validate_cross_registry(style_reg, tmpl_reg)


# ---------------------------------------------------------------------------
# Top-level orchestrator tests
# ---------------------------------------------------------------------------

class TestValidatePhase1Contracts:

    def test_valid_passes(self):
        validate_phase1_contracts(_minimal_style_registry(), _minimal_template_registry())

    def test_template_failure_stops_early(self):
        """Template validation failure should prevent style/cross checks."""
        with pytest.raises(ValueError, match="must be a JSON object"):
            validate_phase1_contracts(_minimal_style_registry(), [])

    def test_style_failure_after_template_passes(self):
        style_reg = _minimal_style_registry()
        style_reg["version"] = 99
        with pytest.raises(ValueError, match="version must be 1"):
            validate_phase1_contracts(style_reg, _minimal_template_registry())

    def test_cross_failure_after_both_pass(self):
        style_reg = _minimal_style_registry()
        tmpl_reg = _minimal_template_registry()
        tmpl_reg["styles"]["style_defs"] = [
            s for s in tmpl_reg["styles"]["style_defs"]
            if s["style_id"] == "Normal"
        ]
        with pytest.raises(ValueError, match="not found in template registry"):
            validate_phase1_contracts(style_reg, tmpl_reg)
