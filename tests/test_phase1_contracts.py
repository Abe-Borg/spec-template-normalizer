from __future__ import annotations

from pathlib import Path

import pytest

from arch_env_extractor import extract_arch_template_registry
from docx_decomposer import (
    extract_paragraph_rpr_inner,
    paragraph_ppr_hints_from_block,
    paragraph_rpr_hints_from_block,
    validate_instructions,
)
from gui import _load_prompt_file
from llm_classifier import _parse_response


def _bundle() -> dict:
    return {
        "paragraphs": [
            {"paragraph_index": 0, "text": "SECTION 01 11 00", "contains_sectPr": False, "in_table": False},
            {"paragraph_index": 1, "text": "PROJECT TITLE", "contains_sectPr": False, "in_table": False},
            {"paragraph_index": 2, "text": "PART 1 - GENERAL", "contains_sectPr": False, "in_table": False},
            {"paragraph_index": 3, "text": "[SPECIFIER NOTE: x]", "contains_sectPr": False, "in_table": False},
            {"paragraph_index": 4, "text": "", "contains_sectPr": False, "in_table": False},
            {"paragraph_index": 5, "text": "A. Scope", "contains_sectPr": False, "in_table": True},
        ],
        "style_catalog": {"Normal": {}, "CSI_SectionID__ARCH": {}},
    }


def _instructions() -> dict:
    return {
        "create_styles": [
            {
                "styleId": "CSI_SectionTitle__ARCH",
                "type": "paragraph",
                "derive_from_paragraph_index": 1,
            },
            {
                "styleId": "CSI_Part__ARCH",
                "type": "paragraph",
                "derive_from_paragraph_index": 2,
            },
        ],
        "apply_pStyle": [
            {"paragraph_index": 0, "styleId": "CSI_SectionID__ARCH"},
            {"paragraph_index": 1, "styleId": "CSI_SectionTitle__ARCH"},
            {"paragraph_index": 2, "styleId": "CSI_Part__ARCH"},
        ],
        "roles": {
            "SectionID": {"styleId": "CSI_SectionID__ARCH", "exemplar_paragraph_index": 0},
            "SectionTitle": {"styleId": "CSI_SectionTitle__ARCH", "exemplar_paragraph_index": 1},
            "PART": {"styleId": "CSI_Part__ARCH", "exemplar_paragraph_index": 2},
        },
    }


def test_invalid_json_from_model_rejected():
    with pytest.raises(ValueError, match="not valid JSON"):
        _parse_response("```json\n{invalid\n```")


def test_duplicate_paragraph_indices_rejected():
    data = _instructions()
    data["apply_pStyle"].append({"paragraph_index": 2, "styleId": "CSI_Part__ARCH"})
    with pytest.raises(ValueError, match="Duplicate paragraph_index"):
        validate_instructions(data, slim_bundle=_bundle())


def test_duplicate_created_style_ids_rejected():
    data = _instructions()
    data["create_styles"].append(data["create_styles"][0].copy())
    with pytest.raises(ValueError, match="Duplicate create_styles styleId"):
        validate_instructions(data, slim_bundle=_bundle())


def test_coverage_gap_rejected():
    data = _instructions()
    data["apply_pStyle"] = data["apply_pStyle"][:-1]
    with pytest.raises(ValueError, match="coverage mismatch"):
        validate_instructions(data, slim_bundle=_bundle())


@pytest.mark.parametrize(
    "field,value,match",
    [
        ("text", "", "blank"),
        ("contains_sectPr", True, "contains sectPr"),
        ("in_table", True, "inside a table"),
    ],
)
def test_exemplar_rejection(field, value, match):
    data = _instructions()
    b = _bundle()
    b["paragraphs"][1][field] = value
    with pytest.raises(ValueError, match=match):
        validate_instructions(data, slim_bundle=b)


def test_sectiontitle_naming_consistency():
    assert "SectionName" not in Path("master_prompt.txt").read_text(encoding="utf-8")
    assert "SectionName" not in Path("run_instruction_prompt.txt").read_text(encoding="utf-8")
    assert "SectionName" not in Path("instructions.json").read_text(encoding="utf-8")


def test_rpr_hints_and_whitespace_run_handling():
    para_xml = (
        '<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
        '<w:r><w:rPr><w:b/><w:u w:val="single"/><w:caps/><w:sz w:val="28"/>'
        '<w:rFonts w:ascii="Arial"/></w:rPr><w:t> </w:t></w:r>'
        '<w:r><w:rPr><w:b/><w:u w:val="single"/><w:caps/><w:sz w:val="28"/>'
        '<w:rFonts w:ascii="Arial"/></w:rPr><w:t>Title</w:t></w:r>'
        '</w:p>'
    )
    rpr = extract_paragraph_rpr_inner(para_xml)
    assert "w:b" in rpr
    hints = paragraph_rpr_hints_from_block(para_xml)
    assert hints["bold"] is True
    assert hints["underline"] is True
    assert hints["caps"] is True
    assert hints["sz"] == "28"
    assert hints["font"] == "Arial"


def test_prompt_loader_missing_file_error(tmp_path: Path):
    missing = tmp_path / "master_prompt.txt"
    with pytest.raises(FileNotFoundError, match="Missing required prompt file"):
        _load_prompt_file(missing)


def test_template_registry_contains_raw_style_xml(tmp_path: Path):
    extract_dir = tmp_path / "x"
    (extract_dir / "word" / "theme").mkdir(parents=True)
    (extract_dir / "word" / "_rels").mkdir(parents=True)
    (extract_dir / "word" / "styles.xml").write_text(
        '<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
        '<w:style w:type="paragraph" w:styleId="Normal"><w:name w:val="Normal"/></w:style>'
        '</w:styles>', encoding="utf-8"
    )
    (extract_dir / "word" / "document.xml").write_text(
        '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>'
        '<w:p><w:r><w:t>x</w:t></w:r></w:p><w:sectPr/></w:body></w:document>',
        encoding="utf-8",
    )
    reg = extract_arch_template_registry(extract_dir)
    assert reg["styles"]["style_defs"][0]["raw_style_xml"].startswith("<w:style")


def test_ppr_hints_include_line_rule():
    para_xml = (
        '<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
        '<w:pPr><w:spacing w:before="120" w:line="240" w:lineRule="auto"/></w:pPr>'
        '</w:p>'
    )
    hints = paragraph_ppr_hints_from_block(para_xml)
    assert hints["spacing"]["lineRule"] == "auto"
