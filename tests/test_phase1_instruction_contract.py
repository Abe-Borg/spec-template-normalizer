from __future__ import annotations

import json
from pathlib import Path

import pytest

from docx_decomposer import validate_instructions
from phase1_validator import validate_instruction_contract, validate_style_registry


def _instructions() -> dict:
    return {
        "create_styles": [
            {
                "styleId": "CSI_EndOfSection__ARCH",
                "derive_from_paragraph_index": 2,
                "role": "END_OF_SECTION",
            }
        ],
        "apply_pStyle": [
            {"paragraph_index": 0, "styleId": "CSI_Paragraph__ARCH"},
            {"paragraph_index": 2, "styleId": "CSI_EndOfSection__ARCH"},
        ],
        "ignored_paragraphs": [
            {"paragraph_index": 1, "reason": "NON_CSI editor guidance"},
        ],
        "roles": {
            "END_OF_SECTION": {
                "styleId": "CSI_EndOfSection__ARCH",
                "exemplar_paragraph_index": 2,
            }
        },
    }


def test_instruction_contract_supports_end_of_section_and_partitioned_coverage():
    validate_instruction_contract(_instructions(), expected_paragraph_indices=[0, 1, 2])


def test_instruction_and_registry_contracts_accept_deep_subparagraph_role():
    instructions = {
        "create_styles": [
            {
                "styleId": "CSI_SubparagraphLevel5__ARCH",
                "derive_from_paragraph_index": 0,
                "role": "SUBPARAGRAPH_LEVEL_5",
            }
        ],
        "apply_pStyle": [
            {"paragraph_index": 0, "styleId": "CSI_SubparagraphLevel5__ARCH"}
        ],
        "ignored_paragraphs": [],
        "roles": {
            "SUBPARAGRAPH_LEVEL_5": {
                "styleId": "CSI_SubparagraphLevel5__ARCH",
                "exemplar_paragraph_index": 0,
            }
        },
    }
    validate_instruction_contract(instructions, expected_paragraph_indices=[0])
    validate_style_registry(
        {
            "version": 2,
            "source_docx": "source.docx",
            "source_sha256": "a" * 64,
            "source_tokens": {},
            "roles": {
                "SUBPARAGRAPH_LEVEL_5": {
                    "style_id": "CSI_SubparagraphLevel5__ARCH",
                    "exemplar_paragraph_index": 0,
                    "numbering_provenance": "direct_numpr",
                    "numbering_pattern": {
                        "numId": "9",
                        "ilvl": "5",
                        "numFmt": "decimal",
                        "lvlText": ".%6",
                    },
                }
            },
        }
    )


def test_instruction_contract_rejects_overlap_between_applied_and_ignored():
    instructions = _instructions()
    instructions["ignored_paragraphs"].append({"paragraph_index": 0, "reason": "wrong"})
    with pytest.raises(ValueError, match="must be disjoint"):
        validate_instruction_contract(instructions, expected_paragraph_indices=[0, 1, 2])


def test_instruction_contract_rejects_partition_gap():
    instructions = _instructions()
    instructions["ignored_paragraphs"] = []
    with pytest.raises(ValueError, match="partition mismatch"):
        validate_instruction_contract(instructions, expected_paragraph_indices=[0, 1, 2])


def test_instruction_contract_rejects_blank_ignore_reason():
    instructions = _instructions()
    instructions["ignored_paragraphs"][0]["reason"] = "  "
    with pytest.raises(ValueError, match="reason must be a non-empty string"):
        validate_instruction_contract(instructions)


def test_existing_style_cannot_hide_uniform_direct_exemplar_formatting():
    bundle = {
        "paragraphs": [{
            "paragraph_index": 0,
            "text": "PART 1 - GENERAL",
            "pStyle": "ArchitectHeading",
            "numPr": None,
            "effective_numPr": None,
            "pPr_hints": None,
            "rPr_hints": {"bold": True},
            "has_direct_pPr": False,
            "has_uniform_direct_rPr": True,
            "skip_reason": None,
        }],
        "style_catalog": {
            "ArchitectHeading": {
                "styleId": "ArchitectHeading",
                "type": "paragraph",
                "default": False,
            }
        },
        "numbering_catalog": {"nums": {}, "abstracts": {}},
    }
    instructions = {
        "create_styles": [],
        "apply_pStyle": [{"paragraph_index": 0, "styleId": "ArchitectHeading"}],
        "ignored_paragraphs": [],
        "roles": {
            "PART": {
                "styleId": "ArchitectHeading",
                "exemplar_paragraph_index": 0,
            }
        },
    }
    with pytest.raises(ValueError, match="contains direct formatting"):
        validate_instructions(instructions, slim_bundle=bundle)


def test_v2_style_contract_accepts_source_tokens_hash_end_role_and_extended_numbering():
    validate_style_registry(
        {
            "version": 2,
            "source_docx": "source.docx",
            "source_sha256": "a" * 64,
            "source_tokens": {"SectionID": "SECTION 01 00 00"},
            "roles": {
                "END_OF_SECTION": {
                    "style_id": "CSI_EndOfSection__ARCH",
                    "exemplar_paragraph_index": 4,
                    "numbering_provenance": "direct_numpr",
                    "numbering_pattern": {
                        "numId": "7",
                        "ilvl": "0",
                        "abstractNumId": "12",
                        "startOverride": "1",
                    },
                }
            },
        }
    )


@pytest.mark.parametrize("missing", ["source_sha256", "source_tokens"])
def test_v2_style_contract_requires_bundle_identity_fields(missing):
    registry = {
        "version": 2,
        "source_docx": "source.docx",
        "source_sha256": "a" * 64,
        "source_tokens": {},
        "roles": {
            "PART": {
                "style_id": "PartStyle",
                "exemplar_paragraph_index": 0,
                "numbering_provenance": "none",
            }
        },
    }
    del registry[missing]
    with pytest.raises(ValueError, match=f"requires {missing}"):
        validate_style_registry(registry)


def test_v2_style_contract_rejects_empty_roles_and_missing_numbering_metadata():
    base = {
        "version": 2,
        "source_docx": "source.docx",
        "source_sha256": "a" * 64,
        "source_tokens": {},
        "roles": {},
    }
    with pytest.raises(ValueError, match="roles must not be empty"):
        validate_style_registry(base)

    base["roles"] = {"PART": {"style_id": "PartStyle", "exemplar_paragraph_index": 0}}
    with pytest.raises(ValueError, match="numbering_provenance is required"):
        validate_style_registry(base)

    base["roles"]["PART"]["numbering_provenance"] = "style_numpr"
    with pytest.raises(ValueError, match="requires numbering_pattern.numId"):
        validate_style_registry(base)


def test_json_schemas_expose_runtime_contract_fields():
    schema_dir = Path(__file__).parents[1] / "schemas"
    instructions = json.loads((schema_dir / "phase1_instructions.schema.json").read_text(encoding="utf-8"))
    style_v2 = json.loads((schema_dir / "arch_style_registry.v2.schema.json").read_text(encoding="utf-8"))

    assert "ignored_paragraphs" in instructions["properties"]
    assert "END_OF_SECTION" in instructions["properties"]["roles"]["properties"]
    assert "CSI_EndOfSection__ARCH" in instructions["$defs"]["reservedStyleId"]["enum"]
    assert "source_tokens" in style_v2["properties"]
    assert "source_sha256" in style_v2["properties"]
    assert "END_OF_SECTION" in style_v2["properties"]["roles"]["properties"]
    for level in range(5, 9):
        role = f"SUBPARAGRAPH_LEVEL_{level}"
        style_id = f"CSI_SubparagraphLevel{level}__ARCH"
        assert role in instructions["properties"]["roles"]["properties"]
        assert style_id in instructions["$defs"]["reservedStyleId"]["enum"]
        assert role in style_v2["properties"]["roles"]["properties"]
