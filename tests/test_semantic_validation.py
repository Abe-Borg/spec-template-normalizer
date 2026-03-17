from __future__ import annotations

import pytest

from docx_decomposer import validate_instructions


def _auto_numbered_bundle() -> dict:
    return {
        "paragraphs": [
            {"paragraph_index": 0, "text": "General", "contains_sectPr": False, "in_table": False, "skip_reason": None, "pStyle": "Auto", "numPr": {"numId": "5", "ilvl": "0"}},
            {"paragraph_index": 1, "text": "Scope", "contains_sectPr": False, "in_table": False, "skip_reason": None, "pStyle": "Auto", "numPr": {"numId": "5", "ilvl": "1"}},
        ],
        "style_catalog": {
            "Auto": {"styleId": "Auto", "type": "paragraph", "resolved_numPr": {"numId": "5", "ilvl": "0"}}
        },
        "numbering_catalog": {
            "nums": {"5": {"numId": "5", "abstractNumId": "10"}},
            "abstracts": {"10": {"abstractNumId": "10", "levels": [{"ilvl": "0"}, {"ilvl": "1"}]}}
        },
    }


def test_auto_numbered_roles_empty_fails():
    with pytest.raises(ValueError, match="numbered structure detected but roles is empty"):
        validate_instructions({"apply_pStyle": [{"paragraph_index": 0, "styleId": "Auto"}, {"paragraph_index": 1, "styleId": "Auto"}], "roles": {}}, slim_bundle=_auto_numbered_bundle())


def test_shared_style_part_article_allowed():
    bundle = {
        "paragraphs": [
            {"paragraph_index": 0, "text": "PART 1 - GENERAL", "contains_sectPr": False, "in_table": False, "skip_reason": None},
            {"paragraph_index": 1, "text": "1.01 SUMMARY", "contains_sectPr": False, "in_table": False, "skip_reason": None},
        ],
        "style_catalog": {"Normal": {"styleId": "Normal"}},
    }
    validate_instructions(
        {
            "apply_pStyle": [
                {"paragraph_index": 0, "styleId": "CSI_Part__ARCH"},
                {"paragraph_index": 1, "styleId": "CSI_Part__ARCH"},
            ],
            "roles": {
                "PART": {"styleId": "CSI_Part__ARCH", "exemplar_paragraph_index": 0},
                "ARTICLE": {"styleId": "CSI_Part__ARCH", "exemplar_paragraph_index": 1},
            },
        },
        slim_bundle=bundle,
    )
