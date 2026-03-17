from __future__ import annotations

import pytest

from docx_decomposer import validate_instructions


def test_reserved_arch_style_collision_fails_before_apply():
    bundle = {
        "paragraphs": [{"paragraph_index": 0, "text": "1.01 SUMMARY", "contains_sectPr": False, "in_table": False, "skip_reason": None}],
        "style_catalog": {"CSI_Article__ARCH": {"styleId": "CSI_Article__ARCH"}},
    }
    instructions = {
        "create_styles": [{"styleId": "CSI_Article__ARCH", "derive_from_paragraph_index": 0}],
        "apply_pStyle": [{"paragraph_index": 0, "styleId": "CSI_Article__ARCH"}],
        "roles": {"ARTICLE": {"styleId": "CSI_Article__ARCH", "exemplar_paragraph_index": 0}},
    }
    with pytest.raises(ValueError, match="Reserved ARCH style collision"):
        validate_instructions(instructions, slim_bundle=bundle)
