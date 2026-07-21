from __future__ import annotations

from docx_decomposer import validate_instructions
from llm_classifier import _repair_missing_roles
from paragraph_rules import infer_expected_roles
from spec_formatter.numbering_roles import (
    role_from_numbering_catalog,
    role_from_numbering_signature,
)


def _pageformat_catalog(*, through_level: int = 5) -> dict:
    levels = [
        {"ilvl": "0", "numFmt": "decimal", "lvlText": "%1."},
        {"ilvl": "1", "numFmt": "decimal", "lvlText": "%1.%2"},
    ]
    levels.extend(
        {
            "ilvl": str(level),
            "numFmt": "decimal",
            "lvlText": f".%{level + 1}",
        }
        for level in range(2, through_level + 1)
    )
    return {
        "nums": {"1": {"numId": "1", "abstractNumId": "10"}},
        "abstracts": {"10": {"abstractNumId": "10", "levels": levels}},
    }


def _sparse_masterspec_catalog(
    *,
    part_text: str = "PART %1 -",
    article_text: str = "%1.%4",
) -> dict:
    return {
        "nums": {"2": {"numId": "2", "abstractNumId": "1"}},
        "abstracts": {
            "1": {
                "abstractNumId": "1",
                "levels": [
                    {"ilvl": "0", "numFmt": "decimal", "lvlText": part_text},
                    {"ilvl": "1", "numFmt": "lowerLetter", "lvlText": "%2."},
                    {"ilvl": "2", "numFmt": "lowerRoman", "lvlText": "%3."},
                    {"ilvl": "3", "numFmt": "decimal", "lvlText": article_text},
                ],
            }
        },
    }


def test_context_proves_sixth_pageformat_level_without_global_guessing():
    catalog = _pageformat_catalog()

    assert role_from_numbering_signature("decimal", ".%6", "5") is None
    assert (
        role_from_numbering_catalog(catalog, "1", "5")
        == "SUBPARAGRAPH_LEVEL_5"
    )

    unrelated = {
        "nums": {"1": {"abstractNumId": "10"}},
        "abstracts": {
            "10": {
                "levels": [
                    {"ilvl": "5", "numFmt": "decimal", "lvlText": ".%6"}
                ]
            }
        },
    }
    assert role_from_numbering_catalog(unrelated, "1", "5") is None


def test_context_proves_sparse_masterspec_article_without_global_guessing():
    catalog = _sparse_masterspec_catalog()

    assert role_from_numbering_signature("decimal", "%1.%4", "3") is None
    assert role_from_numbering_catalog(catalog, "2", "3") == "ARTICLE"

    assert (
        role_from_numbering_catalog(
            _sparse_masterspec_catalog(part_text="%1."),
            "2",
            "3",
        )
        is None
    )
    assert (
        role_from_numbering_catalog(
            _sparse_masterspec_catalog(article_text="%1.%3"),
            "2",
            "3",
        )
        is None
    )
    assert (
        role_from_numbering_catalog(
            _sparse_masterspec_catalog(article_text="%2.%4"),
            "2",
            "3",
        )
        is None
    )
    assert (
        role_from_numbering_catalog(
            _sparse_masterspec_catalog(article_text="%1.%2.%4"),
            "2",
            "3",
        )
        is None
    )


def test_distinctive_csi_deep_signatures_map_through_word_limit():
    assert (
        role_from_numbering_signature("decimal", "%6)", "5")
        == "SUBPARAGRAPH_LEVEL_5"
    )
    assert (
        role_from_numbering_signature("lowerLetter", "%7)", "6")
        == "SUBPARAGRAPH_LEVEL_6"
    )
    assert (
        role_from_numbering_signature("decimal", "(%8)", "7")
        == "SUBPARAGRAPH_LEVEL_7"
    )
    assert (
        role_from_numbering_signature("lowerLetter", "(%9)", "8")
        == "SUBPARAGRAPH_LEVEL_8"
    )


def test_phase1_repairs_and_validates_missing_sixth_level_role():
    catalog = _pageformat_catalog()
    paragraph = {
        "paragraph_index": 0,
        "text": "Welding requirements",
        "pStyle": "Body",
        "numPr": {"numId": "1", "ilvl": "5"},
        "effective_numPr": {"numId": "1", "ilvl": "5"},
        "pPr_hints": None,
        "rPr_hints": None,
        "has_direct_pPr": True,
        "has_uniform_direct_rPr": False,
        "contains_sectPr": False,
        "in_table": False,
        "skip_reason": None,
    }
    bundle = {
        "paragraphs": [paragraph],
        "style_catalog": {
            "Body": {
                "styleId": "Body",
                "type": "paragraph",
                "default": True,
                "resolved_numPr": None,
            }
        },
        "numbering_catalog": catalog,
    }
    expected, hits = infer_expected_roles(
        bundle["paragraphs"], numbering_catalog=catalog
    )
    assert "SUBPARAGRAPH_LEVEL_5" in expected
    assert hits["SUBPARAGRAPH_LEVEL_5"] == [0]

    instructions = {
        "create_styles": [],
        "apply_pStyle": [{"paragraph_index": 0, "styleId": "Body"}],
        "ignored_paragraphs": [],
        "roles": {},
    }
    assert _repair_missing_roles(instructions, bundle) == 1
    assert instructions["roles"]["SUBPARAGRAPH_LEVEL_5"] == {
        "styleId": "CSI_SubparagraphLevel5__ARCH",
        "exemplar_paragraph_index": 0,
    }
    assert instructions["apply_pStyle"] == [
        {"paragraph_index": 0, "styleId": "CSI_SubparagraphLevel5__ARCH"}
    ]

    validate_instructions(instructions, slim_bundle=bundle)
