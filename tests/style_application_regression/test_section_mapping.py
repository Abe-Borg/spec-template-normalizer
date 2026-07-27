import copy

import pytest

from spec_formatter.style_application.core.section_mapping import (
    choose_section_sources,
    resolve_effective_section_chain,
)


def test_mismatched_sections_use_declared_default_for_every_target_section():
    page_layout = {
        "section_chain": [{"name": "s0"}, {"name": "s1"}],
        "default_section": {"name": "d"},
    }
    out = choose_section_sources(4, page_layout, require_default=True, log=[])
    assert [x["name"] for x in out] == ["d", "d", "d", "d"]


def test_equal_section_counts_apply_one_canonical_shell_to_every_section():
    page_layout = {
        "section_chain": [
            {"name": "first", "page_size": {"w": 12240, "h": 15840}},
            {"name": "body", "page_size": {"w": 12240, "h": 15840}},
        ],
        "default_section": {
            "name": "body",
            "page_size": {"w": 12240, "h": 15840},
        },
    }
    out = choose_section_sources(2, page_layout, require_default=True, log=[])
    assert [x["name"] for x in out] == ["body", "body"]


def test_conflicting_architect_section_shells_are_rejected():
    page_layout = {
        "section_chain": [
            {"page_size": {"w": 10000, "h": 15000}},
            {"page_size": {"w": 12240, "h": 15840}},
        ],
        "default_section": {"page_size": {"w": 12240, "h": 15840}},
    }

    with pytest.raises(ValueError, match="conflicting section shells"):
        choose_section_sources(2, page_layout, require_default=True, log=[])


def test_inherited_header_footer_refs_form_one_effective_shell():
    page_layout = {
        "section_chain": [
            {
                "section_index": 0,
                "page_size": {"w": 11910, "h": 16840},
                "header_refs": {"default": "rId5", "even": None, "first": None},
                "footer_refs": {"default": "rId6", "even": None, "first": None},
            },
            {
                "section_index": 1,
                "page_size": {"w": 11910, "h": 16840},
                "header_refs": {"default": None, "even": None, "first": None},
                "footer_refs": {"default": None, "even": None, "first": None},
            },
        ],
        "default_section": {
            "section_index": 1,
            "page_size": {"w": 11910, "h": 16840},
            "header_refs": {"default": None, "even": None, "first": None},
            "footer_refs": {"default": None, "even": None, "first": None},
        },
    }

    out = choose_section_sources(3, page_layout, require_default=True, log=[])

    assert [section["section_index"] for section in out] == [1, 1, 1]
    assert all(section["header_refs"]["default"] == "rId5" for section in out)
    assert all(section["footer_refs"]["default"] == "rId6" for section in out)


def test_reference_types_and_header_footer_inherit_independently():
    page_layout = {
        "section_chain": [
            {
                "section_index": 0,
                "header_refs": {
                    "default": "header-default",
                    "even": "header-even",
                    "first": "header-first",
                },
                "footer_refs": {
                    "default": "footer-default",
                    "even": "footer-even",
                    "first": "footer-first",
                },
            },
            {
                "section_index": 1,
                "header_refs": {
                    "default": "header-default",
                    "even": None,
                    "first": None,
                },
                "footer_refs": {
                    "default": None,
                    "even": "footer-even",
                    "first": None,
                },
            },
        ],
        "default_section": {"section_index": 1},
    }

    resolved = resolve_effective_section_chain(page_layout)

    assert resolved[1]["header_refs"] == {
        "default": "header-default",
        "even": "header-even",
        "first": "header-first",
    }
    assert resolved[1]["footer_refs"] == {
        "default": "footer-default",
        "even": "footer-even",
        "first": "footer-first",
    }


def test_first_section_null_references_remain_null():
    page_layout = {
        "section_chain": [
            {
                "section_index": 0,
                "header_refs": {"default": None},
                "footer_refs": {},
            }
        ],
        "default_section": {"section_index": 0},
    }

    out = choose_section_sources(1, page_layout, require_default=True, log=[])

    assert out[0]["header_refs"] == {
        "default": None,
        "even": None,
        "first": None,
    }
    assert out[0]["footer_refs"] == {
        "default": None,
        "even": None,
        "first": None,
    }


def test_explicit_header_override_is_a_genuine_shell_conflict():
    page_layout = {
        "section_chain": [
            {
                "section_index": 0,
                "header_refs": {"default": "header-a"},
                "footer_refs": {"default": "footer"},
            },
            {
                "section_index": 1,
                "header_refs": {"default": "header-b"},
                "footer_refs": {"default": None},
            },
        ],
        "default_section": {"section_index": 1},
    }

    with pytest.raises(ValueError, match="conflicting section shells"):
        choose_section_sources(2, page_layout, require_default=True, log=[])


def test_explicit_default_override_is_rejected_after_index_mapping():
    page_layout = {
        "section_chain": [
            {
                "section_index": 0,
                "header_refs": {"default": "header"},
                "footer_refs": {"default": "footer"},
            },
            {
                "section_index": 1,
                "header_refs": {"default": None},
                "footer_refs": {"default": None},
            },
        ],
        "default_section": {
            "section_index": 1,
            "header_refs": {"default": "different-header"},
            "footer_refs": {"default": None},
        },
    }

    with pytest.raises(ValueError, match="default section conflicts"):
        choose_section_sources(2, page_layout, require_default=True, log=[])


def test_resolving_and_mapping_sections_does_not_mutate_raw_registry():
    page_layout = {
        "section_chain": [
            {
                "section_index": 0,
                "header_refs": {"default": "header"},
                "footer_refs": {"default": "footer"},
            },
            {
                "section_index": 1,
                "header_refs": {"default": None},
                "footer_refs": {"default": None},
            },
        ],
        "default_section": {
            "section_index": 1,
            "header_refs": {"default": None},
            "footer_refs": {"default": None},
        },
    }
    original = copy.deepcopy(page_layout)

    resolved = resolve_effective_section_chain(page_layout)
    mapped = choose_section_sources(2, page_layout, require_default=True, log=[])
    resolved[0]["header_refs"]["default"] = "changed"
    mapped[0]["footer_refs"]["default"] = "changed"

    assert page_layout == original
