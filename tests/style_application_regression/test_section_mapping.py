import pytest

from spec_formatter.style_application.core.section_mapping import choose_section_sources


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
