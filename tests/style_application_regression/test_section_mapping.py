from spec_formatter.style_application.core.section_mapping import choose_section_sources


def test_mismatched_sections_use_declared_default_for_every_target_section():
    page_layout = {
        "section_chain": [{"name": "s0"}, {"name": "s1"}],
        "default_section": {"name": "d"},
    }
    out = choose_section_sources(4, page_layout, require_default=True, log=[])
    assert [x["name"] for x in out] == ["d", "d", "d", "d"]


def test_equal_section_counts_preserve_positionally_mapped_sections():
    page_layout = {
        "section_chain": [{"name": "cover"}, {"name": "body"}],
        "default_section": {"name": "body"},
    }
    out = choose_section_sources(2, page_layout, require_default=True, log=[])
    assert [x["name"] for x in out] == ["cover", "body"]
