"""The built-in CSC scheme must earn its way past the architect's own contract.

The point of these tests is that they call the *unmodified* validators from
``core/csi_to_canadian`` -- the ones a real architect template has to satisfy
before Canadian conversion will touch a target. Nothing here is a parallel,
more forgiving check written for the built-in scheme's benefit. If a future
edit to ``builtin_scheme`` breaks the Canadian hierarchy, it fails here rather
than in somebody's published document.
"""

from __future__ import annotations

import re

import pytest

from spec_formatter import builtin_scheme
from spec_formatter.role_contract import (
    BODY_HIERARCHY_ROLES,
    ROLE_LEVEL,
    ROLE_ORDER,
    ROLE_TO_ARCH_STYLE,
)
from spec_formatter.style_application.batch_runner import builtin_shared_config
from spec_formatter.style_application.core import csi_to_canadian
from spec_formatter.style_application.core.registry import preflight_validate_registries


BODY_ROLES = set(BODY_HIERARCHY_ROLES)


def test_every_body_role_satisfies_the_canadian_role_contract() -> None:
    specs = builtin_scheme.build_role_specs()
    for role in sorted(BODY_ROLES):
        csi_to_canadian._validate_canadian_role_contract(
            role,
            specs[role],
            allow_numeric_part=(role == "PART"),
        )


def test_roles_form_one_coherent_multilevel_list() -> None:
    specs = builtin_scheme.build_role_specs()
    csi_to_canadian._validate_complete_article_hierarchy(specs, BODY_ROLES)

    num_ids = {
        specs[role]["numbering_pattern"]["numId"] for role in BODY_ROLES
    }
    assert num_ids == {builtin_scheme.BUILTIN_NUM_ID}


def test_generated_numbering_resolves_and_starts_at_one() -> None:
    csi_to_canadian._validate_architect_numbering(
        builtin_scheme.build_numbering_xml(),
        builtin_scheme.build_role_specs(),
        BODY_ROLES,
    )


def test_level_text_matches_csc_pageformat() -> None:
    numbering = builtin_scheme.build_numbering_xml()
    assert re.findall(r'<w:lvlText w:val="([^"]+)"/>', numbering) == [
        "PART %1",
        "%1.%2",
        ".%3",
        ".%4",
        ".%5",
        ".%6",
        ".%7",
        ".%8",
        ".%9",
    ]
    assert "<w:lvlRestart" not in numbering
    assert numbering.count('<w:numFmt w:val="decimal"/>') == len(BODY_HIERARCHY_ROLES)


def test_role_styles_carry_their_own_level_and_no_font_overrides() -> None:
    styles = builtin_scheme.build_styles_xml()
    for role in ROLE_ORDER:
        assert f'w:styleId="{ROLE_TO_ARCH_STYLE[role]}"' in styles
    for role in BODY_HIERARCHY_ROLES:
        block = styles[styles.index(f'w:styleId="{ROLE_TO_ARCH_STYLE[role]}"'):]
        block = block[: block.index("</w:style>")]
        assert f'<w:ilvl w:val="{ROLE_LEVEL[role]}"/>' in block
        assert f'<w:numId w:val="{builtin_scheme.BUILTIN_NUM_ID}"/>' in block
    # No run properties at all: the target's own theme and defaults decide how
    # the text looks, which is what "the shell is left untouched" has to mean
    # on the page and not only in the headers and footers.
    assert "<w:rPr>" not in styles.replace("<w:rPr/>", "")


def test_shared_config_carries_no_shell_and_no_bundle() -> None:
    config = builtin_shared_config()
    assert config.builtin_scheme is True
    assert config.arch_root is None
    assert config.bundle_manifest is None
    assert config.source_tokens == {}
    assert config.env_registry["headers_footers"] == {
        "headers": [],
        "footers": [],
        "header_footer_media": [],
    }
    assert config.env_registry["page_layout"]["section_chain"] == []
    assert set(config.available_roles) == set(ROLE_ORDER)


def test_registry_preflight_passes_without_the_shell_checks() -> None:
    assert (
        preflight_validate_registries(
            builtin_scheme.build_arch_registry(),
            builtin_scheme.build_env_registry(),
            additional_known_style_ids=set(builtin_scheme.build_arch_registry().values()),
            applies_shell=False,
        )
        == []
    )


def test_preflight_still_demands_a_page_layout_when_a_shell_is_applied() -> None:
    """``applies_shell`` narrows the check; it does not weaken it for everyone."""

    errors = preflight_validate_registries(
        builtin_scheme.build_arch_registry(),
        builtin_scheme.build_env_registry(),
        additional_known_style_ids=set(builtin_scheme.build_arch_registry().values()),
    )
    assert any("page_layout" in error for error in errors)


def test_digest_is_stable_across_calls_and_moves_with_the_scheme() -> None:
    first = builtin_scheme.scheme_digest()
    assert first == builtin_scheme.scheme_digest()
    assert len(first) == 64

    original = builtin_scheme._level_indent
    try:
        builtin_scheme._level_indent = lambda level: (99, 99)
        assert builtin_scheme.scheme_digest() != first
    finally:
        builtin_scheme._level_indent = original
    assert builtin_scheme.scheme_digest() == first


@pytest.mark.parametrize("part", ["numbering", "styles"])
def test_generated_parts_are_well_formed(part: str) -> None:
    import xml.etree.ElementTree as ET

    builder = {
        "numbering": builtin_scheme.build_numbering_xml,
        "styles": builtin_scheme.build_styles_xml,
    }[part]
    ET.fromstring(builder().encode("utf-8"))
