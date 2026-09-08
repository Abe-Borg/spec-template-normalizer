"""The one section-number grammar shared by every target-side consumer."""

from __future__ import annotations

import pytest

from spec_formatter.style_application.core.section_numbers import (
    LABELED_SECTION_RE,
    SECTION_HEADING_RE,
    SECTION_NUMBER_RE,
    canonical_section_number,
    render_section_number_like,
    section_number_display_form,
)


@pytest.mark.parametrize(
    ("value", "canonical"),
    [
        ("230500", "230500"),
        ("23 05 00", "230500"),
        ("23 0500", "230500"),
        ("23\t05\t00", "230500"),
        ("23\u00a005\u00a000", "230500"),
        (" 23 05 00 ", "230500"),
        ("23 05 00.13", "230500.13"),
        ("230500.13", "230500.13"),
    ],
)
def test_canonical_section_number_accepts_every_spacing(value, canonical):
    assert canonical_section_number(value) == canonical


@pytest.mark.parametrize(
    "value",
    [
        "",
        None,
        "2305001",
        "23050013",
        "23 05 00 13",
        "230500.1",
        "230500.",
        "23-05-00",
        "SECTION 230500",
        "230500 - PIPING",
    ],
)
def test_canonical_section_number_fails_closed(value):
    assert canonical_section_number(value) == ""


@pytest.mark.parametrize(
    ("value", "display"),
    [
        ("SECTION 23 05 00", "23 05 00"),
        ("Section 230500", "230500"),
        ("SECTION 23  05\u00a000 - PIPING", "23 05 00"),
        ("SECTION 23 05 00.13 PIPING", "23 05 00.13"),
        ("23 05 00 - PIPING", "23 05 00"),
        ("230500", "230500"),
        ("SECTION 2305001", ""),
        ("PIPING 23 05 00", ""),
        ("", ""),
    ],
)
def test_display_form_keeps_the_source_grouping(value, display):
    assert section_number_display_form(value) == display


def test_heading_regex_is_anchored_and_bounded():
    assert SECTION_HEADING_RE.match("SECTION 23 0500").group("number") == "23 0500"
    assert SECTION_HEADING_RE.match("  section 230500.13").group("number") == "230500.13"
    assert SECTION_HEADING_RE.match("SECTION 230500A") is None
    assert SECTION_HEADING_RE.match("SECTION 230500.1") is None
    assert SECTION_HEADING_RE.match("See SECTION 230500") is None


def test_labeled_and_bare_regexes_find_numbers_inside_text():
    text = (
        "See SECTION 23 05 00.13 and Section 260513, not SECTION 2305001, "
        "SECTION 230500A, or 1230500."
    )
    assert [m.group("number") for m in LABELED_SECTION_RE.finditer(text)] == [
        "23 05 00.13",
        "260513",
    ]
    assert [m.group("number") for m in SECTION_NUMBER_RE.finditer(text)] == [
        "23 05 00.13",
        "260513",
    ]


@pytest.mark.parametrize(
    ("source_form", "target", "rendered"),
    [
        ("23 05 00", "260513", "26 05 13"),
        ("230500", "260513", "260513"),
        ("23 0500", "260513", "26 0513"),
        ("23 05 00.13", "260513.23", "26 05 13.23"),
        ("23 05 00", "260513.23", "260513.23"),
        ("", "260513", "260513"),
        ("23 05 00", "", ""),
    ],
)
def test_render_like_source_grouping(source_form, target, rendered):
    assert render_section_number_like(source_form, target) == rendered
