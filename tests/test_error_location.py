"""The published location of a fail-closed decision.

Every engine remediation sentence is fixed, so the most it can say is "the
reported paragraph". Nothing reported it: the engine built a Word-findable
placement for its developer detail and the redaction boundary discarded that
detail wholesale, leaving a user with a stable error code and no way to find
the paragraph it was about.

These tests hold both halves of the repair. The placement must reach the
artifacts, and it must stay incapable of carrying document text -- a location
that could quote the document would be a redaction hole, not a diagnostic.
"""

from __future__ import annotations

import pytest

from spec_formatter import builtin_scheme
from spec_formatter.pipeline import describe_error_location
from spec_formatter.role_contract import ROLE_TO_ARCH_STYLE
from spec_formatter.style_application.core.canadian_to_csi import plan_canadian_to_csi
from spec_formatter.style_application.core.errors import (
    EngineError,
    ErrorLocation,
    safe_error_location,
)

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"

#: Distinctive body text. Nothing derived from these paragraphs may appear in a
#: published location, however the failure is reported.
SECRET_BODY = "Provide listed CPVC sprinkler piping per the confidential basis of design."


def _styled(role: str, text: str) -> str:
    return (
        f'<w:p><w:pPr><w:pStyle w:val="{ROLE_TO_ARCH_STYLE[role]}"/></w:pPr>'
        f'<w:r><w:t xml:space="preserve">{text}</w:t></w:r></w:p>'
    )


def _document(paragraphs: list[str]) -> str:
    return (
        f'<w:document xmlns:w="{W_NS}"><w:body>{"".join(paragraphs)}</w:body></w:document>'
    )


def _classifications(roles: list[str]) -> dict:
    return {
        "classifications": [
            {"paragraph_index": index, "csi_role": role}
            for index, role in enumerate(roles)
        ]
    }


def _plan(paragraphs: list[str], roles: list[str], **kwargs):
    return plan_canadian_to_csi(
        _document(paragraphs),
        builtin_scheme.build_styles_xml(),
        _classifications(roles),
        numbering_xml=kwargs.pop("numbering_xml", builtin_scheme.build_numbering_xml()),
        **kwargs,
    )


# --------------------------------------------------------------------------
# The type itself
# --------------------------------------------------------------------------


def test_a_section_number_must_survive_the_one_grammar() -> None:
    """The gate that makes a location safe to persist.

    A location reaches ``run.json`` verbatim, so the only free-form field it
    has is re-validated through the same section-number grammar the rest of
    the application uses. Body text cannot pass it.
    """

    with pytest.raises(ValueError, match="canonical section number"):
        ErrorLocation(
            paragraph_index=3,
            section_number=SECRET_BODY,
            section_state="numbered",
        )
    assert (
        ErrorLocation(
            paragraph_index=3, section_number="21 13 13", section_state="numbered"
        ).section_number
        == "21 13 13"
    )


@pytest.mark.parametrize(
    "kwargs",
    [
        {"paragraph_index": -1},
        {"paragraph_index": True},
        {"paragraph_index": "3"},
        {"paragraph_index": 3, "placement": "sideways"},
        {"paragraph_index": 3, "section_state": "maybe"},
        {"paragraph_index": 3, "heading_ordinal": -2},
        # A number with no state to hang it on, and a state with no number.
        {"paragraph_index": 3, "section_number": "21 13 13"},
        {"paragraph_index": 3, "section_state": "numbered"},
    ],
)
def test_a_location_refuses_anything_but_validated_scalars(kwargs: dict) -> None:
    with pytest.raises(ValueError):
        ErrorLocation(**kwargs)


def test_only_a_real_location_can_be_published() -> None:
    """``safe_error_location`` is duck-typed but not gullible.

    Wrapper exceptions forward the attribute without importing the type, so the
    reader has to accept any object -- and therefore has to reject one that
    merely claims the attribute, or an arbitrary dict would reach an artifact.
    """

    class Spoofed(Exception):
        safe_error_location = {"paragraph_index": SECRET_BODY}

    assert safe_error_location(Spoofed()) is None
    assert safe_error_location(Exception()) is None
    located = EngineError(
        "canadian_to_csi_hierarchy",
        "detail",
        ErrorLocation(paragraph_index=7),
    )
    assert safe_error_location(located) == {
        "paragraph_index": 7,
        "section_number": None,
        "heading_ordinal": None,
        "placement": "unknown",
        "section_state": "unknown",
        "description": "paragraph index 7",
    }


def test_a_location_renders_the_same_sentence_everywhere() -> None:
    """One rendering, so run.log, run.json and the GUI cannot disagree."""

    location = ErrorLocation(
        paragraph_index=119,
        section_number="21 13 13",
        heading_ordinal=5,
        placement="after_heading",
        section_state="numbered",
    )
    expected = "Section 21 13 13, after heading 5, paragraph index 119"
    assert location.describe() == expected
    assert location.as_dict()["description"] == expected
    assert describe_error_location(location.as_dict()) == expected
    # The message suffix omits the index, because the sentence it is appended
    # to already names the paragraph.
    assert location.describe_position() == "Section 21 13 13, after heading 5"


def test_describe_error_location_tolerates_a_missing_or_odd_payload() -> None:
    assert describe_error_location(None) == ""
    assert describe_error_location({}) == ""
    assert describe_error_location({"description": 5}) == ""
    assert describe_error_location("not a payload") == ""


# --------------------------------------------------------------------------
# What the reverse converter reports
# --------------------------------------------------------------------------


def test_an_unprovable_list_is_reported_at_its_first_paragraph() -> None:
    """The failure is about a list level; the answer has to be a paragraph.

    "Some level of some list starts at 4" is not something anyone can act on
    in Word. The level is reported at the first paragraph sitting on it, which
    is a place a user can actually open.
    """

    rows = [
        ("SectionID", "SECTION 21 13 13"),
        ("PART", "GENERAL"),
        ("ARTICLE", "SUMMARY"),
        ("PARAGRAPH", SECRET_BODY),
    ]
    numbering = builtin_scheme.build_numbering_xml().replace(
        '<w:start w:val="1"/>', '<w:start w:val="4"/>', 1
    )
    with pytest.raises(EngineError) as raised:
        _plan(
            [_styled(role, text) for role, text in rows],
            [role for role, _text in rows],
            numbering_xml=numbering,
        )

    # Fix 2: a reverse-mode failure reports the reverse code. The shared
    # numbering helpers used to hard-code the forward one, so this run was
    # told to fix its Canadian conversion.
    assert raised.value.code == "canadian_to_csi_numbering_unprovable"
    location = raised.value.location
    assert location is not None
    assert location.section_number == "21 13 13"
    assert location.paragraph_index == 1
    assert location.describe() == "Section 21 13 13, heading 1, paragraph index 1"
    assert SECRET_BODY not in str(location.as_dict())


def test_a_paragraph_left_on_the_converted_list_names_that_paragraph() -> None:
    """The check a spec with a stray blank list item trips."""

    rows = [
        ("SectionID", "SECTION 21 13 13"),
        ("PART", "GENERAL"),
        ("ARTICLE", "SUMMARY"),
        ("PARAGRAPH", "Classified."),
    ]
    paragraphs = [_styled(role, text) for role, text in rows]
    paragraphs.append(_styled("PARAGRAPH", SECRET_BODY))

    with pytest.raises(EngineError) as raised:
        _plan(paragraphs, [role for role, _text in rows])

    assert raised.value.code == "canadian_to_csi_hierarchy"
    location = raised.value.location
    assert location is not None
    assert location.paragraph_index == 4
    assert location.section_number == "21 13 13"
    assert SECRET_BODY not in str(location.as_dict())


def test_a_tracked_paragraph_mark_names_that_paragraph() -> None:
    rows = [
        ("SectionID", "SECTION 21 13 13"),
        ("PART", "GENERAL"),
        ("ARTICLE", "SUMMARY"),
        ("ARTICLE", "REFERENCES"),
    ]
    paragraphs = [
        f'<w:p><w:pPr><w:pStyle w:val="{ROLE_TO_ARCH_STYLE[role]}"/>'
        + (
            '<w:rPr><w:ins w:id="9" w:author="A" w:date="2026-01-01T00:00:00Z"/></w:rPr>'
            if index == 2
            else ""
        )
        + f'</w:pPr><w:r><w:t xml:space="preserve">{text}</w:t></w:r></w:p>'
        for index, (role, text) in enumerate(rows)
    ]

    with pytest.raises(EngineError) as raised:
        _plan(paragraphs, [role for role, _text in rows])

    assert raised.value.code == "canadian_to_csi_tracked_hierarchy"
    assert raised.value.location is not None
    assert raised.value.location.paragraph_index == 2


def test_the_developer_message_still_carries_its_locator_suffix() -> None:
    """The structured location is additive; the message shape is unchanged.

    Both come from one set of index tables, so a message and an artifact that
    disagreed about which paragraph failed would be worse than either alone.
    """

    rows = [
        ("SectionID", "SECTION 21 13 13"),
        ("PART", "GENERAL"),
        ("ARTICLE", "SUMMARY"),
        ("PARAGRAPH", "Classified."),
    ]
    paragraphs = [_styled(role, text) for role, text in rows]
    paragraphs.append(_styled("PARAGRAPH", SECRET_BODY))

    with pytest.raises(EngineError) as raised:
        _plan(paragraphs, [role for role, _text in rows])

    detail = str(raised.value)
    assert "(Section 21 13 13, after heading 3)" in detail
    assert raised.value.location.describe_position() == (
        "Section 21 13 13, after heading 3"
    )


def test_a_diagnostics_event_carries_the_failing_paragraph_index() -> None:
    """``diagnostics.jsonl`` is where a developer looks first.

    The index is an int, so it survives the diagnostics field boundary that
    drops anything which could hold document text.
    """

    from spec_formatter.diagnostics import DEBUG, DiagnosticsRecorder

    recorder = DiagnosticsRecorder(min_level=DEBUG)
    with pytest.raises(EngineError):
        with recorder.timer("target", "canadian_to_csi", target=1):
            raise EngineError(
                "canadian_to_csi_hierarchy",
                "detail naming " + SECRET_BODY,
                ErrorLocation(
                    paragraph_index=119,
                    section_number="21 13 13",
                    heading_ordinal=5,
                    placement="after_heading",
                    section_state="numbered",
                ),
            )

    event = recorder.iter_dicts()[-1]
    assert event["level"] == "ERROR"
    assert event["fields"]["paragraph_index"] == 119
    # Only the index crosses; the section number has whitespace and the
    # diagnostics boundary drops it rather than truncating.
    assert SECRET_BODY not in str(event)
    assert "21 13 13" not in str(event)
