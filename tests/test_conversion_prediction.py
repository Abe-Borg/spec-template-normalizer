"""Both converters commit to an exact prediction of every paragraph they change.

WI-03 of the DOCX Method Hardening program. A converter states, before it
edits, what each paragraph it changes will read afterwards: the visible text,
and the exact run content (every run's text, tabs, breaks and space
preservation) worked out from the *source* paragraph's run-content signature.
The prediction is then asserted twice: against the converter's own assembled
document, and again by the final gate against the packaged output after
every later stage has run.

Working the prediction out on the signature rather than on the XML is what
keeps it from agreeing with the edit by construction. The expected values in
these tests are written out by hand for the same reason.
"""

from __future__ import annotations

import pytest

from spec_formatter.style_application.core import canadian_to_csi, csi_to_canadian
from spec_formatter.style_application.core.csi_to_canadian import plan_csi_to_canadian
from spec_formatter.style_application.core.errors import EngineError
from spec_formatter.style_application.core.expected_changes import (
    NO_EXPECTED_PARAGRAPH_CHANGES,
    ExpectedParagraphChanges,
)
from tests.test_canadian_to_csi import _TRACKING_ON, _convert, _convert_tracked
from tests.test_csi_to_canadian import (
    _canadian_role_specs,
    _classifications,
    _document,
    _paragraph,
    _source_numbering,
    _styles,
)

PREDICTION_MISMATCH = "conversion_prediction_mismatch"


def _forward(source: str, *roles: str, **kwargs):
    return plan_csi_to_canadian(
        source,
        kwargs.pop("styles_xml", _styles()),
        _classifications(*roles),
        _canadian_role_specs(*roles),
        **kwargs,
    )


# --- The value itself ------------------------------------------------------


def test_no_expected_changes_predicts_nothing() -> None:
    assert dict(NO_EXPECTED_PARAGRAPH_CHANGES.changes) == {}
    assert dict(NO_EXPECTED_PARAGRAPH_CHANGES.roles) == {}
    assert len(NO_EXPECTED_PARAGRAPH_CHANGES) == 0


def test_a_prediction_never_shows_the_document_text_it_holds() -> None:
    # It carries expected text by necessity. Nothing that renders it -- a log
    # line, an assertion message, a traceback -- may show that text.
    plan = _convert([("PART", "GENERAL"), ("ARTICLE", "SUMMARY")])
    expected = plan.expected_paragraph_changes
    assert isinstance(expected, ExpectedParagraphChanges)
    for rendered in (repr(expected), str(expected), repr(expected.changes[0])):
        assert "GENERAL" not in rendered
        assert "SUMMARY" not in rendered


def test_a_prediction_is_immutable() -> None:
    plan = _convert([("PART", "GENERAL")])
    expected = plan.expected_paragraph_changes
    with pytest.raises(TypeError):
        expected.changes[5] = expected.changes[0]  # type: ignore[index]
    with pytest.raises(TypeError):
        expected.roles[5] = "PART"  # type: ignore[index]


# --- csi_to_canadian -------------------------------------------------------


def test_forward_prediction_covers_the_typed_markers_it_removes() -> None:
    plan = _forward(
        _document(
            _paragraph("PART 1 - GENERAL"),
            _paragraph("1.01 SUMMARY"),
            _paragraph("A. Work Included"),
        ),
        "PART",
        "ARTICLE",
        "PARAGRAPH",
    )

    changes = plan.expected_paragraph_changes.changes
    assert sorted(changes) == [0, 1, 2]
    assert [changes[index].visible_text for index in (0, 1, 2)] == [
        "GENERAL",
        "SUMMARY",
        "Work Included",
    ]
    # Every run survives; only the marker's text comes out of it.
    assert changes[2].run_content == ((("t", "Work Included", False),),)
    # Nothing in the forward direction is tracked.
    assert changes[2].run_content_outside_own_revisions == changes[2].run_content
    # The roles let the gate place a failure by SECTION and heading.
    assert dict(plan.expected_paragraph_changes.roles) == {
        0: "PART",
        1: "ARTICLE",
        2: "PARAGRAPH",
    }


def test_forward_prediction_leaves_automatic_numbering_out() -> None:
    # A retargeted automatic list item changes numbering, not text, so it is
    # not in the prediction -- and the gate therefore holds it to identity.
    style = (
        '<w:style w:type="paragraph" w:styleId="TargetAlpha"><w:pPr><w:numPr>'
        '<w:ilvl w:val="0"/><w:numId w:val="7"/>'
        "</w:numPr></w:pPr></w:style>"
    )
    source = _document(_paragraph("Scope", '<w:pStyle w:val="TargetAlpha"/>'))

    plan = _forward(
        source,
        "PARAGRAPH",
        styles_xml=_styles(style),
        numbering_xml=_source_numbering("7"),
    )

    assert plan.report.automatic_numbering_retargeted == 1
    assert dict(plan.expected_paragraph_changes.changes) == {}
    assert dict(plan.expected_paragraph_changes.roles) == {0: "PARAGRAPH"}


def test_forward_prediction_removes_the_delimiter_tab_the_converter_removes() -> None:
    # The marker ends exactly at the end of a text node and a structural tab
    # follows it, in the next run: that tab is the marker's delimiter.
    plan = _forward(
        _document(
            '<w:p><w:r><w:rPr><w:b/></w:rPr><w:t>A</w:t></w:r>'
            "<w:r><w:t>.</w:t><w:tab/></w:r>"
            "<w:r><w:rPr><w:i/></w:rPr><w:t>Scope &amp; coordination</w:t></w:r></w:p>"
        ),
        "PARAGRAPH",
    )

    assert plan.expected_paragraph_changes.changes[0].run_content == (
        (("t", "", False),),
        (("t", "", False),),
        (("t", "Scope & coordination", False),),
    )


def test_forward_prediction_keeps_a_tab_the_separator_runs_past() -> None:
    # "PART 1" ends one node; its "- " separator continues in the next, so the
    # removal ends inside that node and the tab between them stays. This is
    # the converter's existing behaviour, predicted rather than repaired.
    plan = _forward(
        _document(
            "<w:p><w:r><w:t>PART 1</w:t></w:r><w:r><w:tab/></w:r>"
            "<w:r><w:t>- GENERAL</w:t></w:r></w:p>"
        ),
        "PART",
    )

    change = plan.expected_paragraph_changes.changes[0]
    assert change.visible_text == "GENERAL"
    assert change.run_content == (
        (("t", "", False),),
        (("tab",),),
        (("t", "GENERAL", False),),
    )


def test_forward_conversion_is_held_to_its_exact_prediction(monkeypatch) -> None:
    """An edit that only a whitespace-normalized check would pass is refused.

    The converter's own per-paragraph check compares normalized text and the
    text-free skeleton, so a doubled space collapsed inside the text it keeps
    passes both. The exact prediction does not.
    """

    real_remove = csi_to_canadian._remove_literal_marker

    def remove_and_collapse(paragraph_xml: str, role: str):
        edited, tab_removed = real_remove(paragraph_xml, role)
        return edited.replace("Work  Included", "Work Included"), tab_removed

    monkeypatch.setattr(csi_to_canadian, "_remove_literal_marker", remove_and_collapse)

    with pytest.raises(EngineError) as raised:
        _forward(_document(_paragraph("A. Work  Included")), "PARAGRAPH")

    assert raised.value.code == PREDICTION_MISMATCH
    assert raised.value.location is not None
    assert raised.value.location.paragraph_index == 0
    assert "Work" not in str(raised.value)


def test_forward_prediction_does_not_share_the_converters_decoding() -> None:
    """A character the edit rewrites without being asked is caught.

    The converter decodes text nodes with ``html.unescape``, which follows
    HTML5 and turns the character reference ``&#x80;`` into U+20AC, the euro
    sign; XML keeps it as U+0080. Rewriting the node that kept the reference
    therefore put a euro sign in the document, and every check the converter
    had compared the text through the same decoding, so none saw it. The
    prediction is worked out from the source's run-content signature, which
    decodes as an XML parser does.
    """

    with pytest.raises(EngineError) as raised:
        _forward(_document(_paragraph("A. Unit price &#x80; each")), "PARAGRAPH")

    assert raised.value.code == PREDICTION_MISMATCH


# --- canadian_to_csi -------------------------------------------------------


def test_reverse_prediction_covers_every_marked_paragraph() -> None:
    plan = _convert(
        [("PART", "GENERAL"), ("ARTICLE", "SUMMARY"), ("PARAGRAPH", "First requirement.")]
    )

    changes = plan.expected_paragraph_changes.changes
    assert sorted(changes) == [0, 1, 2]
    assert [changes[index].visible_text for index in (0, 1, 2)] == [
        "PART 1 GENERAL",
        "1.1 SUMMARY",
        "A. First requirement.",
    ]
    # Untracked, the marker and its tab join the paragraph's first text run.
    assert changes[0].run_content == (
        (("t", "PART 1", True), ("tab",), ("t", "GENERAL", True)),
    )
    assert changes[0].run_content_outside_own_revisions == changes[0].run_content


def test_reverse_tracked_prediction_is_the_source_outside_the_marker_revision() -> None:
    plan = _convert_tracked([("PART", "GENERAL"), ("ARTICLE", "SUMMARY")], _TRACKING_ON)

    change = plan.expected_paragraph_changes.changes[0]
    # Tracked, the marker is its own run, before the run holding the text...
    assert change.run_content == (
        (("t", "PART 1", True), ("tab",)),
        (("t", "GENERAL", True),),
    )
    # ...and with this application's revisions projected out, the paragraph
    # is exactly the source again.
    assert change.run_content_outside_own_revisions == ((("t", "GENERAL", True),),)


def test_reverse_conversion_is_held_to_its_exact_prediction(monkeypatch) -> None:
    real_insert = canadian_to_csi._insert_marker

    def insert_and_collapse(paragraph_xml: str, marker: str, **kwargs):
        edited = real_insert(paragraph_xml, marker, **kwargs)
        return edited.replace("Provide  listed", "Provide listed")

    monkeypatch.setattr(canadian_to_csi, "_insert_marker", insert_and_collapse)

    with pytest.raises(EngineError) as raised:
        _convert(
            [
                ("PART", "GENERAL"),
                ("ARTICLE", "SUMMARY"),
                ("PARAGRAPH", "Provide  listed sprinklers."),
            ]
        )

    assert raised.value.code == PREDICTION_MISMATCH
    assert raised.value.location is not None
    assert raised.value.location.paragraph_index == 2
    assert "Provide" not in str(raised.value)


def test_reverse_marker_written_without_its_tab_is_refused(monkeypatch) -> None:
    """The prediction is the source plus the planned marker, never the edit.

    ``_verify_marked_paragraph`` accepts a marker that runs straight into the
    text ("PART 1GENERAL"): it checks that the text starts with the marker and
    that the rest, stripped, is the body. A prediction read back off the edited
    paragraph would agree with that edit too. One made from the source before
    the edit expects the tab, and refuses its absence.
    """

    real_insert = canadian_to_csi._insert_marker

    def insert_without_the_tab(paragraph_xml: str, marker: str, **kwargs):
        return real_insert(paragraph_xml, marker, **kwargs).replace("<w:tab/>", "", 1)

    monkeypatch.setattr(canadian_to_csi, "_insert_marker", insert_without_the_tab)

    with pytest.raises(EngineError) as raised:
        _convert([("PART", "GENERAL")])

    assert raised.value.code == PREDICTION_MISMATCH
    assert raised.value.location is not None
    assert raised.value.location.paragraph_index == 0
