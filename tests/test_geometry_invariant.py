"""Coverage for the differential effective-paragraph-geometry invariant.

This is the one check in ``phase2_invariants`` aimed at a defect that leaves no
trace in text: identical wording, identical numbering semantics, identical run
structure and valid XSD, with every paragraph rendering somewhere else. Every
other invariant in that module passed on the document that motivated it.
"""

from __future__ import annotations

import pytest

from spec_formatter.style_application.core.errors import EngineError
from spec_formatter.style_application.phase2_invariants import (
    _verify_effective_paragraph_geometry,
)

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"


def _document(*paragraphs: str) -> str:
    return f'<w:document xmlns:w="{W_NS}"><w:body>{"".join(paragraphs)}</w:body></w:document>'


def _para(ppr: str, text: str = "Body") -> str:
    return f"<w:p><w:pPr>{ppr}</w:pPr><w:r><w:t>{text}</w:t></w:r></w:p>"


def _numbering(level_ind: str = '<w:ind w:left="1259" w:hanging="539"/>') -> str:
    return (
        f'<w:numbering xmlns:w="{W_NS}">'
        f'<w:abstractNum w:abstractNumId="1">'
        f'<w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/>'
        f"<w:pPr>{level_ind}</w:pPr></w:lvl></w:abstractNum>"
        f'<w:num w:numId="4"><w:abstractNumId w:val="1"/></w:num></w:numbering>'
    )


def _styles(style_ppr: str) -> str:
    return (
        f'<w:styles xmlns:w="{W_NS}">'
        f'<w:style w:type="paragraph" w:styleId="Normal"><w:name w:val="Normal"/></w:style>'
        f'<w:style w:type="paragraph" w:styleId="List1"><w:name w:val="List 1"/>'
        f'<w:basedOn w:val="Normal"/><w:pPr>{style_ppr}</w:pPr></w:style>'
        f"</w:styles>"
    )


#: The real shape: the style routes the paragraph to a list and supplies no
#: indent of its own, so the numbering level is the only source of geometry.
_LEVEL_ONLY_STYLES = _styles('<w:numPr><w:ilvl w:val="0"/><w:numId w:val="4"/></w:numPr>')

_IN_LIST = _para('<w:pStyle w:val="List1"/>')
_NUMBERING_CANCELLED = _para(
    '<w:pStyle w:val="List1"/><w:numPr><w:ilvl w:val="0"/><w:numId w:val="0"/></w:numPr>'
)
_NUMBERING_CANCELLED_WITH_IND = _para(
    '<w:pStyle w:val="List1"/><w:numPr><w:ilvl w:val="0"/><w:numId w:val="0"/></w:numPr>'
    '<w:ind w:left="1259" w:hanging="539"/>'
)


def _verify(before: str, after: str, styles: str = _LEVEL_ONLY_STYLES) -> None:
    _verify_effective_paragraph_geometry(
        _document(before), _document(after), styles, styles, _numbering(), _numbering()
    )


def test_cancelling_numbering_without_restoring_indent_fails() -> None:
    """The defect this exists for: the indent leaves with the list."""

    with pytest.raises(EngineError) as raised:
        _verify(_IN_LIST, _NUMBERING_CANCELLED)
    assert raised.value.code == "geometry_not_preserved"
    assert "paragraph index 0" in str(raised.value)


def test_cancelling_numbering_with_the_indent_restored_passes() -> None:
    _verify(_IN_LIST, _NUMBERING_CANCELLED_WITH_IND)


def test_identical_paragraph_and_identical_parts_is_a_pass() -> None:
    """Unchanged everywhere is unchanged.

    Note what this does *not* say: an identical paragraph on its own proves
    nothing, because it resolves through styles, numbering and docDefaults that
    the run may have replaced. It is the parts being identical too that makes
    this one safe -- see the shared-part cases at the end of this module.
    """

    _verify(_IN_LIST, _IN_LIST)


def test_a_restored_indent_that_disagrees_with_the_level_fails() -> None:
    """Restoring the *wrong* geometry is not better than restoring none."""

    wrong = _para(
        '<w:pStyle w:val="List1"/><w:numPr><w:ilvl w:val="0"/><w:numId w:val="0"/></w:numPr>'
        '<w:ind w:left="2880" w:hanging="539"/>'
    )
    with pytest.raises(EngineError) as raised:
        _verify(_IN_LIST, wrong)
    assert raised.value.code == "geometry_not_preserved"


def test_ambiguous_style_and_level_indents_do_not_fail() -> None:
    """Where precedence is ambiguous, the check declines to claim a defect.

    A style that sets ``w:ind`` *and* references a level that sets a different
    one is the case OOXML does not settle: the spec's style hierarchy applies
    paragraph styles after numbering, Word's observed behaviour for a directly
    referenced list is the reverse. Resolving under one reading only would make
    this document fail on a rendering that never moved -- so the check resolves
    under both and fails only when both agree it changed.
    """

    styles = _styles(
        '<w:numPr><w:ilvl w:val="0"/><w:numId w:val="4"/></w:numPr>'
        '<w:ind w:left="999" w:hanging="111"/>'
    )
    # Numbering cancelled, nothing injected: under style-first precedence the
    # style supplied 999 before and still does, so nothing moved.
    _verify(_IN_LIST, _NUMBERING_CANCELLED, styles=styles)


def test_geometry_changing_under_both_readings_still_fails() -> None:
    """Ambiguity is not a blanket exemption."""

    styles = _styles(
        '<w:numPr><w:ilvl w:val="0"/><w:numId w:val="4"/></w:numPr>'
        '<w:ind w:left="999" w:hanging="111"/>'
    )
    moved = _para(
        '<w:pStyle w:val="List1"/><w:numPr><w:ilvl w:val="0"/><w:numId w:val="0"/></w:numPr>'
        '<w:ind w:left="4000"/>'
    )
    with pytest.raises(EngineError) as raised:
        _verify(_IN_LIST, moved, styles=styles)
    assert raised.value.code == "geometry_not_preserved"


def test_direct_indent_is_reported_by_index() -> None:
    """The message names the paragraph, and carries no document text."""

    with pytest.raises(EngineError) as raised:
        _verify_effective_paragraph_geometry(
            _document(_IN_LIST, _IN_LIST),
            _document(_IN_LIST, _NUMBERING_CANCELLED),
            _LEVEL_ONLY_STYLES,
            _LEVEL_ONLY_STYLES,
            _numbering(),
            _numbering(),
        )
    message = str(raised.value)
    assert "paragraph index 1" in message
    assert "Body" not in message


# --- Paragraphs the engine never edited -----------------------------------
#
# A paragraph's own XML being untouched does not prove it still renders where
# it did: it resolves through styles, numbering and docDefaults, all shared.
# Whether that movement is a defect depends on the mode.


def _doc_defaults(ind: str) -> str:
    return (
        f'<w:styles xmlns:w="{W_NS}"><w:docDefaults><w:pPrDefault><w:pPr>{ind}'
        f"</w:pPr></w:pPrDefault></w:docDefaults>"
        f'<w:style w:type="paragraph" w:styleId="Body"><w:name w:val="Body"/></w:style>'
        f"</w:styles>"
    )


_PLAIN = _para('<w:pStyle w:val="Body"/>', "Untouched.")


def test_untouched_paragraph_moved_by_doc_defaults_fails_without_a_shell() -> None:
    """Nothing global should move a paragraph when no shell is being applied.

    The architect-free modes import nothing, so a changed ``docDefaults`` here
    means something moved the target's own geometry -- exactly what this check
    is for, and invisible in the paragraph's own XML.
    """

    with pytest.raises(EngineError) as raised:
        _verify_effective_paragraph_geometry(
            _document(_PLAIN),
            _document(_PLAIN),
            _doc_defaults('<w:ind w:left="0"/>'),
            _doc_defaults('<w:ind w:left="1440"/>'),
            "",
            "",
            applies_document_shell=False,
        )
    assert raised.value.code == "geometry_not_preserved"


def test_untouched_paragraph_moved_by_the_architect_shell_is_expected() -> None:
    """Applying the architect's document-global shell is meant to reflow.

    The guide is explicit that leaving a paragraph's XML unedited proves the
    engine did not touch it, not that it renders identically, and that
    shell-driven reflow of untouched content is expected behaviour. Failing it
    here would refuse every legitimate architect run.
    """

    _verify_effective_paragraph_geometry(
        _document(_PLAIN),
        _document(_PLAIN),
        _doc_defaults('<w:ind w:left="0"/>'),
        _doc_defaults('<w:ind w:left="1440"/>'),
        "",
        "",
        applies_document_shell=True,
    )


def test_an_edited_paragraph_is_checked_even_under_a_shell() -> None:
    """The shell exemption covers untouched paragraphs, not edited ones."""

    with pytest.raises(EngineError) as raised:
        _verify_effective_paragraph_geometry(
            _document(_IN_LIST),
            _document(_NUMBERING_CANCELLED),
            _LEVEL_ONLY_STYLES,
            _LEVEL_ONLY_STYLES,
            _numbering(),
            _numbering(),
            applies_document_shell=True,
        )
    assert raised.value.code == "geometry_not_preserved"


# --- Reporting ------------------------------------------------------------


def test_the_check_reports_its_work_on_success() -> None:
    """An invariant that only speaks when it trips cannot be shown to have run.

    Silence is then indistinguishable from being skipped -- which is precisely
    the state the engine was in before this check existed, and the reason a
    flattened document shipped reporting complete success.
    """

    from spec_formatter.style_application.phase2_invariants import (
        _verify_effective_paragraph_geometry as check,
    )

    compared = check(
        _document(_IN_LIST, _NUMBERING_CANCELLED_WITH_IND),
        _document(_NUMBERING_CANCELLED_WITH_IND, _NUMBERING_CANCELLED_WITH_IND),
        _LEVEL_ONLY_STYLES,
        _LEVEL_ONLY_STYLES,
        _numbering(),
        _numbering(),
    )
    assert compared == 2


def test_paragraph_count_change_reports_zero_rather_than_guessing() -> None:
    """Another invariant owns that failure; this one declines to comment."""

    from spec_formatter.style_application.phase2_invariants import (
        _verify_effective_paragraph_geometry as check,
    )

    assert (
        check(
            _document(_IN_LIST),
            _document(_IN_LIST, _IN_LIST),
            _LEVEL_ONLY_STYLES,
            _LEVEL_ONLY_STYLES,
            _numbering(),
            _numbering(),
        )
        == 0
    )
