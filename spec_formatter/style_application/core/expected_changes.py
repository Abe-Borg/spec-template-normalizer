"""Identity except the enumerated diff: what a conversion predicts it changes.

The final gate in ``phase2_invariants`` used to prove body text only for
Format-only. The conversion modes verified their own edits in memory, at the
conversion stage, and every stage after that -- environment application,
numbering and style import, classification application, repackaging -- could
change a paragraph's text and still publish.

Each converter now states, before it edits, what every paragraph it changes
will read afterwards: an :class:`ExpectedParagraphChange` per paragraph, keyed
by its index in ``word/document.xml``. The gate then holds

* every paragraph *not* in the prediction to its exact run content
  (:func:`~.xml_helpers.paragraph_run_content_signature`), and
* every paragraph *in* it to that prediction: its visible text, its exact run
  content, and its run content outside this application's own tracked
  revisions.

Format-only predicts nothing, so the same comparison holds every paragraph to
identity, which is the check WI-02 added.

**A prediction is worked out, never read back.** A converter computes a
paragraph's expected run content from the *source* paragraph's signature and
the edit it plans -- a marker's text and delimiter removed, a marker and its
tab added -- on the signature, not on the XML the edit writes. A check derived
from what the edit did would agree with the edit by construction; this one can
disagree with it.

A prediction holds document text, because it has to. It is never published:
it travels from the converter to the gate as a value and nowhere else, and it
renders without its contents.
"""

from __future__ import annotations

import re
from dataclasses import dataclass, field
from types import MappingProxyType
from typing import Dict, Mapping, Optional, Sequence, Tuple

from .errors import EngineError
from .xml_helpers import (
    RunContentSignature,
    paragraph_run_content_signature,
    paragraph_text_from_block,
    run_content_difference,
)


#: The engine error code for any disagreement between a conversion's
#: prediction and the document, in the converter or at the final gate.
PREDICTION_MISMATCH = "conversion_prediction_mismatch"

#: Author recorded on marker insertions when the source is under review.
#:
#: Deliberately *not* the document author. Three things depend on it being
#: distinguishable: Word's markup pane should show a machine conversion apart
#: from a person's own edits; the invariants project the app's revisions back
#: out by author; and a reviewer validating that nothing changed the author's
#: content outside a revision must not have that check pass trivially because
#: the app signed the author's name to its own work.
MARKER_REVISION_AUTHOR = "Specification Formatter"

_OWN_REVISION_RX = re.compile(
    rf'<w:ins\b[^>]*w:author="{re.escape(MARKER_REVISION_AUTHOR)}"[^>]*'
    r"(?:/>|>.*?</w:ins>)",
    re.S,
)


def without_own_revisions(paragraph_xml: str) -> str:
    """Drop insertions this application authored, keeping everyone else's.

    Scoped by author on purpose: a reviewer's pending edits must stay in any
    comparison made through this projection, because losing content inside
    one of those is exactly as damaging as losing it anywhere else.
    """

    return _OWN_REVISION_RX.sub("", paragraph_xml)


@dataclass(frozen=True, repr=False)
class ExpectedParagraphChange:
    """What one converted paragraph must read once the whole run is over.

    ``visible_text`` is the paragraph as :func:`paragraph_text_from_block`
    reads it. ``run_content`` is its exact run-content signature, every run
    included. ``run_content_outside_own_revisions`` is the signature once this
    application's own tracked insertions are projected out: for an untracked
    edit that is ``run_content`` itself, and for a tracked marker it is the
    source paragraph, untouched.
    """

    visible_text: str
    run_content: RunContentSignature
    run_content_outside_own_revisions: RunContentSignature

    def __repr__(self) -> str:
        return (
            f"ExpectedParagraphChange(<{len(self.run_content)} run(s); "
            "text withheld>)"
        )


@dataclass(frozen=True, repr=False)
class ExpectedParagraphChanges:
    """Every paragraph a conversion changes, keyed by paragraph index.

    ``roles`` carries the classified role of each paragraph the converter
    placed, so a failure found later can be reported by SECTION and heading
    rather than by a bare index. Both mappings are read-only copies.
    """

    changes: Mapping[int, ExpectedParagraphChange] = field(default_factory=dict)
    roles: Mapping[int, str] = field(default_factory=dict)

    def __post_init__(self) -> None:
        changes = dict(self.changes)
        for index, change in changes.items():
            _require_paragraph_index(index)
            if not isinstance(change, ExpectedParagraphChange):
                raise TypeError(
                    f"predicted change for paragraph {index} is not an "
                    "ExpectedParagraphChange"
                )
        roles = dict(self.roles)
        for index, role in roles.items():
            _require_paragraph_index(index)
            if not isinstance(role, str):
                raise TypeError(f"role for paragraph {index} is not a string")
        object.__setattr__(self, "changes", MappingProxyType(changes))
        object.__setattr__(self, "roles", MappingProxyType(roles))

    def __repr__(self) -> str:
        return (
            f"ExpectedParagraphChanges(<{len(self.changes)} predicted "
            "paragraph(s); text withheld>)"
        )


def _require_paragraph_index(index: object) -> None:
    if isinstance(index, bool) or not isinstance(index, int) or index < 0:
        raise ValueError(f"paragraph index must be a non-negative int: {index!r}")


#: The prediction of a run that changes no paragraph: Format-only's, and the
#: meaning of a prediction nobody supplied.
NO_EXPECTED_PARAGRAPH_CHANGES = ExpectedParagraphChanges()


@dataclass(frozen=True)
class PredictionMismatch:
    """How one paragraph departed from its prediction.

    ``check`` is the assertion that failed; ``kind``, when there is one, is
    the :data:`~.xml_helpers.RUN_CONTENT_DIFFERENCE_KINDS` member naming what
    differed. Neither can carry document text.
    """

    check: str
    kind: Optional[str] = None


#: Every ``PredictionMismatch.check``: an unpredicted paragraph is tested for
#: its visible text and then its run content, a predicted one for the three
#: after it, in this order.
PREDICTION_CHECKS = (
    "unpredicted_visible_text",
    "unpredicted_change",
    "visible_text",
    "run_content",
    "own_revisions",
)


def prediction_mismatch(
    before_block: str,
    after_block: str,
    change: Optional[ExpectedParagraphChange],
) -> Optional[PredictionMismatch]:
    """How ``after_block`` departs from what was predicted for it, or ``None``.

    With no prediction the paragraph must still read as it did and keep its
    exact run content. Both are needed: the signature records runs, not the
    container a run sits in, so a run wrapped in ``w:del`` or ``w:moveFrom``
    keeps its signature while Word stops showing its text -- which the
    visible-text reading, dropping exactly those containers, does see.

    With a prediction the paragraph must read as predicted, then match the
    predicted run content exactly, and then match it again once this
    application's own tracked revisions are projected out -- which is what
    tells a marker still inside its revision from one that became permanent
    text, since the ``w:ins`` wrapper is not run content.
    """

    if change is None:
        if before_block == after_block:
            return None
        if paragraph_text_from_block(before_block) != paragraph_text_from_block(
            after_block
        ):
            return PredictionMismatch("unpredicted_visible_text")
        kind = run_content_difference(
            paragraph_run_content_signature(before_block),
            paragraph_run_content_signature(after_block),
        )
        return None if kind is None else PredictionMismatch("unpredicted_change", kind)

    if paragraph_text_from_block(after_block) != change.visible_text:
        return PredictionMismatch("visible_text")
    actual = paragraph_run_content_signature(after_block)
    kind = run_content_difference(change.run_content, actual)
    if kind is not None:
        return PredictionMismatch("run_content", kind)
    projected = without_own_revisions(after_block)
    outside = actual if projected == after_block else paragraph_run_content_signature(projected)
    kind = run_content_difference(change.run_content_outside_own_revisions, outside)
    if kind is not None:
        return PredictionMismatch("own_revisions", kind)
    return None


def first_prediction_mismatch(
    before_blocks: Sequence[str],
    after_blocks: Sequence[str],
    expected: ExpectedParagraphChanges,
    progress: Optional[Dict[str, int]] = None,
) -> Optional[Tuple[int, PredictionMismatch]]:
    """The first paragraph that departs from ``expected``, with how it does.

    Paragraphs are paired by index; the caller has already proven the counts
    equal. A predicted paragraph is always compared. An unpredicted one is
    compared only when its XML changed, because identical XML has identical
    run content.

    ``progress`` receives ``body_signature_paragraphs_compared`` and
    ``body_paragraphs_expected_changed`` whether the comparison passes, finds
    a mismatch or raises, so a failed run can show the check ran.
    """

    compared = 0
    try:
        for index, (before, after) in enumerate(zip(before_blocks, after_blocks)):
            change = expected.changes.get(index)
            if change is None and before == after:
                continue
            compared += 1
            mismatch = prediction_mismatch(before, after, change)
            if mismatch is not None:
                return index, mismatch
        return None
    finally:
        record_body_check(progress, expected, compared)


def record_body_check(
    progress: Optional[Dict[str, int]],
    expected: ExpectedParagraphChanges,
    compared: int = 0,
) -> None:
    """Write the body check's counters into ``progress``, when there is one.

    Called as the body check starts, with nothing compared yet, and again by
    :func:`first_prediction_mismatch` with the real count. A check that fails
    before comparing anything -- a paragraph added or removed, a Format-only
    word changed -- has still run, and a failed run must be able to show it.
    """

    if progress is not None:
        progress["body_signature_paragraphs_compared"] = compared
        progress["body_paragraphs_expected_changed"] = len(expected.changes)


def describe_prediction_mismatch(
    index: int,
    where: str,
    mismatch: PredictionMismatch,
) -> str:
    """The developer detail for a mismatch: index, placement and kind only."""

    if mismatch.check == "unpredicted_visible_text":
        return (
            f"Paragraph {index}{where} no longer reads as it did, but the "
            "conversion predicted no change to it."
        )
    if mismatch.check == "unpredicted_change":
        return (
            f"Paragraph {index}{where} changed its run content ({mismatch.kind}), "
            "but the conversion predicted no change to it."
        )
    if mismatch.check == "visible_text":
        return f"Paragraph {index}{where} does not read as the conversion predicted."
    if mismatch.check == "run_content":
        return (
            f"Paragraph {index}{where} changed beyond its predicted conversion "
            f"({mismatch.kind})."
        )
    return (
        f"Paragraph {index}{where} is not its predicted conversion outside this "
        f"application's own tracked revisions ({mismatch.kind})."
    )


def require_predicted_paragraphs_exist(
    expected: ExpectedParagraphChanges,
    paragraph_count: int,
) -> None:
    """Refuse a prediction about a paragraph the document does not have.

    Such a prediction could never be checked, so accepting it would let the
    paragraph it meant go unverified.
    """

    missing = sorted(index for index in expected.changes if index >= paragraph_count)
    if missing:
        raise EngineError(
            PREDICTION_MISMATCH,
            f"The conversion predicted a change to paragraph {missing[0]}, but the "
            f"document has {paragraph_count} body paragraphs.",
        )


__all__ = [
    "MARKER_REVISION_AUTHOR",
    "NO_EXPECTED_PARAGRAPH_CHANGES",
    "PREDICTION_CHECKS",
    "PREDICTION_MISMATCH",
    "ExpectedParagraphChange",
    "ExpectedParagraphChanges",
    "PredictionMismatch",
    "describe_prediction_mismatch",
    "first_prediction_mismatch",
    "prediction_mismatch",
    "record_body_check",
    "require_predicted_paragraphs_exist",
    "without_own_revisions",
]
