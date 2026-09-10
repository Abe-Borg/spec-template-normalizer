"""Stable engine error codes with fixed, text-free remediation sentences.

Engine failures used to reach the user as ``[untrusted detail omitted;
sha256=...]`` because their messages can echo document text and the
pipeline's redaction boundary rightly refuses to persist them. An
:class:`EngineError` keeps the detailed message for developers (``str(exc)``
is unchanged, so existing ``match=`` tests still hold) and carries a stable
``code`` plus a fixed remediation sentence on the ``safe_error_code`` /
``safe_error_message`` attributes that :func:`spec_formatter.pipeline.
safe_error_diagnostic` already passes through. The code and sentence are what
``run.json``, each ``audit.json``, and the GUI show.

Keep the code set small and closed: a new code is a contract change that the
guide's "Error codes and stages" section must list.

A remediation sentence is fixed, so it can only say "the reported paragraph".
:class:`ErrorLocation` is what reports it. The engine already builds a
Word-findable locator for its developer detail -- a section number and two
counts, never body text -- and that detail is discarded at the redaction
boundary along with everything else the message might echo. The location
therefore travels as its own validated, scalar-only value rather than as
prose: it is the one part of a failure a user can act on, and dropping it
left every "check the reported paragraph" sentence pointing at nothing.
"""

from __future__ import annotations

from dataclasses import dataclass
from typing import Any, Mapping, Optional

from .section_numbers import section_number_display_form

ERROR_REMEDIATIONS: Mapping[str, str] = {
    "header_footer_target_section_id_required": (
        "The architect header/footer names its section number, but this target "
        "has no recognisable SECTION number to put in its place. Add or fix the "
        "target's SECTION line, or use an architect template without a section "
        "number in its header/footer."
    ),
    "header_footer_target_section_title_required": (
        "The architect header/footer names its section title, but this target "
        "has no section title to put in its place. Add the target's section "
        "title line below its SECTION number."
    ),
    "header_footer_token_residual": (
        "After substitution the imported header/footer still carried the "
        "architect's section number or title, so the output was withheld. "
        "Simplify the architect header/footer text boxes or report the layout."
    ),
    "canadian_architect_contract": (
        "The architect template's PART, article, and list roles do not form the "
        "Word numbering contract Canadian conversion requires. Fix the architect "
        "template's automatic numbering, then re-analyze it."
    ),
    "canadian_target_hierarchy": (
        "The target's PART, article, and paragraph markers do not form a "
        "hierarchy Canadian conversion can prove. Check the target's heading "
        "and list markers around the reported paragraph."
    ),
    "canadian_target_markup": (
        "A target paragraph carries markup inside its numbering marker that "
        "Canadian conversion cannot edit safely. Retype the marker as plain "
        "text or remove the nested formatting."
    ),
    "canadian_numbering_unprovable": (
        "A Word numbering list starts, restarts, or overrides its counter in a "
        "way the converter cannot prove. Use numbering that starts at 1 without "
        "level overrides or explicit restarts."
    ),
    "canadian_to_csi_hierarchy": (
        "The target's Canadian PART, article, and list levels do not form a "
        "hierarchy that can be rewritten as CSI markers. Check the heading and "
        "list levels around the reported paragraph."
    ),
    "canadian_to_csi_numbering_unprovable": (
        "A paragraph's Canadian number could not be proven from the target's "
        "own numbering, so no CSI marker could be written for it. Check that "
        "the paragraph is a real list item with numbering that starts at 1."
    ),
    "canadian_to_csi_tracked_hierarchy": (
        "A heading or list paragraph is itself an unresolved tracked insertion "
        "or deletion, so its CSI number depends on whether that revision is "
        "later accepted or rejected. Accept or reject the tracked changes on "
        "the reported paragraph, then convert."
    ),
    "conversion_prediction_mismatch": (
        "The converted document did not match the set of numbering edits the "
        "converter committed to before writing, so the output was withheld. "
        "This is a defect in the application rather than in your document; "
        "please report the target."
    ),
    "geometry_not_preserved": (
        "A paragraph's effective indentation changed during formatting, so the "
        "output was withheld. This is usually a template whose list styles take "
        "their indents from the numbering definition; please report the target "
        "and the reported paragraph index."
    ),
    "builtin_scheme_contract": (
        "The built-in Canadian CSC PageFormat scheme failed its own contract "
        "check. This is a defect in the application rather than in your "
        "document; please report it."
    ),
    "classification_invalid_payload": (
        "The classification result was malformed (an unknown, duplicate, or "
        "invalid paragraph disposition). Run the target again; if it repeats, "
        "report the target."
    ),
    "classification_deterministic_override": (
        "The classifier tried to override a paragraph whose role was proven "
        "deterministically. Run the target again; if it repeats, report the "
        "target."
    ),
    "classification_coverage_incomplete": (
        "Not every classifiable paragraph received exactly one disposition. Run "
        "the target again; if it repeats, report the target."
    ),
    "numbering_importer_unavailable": (
        "The architect styles need numbering definitions but the numbering "
        "importer is not installed with this build. Reinstall the application."
    ),
    "template_section_shell_conflict": (
        "Architect template has conflicting section shells; use one canonical "
        "page layout and default/even/first header-footer mapping."
    ),
    "template_default_section_conflict": (
        "Architect template default section conflicts with its section chain."
    ),
    "template_duplicate_section_index": (
        "Architect template section chain has duplicate section_index values."
    ),
}

#: The closed set of engine stages recorded on ``BatchResult.stage`` /
#: ``TargetFormatResult.stage`` by the shared application path, in order.
ENGINE_STAGES: tuple[str, ...] = (
    "classification_ready",
    "disposition_verification",
    "application_policy",
    "classification_checkpoint",
    "source_catalog_snapshot",
    "target_token_extraction",
    "csi_conversion",
    "canadian_to_csi_conversion",
    "canadian_classification_mapping",
    "environment_application",
    "header_footer_token_patch",
    "numbering_import",
    "header_footer_numbering_remap",
    "style_import",
    "header_footer_style_remap",
    "stability_snapshot",
    "classification_application",
    "stability_verification",
    "application_reporting",
    "output_publication",
    "complete",
)

#: Stages the target runner records before the shared path is entered.
RUNNER_STAGES: tuple[str, ...] = (
    "validation",
    "extraction",
    "bundle_build",
    "classification_preflight",
    "classification",
    "application",
)

#: Stages the pipeline records around the runner.
PIPELINE_STAGES: tuple[str, ...] = (
    "not_started",
    "processing",
    "publication",
    "complete",
)


#: Where a paragraph sits relative to the headings of its section. Closed, so
#: a consumer can branch on it without parsing the rendered sentence.
ERROR_LOCATION_PLACEMENTS: tuple[str, ...] = (
    "heading",
    "after_heading",
    "before_first_heading",
    "unknown",
)

#: What is known about the SECTION line the paragraph belongs to. ``none``
#: means the paragraph precedes every SECTION line, which is different from
#: ``unnumbered`` (a SECTION line whose number the grammar did not recognise)
#: and from ``unknown`` (the failure had no role map to resolve one).
ERROR_LOCATION_SECTION_STATES: tuple[str, ...] = (
    "numbered",
    "unnumbered",
    "none",
    "unknown",
)


@dataclass(frozen=True)
class ErrorLocation:
    """Where a fail-closed decision happened, in terms a user can act on.

    Deliberately scalars and one grammar-validated section number. A paragraph
    index alone is useless -- nobody can find ``word/document.xml`` paragraph
    32 in Word -- and free prose cannot be persisted, because a message that
    quotes the document is exactly what the redaction boundary exists to stop.
    A section number plus a heading ordinal is findable *and* provably not body
    text, so it is safe to write into ``run.json``, ``audit.json``, ``run.log``
    and the GUI verbatim.

    ``section_number`` is re-validated here through the one section-number
    grammar rather than trusted from the caller. The producer already resolves
    it that way; validating again is what lets this type promise that nothing
    which reaches an artifact through it can be arbitrary document text.
    """

    paragraph_index: int
    section_number: Optional[str] = None
    heading_ordinal: Optional[int] = None
    placement: str = "unknown"
    section_state: str = "unknown"

    def __post_init__(self) -> None:
        index = self.paragraph_index
        if isinstance(index, bool) or not isinstance(index, int) or index < 0:
            raise ValueError(f"paragraph_index must be a non-negative int: {index!r}")
        if self.placement not in ERROR_LOCATION_PLACEMENTS:
            raise ValueError(f"unknown error location placement: {self.placement!r}")
        if self.section_state not in ERROR_LOCATION_SECTION_STATES:
            raise ValueError(
                f"unknown error location section state: {self.section_state!r}"
            )
        ordinal = self.heading_ordinal
        if ordinal is not None and (
            isinstance(ordinal, bool) or not isinstance(ordinal, int) or ordinal < 0
        ):
            raise ValueError(f"heading_ordinal must be a non-negative int: {ordinal!r}")
        number = self.section_number
        if number is not None:
            if not isinstance(number, str) or section_number_display_form(number) != number:
                raise ValueError(
                    "section_number must be a canonical section number, not free text"
                )
            if self.section_state != "numbered":
                raise ValueError(
                    "section_number is only meaningful with section_state='numbered'"
                )
        elif self.section_state == "numbered":
            raise ValueError("section_state='numbered' requires a section_number")

    def describe_position(self) -> str:
        """Render the placement alone, without the paragraph index.

        This is the suffix engine messages append to a sentence that already
        names the paragraph by index, so repeating it here would read as
        ``Paragraph 32 (Section 21 13 13, heading 5, paragraph index 32)``.
        """

        if self.section_state == "numbered":
            where = f"Section {self.section_number}"
        elif self.section_state == "unnumbered":
            where = "an unnumbered SECTION line"
        elif self.section_state == "none":
            where = "before any SECTION line"
        else:
            where = ""
        if self.placement == "heading" and self.heading_ordinal:
            place = f"heading {self.heading_ordinal}"
        elif self.placement == "after_heading" and self.heading_ordinal:
            place = f"after heading {self.heading_ordinal}"
        elif self.placement == "before_first_heading":
            place = "before its first heading"
        else:
            place = ""
        return ", ".join(part for part in (where, place) if part)

    def describe(self) -> str:
        """Render the one-line human form used in ``run.log`` and the GUI."""

        position = self.describe_position()
        index = f"paragraph index {self.paragraph_index}"
        return f"{position}, {index}" if position else index

    def as_dict(self) -> dict:
        """JSON-safe payload for artifacts, including the rendered sentence."""

        return {
            "paragraph_index": self.paragraph_index,
            "section_number": self.section_number,
            "heading_ordinal": self.heading_ordinal,
            "placement": self.placement,
            "section_state": self.section_state,
            "description": self.describe(),
        }


class EngineError(ValueError):
    """A fail-closed engine decision with a stable code and remediation."""

    def __init__(
        self,
        code: str,
        message: str,
        location: Optional[ErrorLocation] = None,
    ) -> None:
        if code not in ERROR_REMEDIATIONS:
            raise KeyError(f"unknown engine error code: {code!r}")
        if location is not None and not isinstance(location, ErrorLocation):
            raise TypeError(f"location must be an ErrorLocation: {location!r}")
        super().__init__(message)
        self.code = code
        self.safe_error_code = code
        self.safe_error_message = ERROR_REMEDIATIONS[code]
        self.location = location
        self.safe_error_location = location


def attach_engine_error(
    error: BaseException,
    code: str,
    location: Optional[ErrorLocation] = None,
) -> BaseException:
    """Give an existing exception (any type) the engine code and remediation."""

    if code not in ERROR_REMEDIATIONS:
        raise KeyError(f"unknown engine error code: {code!r}")
    if location is not None and not isinstance(location, ErrorLocation):
        raise TypeError(f"location must be an ErrorLocation: {location!r}")
    setattr(error, "safe_error_code", code)
    setattr(error, "safe_error_message", ERROR_REMEDIATIONS[code])
    if location is not None:
        setattr(error, "safe_error_location", location)
    return error


def safe_error_location(error: Any) -> Optional[dict]:
    """Return an exception's location as an artifact payload, if it carries one.

    Duck-typed like ``safe_error_code``/``safe_error_message`` so a wrapper
    exception that forwards the attribute needs no import, and so a foreign
    object claiming the attribute cannot smuggle a payload: only a real
    :class:`ErrorLocation` is accepted.
    """

    location = getattr(error, "safe_error_location", None)
    return location.as_dict() if isinstance(location, ErrorLocation) else None


def remediation_for(code: str) -> str | None:
    return ERROR_REMEDIATIONS.get(code)


__all__ = [
    "ENGINE_STAGES",
    "ERROR_LOCATION_PLACEMENTS",
    "ERROR_LOCATION_SECTION_STATES",
    "ERROR_REMEDIATIONS",
    "EngineError",
    "ErrorLocation",
    "PIPELINE_STAGES",
    "RUNNER_STAGES",
    "attach_engine_error",
    "remediation_for",
    "safe_error_location",
]
