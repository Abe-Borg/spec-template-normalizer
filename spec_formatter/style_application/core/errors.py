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
"""

from __future__ import annotations

from typing import Mapping

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
    "geometry_verification",
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


class EngineError(ValueError):
    """A fail-closed engine decision with a stable code and remediation."""

    def __init__(self, code: str, message: str) -> None:
        if code not in ERROR_REMEDIATIONS:
            raise KeyError(f"unknown engine error code: {code!r}")
        super().__init__(message)
        self.code = code
        self.safe_error_code = code
        self.safe_error_message = ERROR_REMEDIATIONS[code]


def attach_engine_error(error: BaseException, code: str) -> BaseException:
    """Give an existing exception (any type) the engine code and remediation."""

    if code not in ERROR_REMEDIATIONS:
        raise KeyError(f"unknown engine error code: {code!r}")
    setattr(error, "safe_error_code", code)
    setattr(error, "safe_error_message", ERROR_REMEDIATIONS[code])
    return error


def remediation_for(code: str) -> str | None:
    return ERROR_REMEDIATIONS.get(code)


__all__ = [
    "ENGINE_STAGES",
    "ERROR_REMEDIATIONS",
    "EngineError",
    "PIPELINE_STAGES",
    "RUNNER_STAGES",
    "attach_engine_error",
    "remediation_for",
]
