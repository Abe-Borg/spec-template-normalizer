"""Explicit, immutable behavior contracts for target application modes."""

from __future__ import annotations

from dataclasses import dataclass

from .conversion_modes import (
    CANADIAN_TO_CSI,
    CSI_TO_CANADIAN,
    CSI_TO_CANADIAN_STANDALONE,
    FORMAT_ONLY,
    validate_conversion_mode,
)


# Contract 4: the architect template became optional for two modes, so
# ``requires_architect_template`` and ``numbering_scheme`` now decide where a
# run's numbering and shell come from.
APPLICATION_POLICY_VERSION = "5"

FORMAT_ONLY_OUTPUT_SUFFIX = "_FORMATTED.docx"
CSI_TO_CANADIAN_OUTPUT_SUFFIX = "_CANADIAN_FORMATTED.docx"
CSI_TO_CANADIAN_STANDALONE_OUTPUT_SUFFIX = "_CANADIAN.docx"
CANADIAN_TO_CSI_OUTPUT_SUFFIX = "_CSI.docx"

#: ``numbering_scheme`` values. The architect's proven multilevel list, the
#: committed built-in CSC PageFormat list, or literal typed CSI markers with
#: no automatic numbering at all.
SCHEME_ARCHITECT = "architect"
SCHEME_BUILTIN_CSC = "builtin_csc"
SCHEME_TYPED_CSI = "typed_csi"


@dataclass(frozen=True)
class ApplicationPolicy:
    """All mode-dependent mutation decisions for one formatting run.

    Keeping these switches together prevents a caller from selecting a mode for
    one phase while accidentally using another mode's numbering or validation
    behavior later in the pipeline.
    """

    conversion_mode: str
    preserve_target_numbering: bool
    convert_to_canadian: bool
    import_body_numbering: bool
    #: Suffix appended to the target stem for the published DOCX. Owned here so
    #: the engine's staged output and the pipeline's planned output paths can
    #: never disagree about a mode's naming.
    output_suffix: str
    #: Rewrite Canadian numbering back to literal typed CSI markers.
    convert_to_csi: bool = False
    #: Whether the run needs an architect template at all. False means the
    #: pipeline must not ask for one *and must reject one that is supplied*:
    #: silently ignoring a selected template would be the worst answer.
    requires_architect_template: bool = True
    #: Where the run's numbering comes from -- one of the ``SCHEME_*`` values.
    numbering_scheme: str = SCHEME_ARCHITECT
    #: Whether classified paragraphs are restyled with a role style. False
    #: for the architect-free modes: with no architect there is no formatting
    #: to apply, and swapping a paragraph's own style for a generated one that
    #: supplies no character formatting would silently flatten whatever its
    #: existing style gave it. Those modes apply numbering only.
    applies_role_styles: bool = True
    #: Output suffixes this application produces that are nonetheless valid
    #: *input* for this mode. Converting a spec to Canadian and later bringing
    #: it back is the whole point of the reverse mode, so its own earlier
    #: output must not be refused as "already formatted".
    accepted_output_suffixes: tuple[str, ...] = ()
    allow_ignored_paragraphs: bool = True
    #: Apply the architect's document-global shell (defaults, theme, settings,
    #: page layout, headers/footers). False leaves the target's own shell
    #: exactly as authored, which is the only honest option when there is no
    #: architect to take a shell from.
    apply_full_architect_shell: bool = True
    #: Whether this mode deliberately moves converted paragraphs' indentation.
    #: The two forward Canadian modes retarget every converted paragraph onto a
    #: different list's level indents, so their geometry is *meant* to change
    #: and the effective-geometry invariant would fire on every run. Every
    #: other mode must leave a paragraph rendering where it rendered before,
    #: and says so by leaving this False rather than by being exempted at the
    #: check.
    reindents_converted_paragraphs: bool = False
    contract_version: str = APPLICATION_POLICY_VERSION

    @property
    def is_format_only(self) -> bool:
        return self.conversion_mode == FORMAT_ONLY

    @property
    def uses_builtin_scheme(self) -> bool:
        return self.numbering_scheme == SCHEME_BUILTIN_CSC

    @property
    def imports_parts(self) -> bool:
        """Whether any architect or built-in part is imported into the target.

        ``canadian_to_csi`` writes literal markers and imports nothing at all;
        every other mode brings in styles, numbering, or both.
        """

        return self.applies_role_styles or self.import_body_numbering


def application_policy_for_mode(conversion_mode: str) -> ApplicationPolicy:
    """Return the single authoritative policy for ``conversion_mode``."""

    mode = validate_conversion_mode(conversion_mode)
    if mode == FORMAT_ONLY:
        return ApplicationPolicy(
            conversion_mode=mode,
            preserve_target_numbering=True,
            convert_to_canadian=False,
            import_body_numbering=False,
            output_suffix=FORMAT_ONLY_OUTPUT_SUFFIX,
        )
    if mode == CSI_TO_CANADIAN:
        return ApplicationPolicy(
            conversion_mode=mode,
            preserve_target_numbering=False,
            convert_to_canadian=True,
            import_body_numbering=True,
            output_suffix=CSI_TO_CANADIAN_OUTPUT_SUFFIX,
            reindents_converted_paragraphs=True,
        )
    if mode == CSI_TO_CANADIAN_STANDALONE:
        return ApplicationPolicy(
            conversion_mode=mode,
            preserve_target_numbering=False,
            convert_to_canadian=True,
            import_body_numbering=True,
            output_suffix=CSI_TO_CANADIAN_STANDALONE_OUTPUT_SUFFIX,
            reindents_converted_paragraphs=True,
            requires_architect_template=False,
            numbering_scheme=SCHEME_BUILTIN_CSC,
            applies_role_styles=False,
            apply_full_architect_shell=False,
        )
    if mode == CANADIAN_TO_CSI:
        return ApplicationPolicy(
            conversion_mode=mode,
            preserve_target_numbering=False,
            convert_to_canadian=False,
            import_body_numbering=False,
            output_suffix=CANADIAN_TO_CSI_OUTPUT_SUFFIX,
            convert_to_csi=True,
            accepted_output_suffixes=(
                CSI_TO_CANADIAN_STANDALONE_OUTPUT_SUFFIX.upper(),
                CSI_TO_CANADIAN_OUTPUT_SUFFIX.upper(),
            ),
            requires_architect_template=False,
            numbering_scheme=SCHEME_TYPED_CSI,
            applies_role_styles=False,
            apply_full_architect_shell=False,
        )
    raise AssertionError(f"Unhandled conversion mode: {mode}")


__all__ = [
    "APPLICATION_POLICY_VERSION",
    "CANADIAN_TO_CSI_OUTPUT_SUFFIX",
    "CSI_TO_CANADIAN_OUTPUT_SUFFIX",
    "CSI_TO_CANADIAN_STANDALONE_OUTPUT_SUFFIX",
    "FORMAT_ONLY_OUTPUT_SUFFIX",
    "SCHEME_ARCHITECT",
    "SCHEME_BUILTIN_CSC",
    "SCHEME_TYPED_CSI",
    "ApplicationPolicy",
    "application_policy_for_mode",
]
