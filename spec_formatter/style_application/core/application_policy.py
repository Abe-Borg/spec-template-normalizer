"""Explicit, immutable behavior contracts for target application modes."""

from __future__ import annotations

from dataclasses import dataclass

from .csi_to_canadian import CSI_TO_CANADIAN, FORMAT_ONLY, validate_conversion_mode


APPLICATION_POLICY_VERSION = "2"


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
    allow_ignored_paragraphs: bool = True
    apply_full_architect_shell: bool = True
    contract_version: str = APPLICATION_POLICY_VERSION

    @property
    def is_format_only(self) -> bool:
        return self.conversion_mode == FORMAT_ONLY


def application_policy_for_mode(conversion_mode: str) -> ApplicationPolicy:
    """Return the single authoritative policy for ``conversion_mode``."""

    mode = validate_conversion_mode(conversion_mode)
    if mode == FORMAT_ONLY:
        return ApplicationPolicy(
            conversion_mode=mode,
            preserve_target_numbering=True,
            convert_to_canadian=False,
            import_body_numbering=False,
        )
    if mode == CSI_TO_CANADIAN:
        return ApplicationPolicy(
            conversion_mode=mode,
            preserve_target_numbering=False,
            convert_to_canadian=True,
            import_body_numbering=True,
        )
    raise AssertionError(f"Unhandled conversion mode: {mode}")


__all__ = [
    "APPLICATION_POLICY_VERSION",
    "ApplicationPolicy",
    "application_policy_for_mode",
]
