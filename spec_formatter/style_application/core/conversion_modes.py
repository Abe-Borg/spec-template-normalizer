"""The closed set of conversion modes and their validator.

These constants used to live in :mod:`csi_to_canadian`, which was fine while
that module was the only converter.  It is not any more: ``canadian_to_csi``
is its inverse, and a reverse converter importing its own name from the
forward one reads backwards.  The mode vocabulary is shared contract, so it
gets its own module and both converters import from here.

Two of the four modes need an architect template and two deliberately do not.
That distinction is *not* expressed here -- it belongs to
:class:`~.application_policy.ApplicationPolicy`, which owns every
mode-dependent decision.  This module only says which strings are real.
"""

from __future__ import annotations


FORMAT_ONLY = "format_only"
CSI_TO_CANADIAN = "csi_to_canadian"
#: CSI -> Canadian CSC PageFormat using the built-in scheme, with no architect.
CSI_TO_CANADIAN_STANDALONE = "csi_to_canadian_standalone"
#: Canadian CSC PageFormat -> typed CSI markers, with no architect.
CANADIAN_TO_CSI = "canadian_to_csi"

VALID_CONVERSION_MODES = frozenset(
    {
        FORMAT_ONLY,
        CSI_TO_CANADIAN,
        CSI_TO_CANADIAN_STANDALONE,
        CANADIAN_TO_CSI,
    }
)

#: Modes that run from the built-in scheme instead of an architect template.
ARCHITECT_FREE_MODES = frozenset({CSI_TO_CANADIAN_STANDALONE, CANADIAN_TO_CSI})


def validate_conversion_mode(value: object) -> str:
    """Return a validated mode string or raise before any document work."""

    if not isinstance(value, str) or value not in VALID_CONVERSION_MODES:
        choices = ", ".join(sorted(VALID_CONVERSION_MODES))
        raise ValueError(f"conversion_mode must be one of: {choices}")
    return value


__all__ = [
    "ARCHITECT_FREE_MODES",
    "CANADIAN_TO_CSI",
    "CSI_TO_CANADIAN",
    "CSI_TO_CANADIAN_STANDALONE",
    "FORMAT_ONLY",
    "VALID_CONVERSION_MODES",
    "validate_conversion_mode",
]
