"""Shared semantic inference for rendered Word numbering signatures.

The identifiers stored in ``numbering.xml`` are package-local, but the
rendered number format and level text are portable.  Both template analysis
and target classification use this helper so a CSI target can be recognized
even when the selected architect template uses Canadian numeric numbering.
"""

from __future__ import annotations

import re
from typing import Optional


def role_from_numbering_signature(
    num_fmt: object,
    lvl_text: object,
    ilvl: object,
) -> Optional[str]:
    """Return the structural role implied by a common CSI/CSC signature.

    Plain decimal lists at shallow levels are intentionally left ambiguous.
    They are common in ordinary prose as well as specifications and still need
    contextual classification.
    """

    fmt = str(num_fmt or "").strip().lower()
    marker = str(lvl_text or "")
    marker_upper = marker.upper()
    placeholders = re.findall(r"%\d+", marker)
    try:
        level = int(ilvl) if ilvl is not None else None
    except (TypeError, ValueError):
        level = None

    if "PART" in marker_upper:
        return "PART"
    # An article is a two-component shallow designation such as 1.01.  Deep
    # cumulative patterns (for example %1.%2.%3) are ordinary nested lists and
    # must not be deterministically promoted to ARTICLE.
    if (
        len(placeholders) == 2
        and level in {0, 1}
        and fmt in {"decimal", "decimalzero"}
        and re.fullmatch(r"\s*%\d+\s*\.\s*%\d+\s*", marker) is not None
    ):
        return "ARTICLE"
    if fmt in {"upperletter", "upperalpha"}:
        return "PARAGRAPH"
    if fmt in {"lowerletter", "loweralpha"}:
        return "SUBSUBPARAGRAPH"

    # Canadian PageFormat/NMS-style subordinate levels render as .1, .2,
    # etc.  When they share one multilevel list, ilvl carries the hierarchy.
    if fmt == "decimal" and re.match(r"^\s*\.\s*%\d+", marker):
        if level in {None, 0}:
            return "PARAGRAPH" if level == 0 else None
        if level == 1:
            return "SUBPARAGRAPH"
        # At deeper levels the absolute ilvl is ambiguous: it may be a third
        # level in a standalone list or the first dot-numbered level after
        # PART/ARTICLE in one larger list.  Leave it for contextual/exact-role
        # classification instead of assigning the wrong role.
        return None

    # Conventional CSI multilevel numbering commonly places the numeric
    # subparagraph at level 2 after article and upper-letter levels.
    if (
        fmt in {"decimal", "decimalzero"}
        and len(placeholders) <= 1
        and level is not None
        and level >= 2
    ):
        return "SUBPARAGRAPH"
    return None


__all__ = ["role_from_numbering_signature"]
