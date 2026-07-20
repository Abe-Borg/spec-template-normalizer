"""Shared semantic inference for rendered Word numbering signatures.

The identifiers stored in ``numbering.xml`` are package-local, but the
rendered number format and level text are portable.  Both template analysis
and target classification use this helper so a CSI target can be recognized
even when the selected architect template uses Canadian numeric numbering.
"""

from __future__ import annotations

import re
from typing import Any, Dict, Optional


PAGEFORMAT_LEVEL_ROLES = {
    0: "PART",
    1: "ARTICLE",
    2: "PARAGRAPH",
    3: "SUBPARAGRAPH",
    4: "SUBSUBPARAGRAPH",
    5: "SUBPARAGRAPH_LEVEL_5",
    6: "SUBPARAGRAPH_LEVEL_6",
    7: "SUBPARAGRAPH_LEVEL_7",
    8: "SUBPARAGRAPH_LEVEL_8",
}


def _level_from_catalog(
    numbering_catalog: Dict[str, Any],
    num_id: object,
    ilvl: object,
) -> Optional[Dict[str, str]]:
    """Return one effective numbering level from either catalog wire shape.

    Phase 1 stores levels/overrides as arrays while the Phase 2 reader stores
    them as dictionaries.  Keeping that compatibility here prevents the two
    engines from making different semantic decisions about the same list.
    """

    level_key = str(ilvl)
    nums = numbering_catalog.get("nums", {})
    abstracts = numbering_catalog.get("abstracts", {})
    if not isinstance(nums, dict) or not isinstance(abstracts, dict):
        return None
    num = nums.get(str(num_id))
    if not isinstance(num, dict):
        return None
    abstract = abstracts.get(str(num.get("abstractNumId")))
    if not isinstance(abstract, dict):
        return None

    raw_levels = abstract.get("levels", {})
    raw_level: object = None
    if isinstance(raw_levels, dict):
        raw_level = raw_levels.get(level_key)
    elif isinstance(raw_levels, list):
        raw_level = next(
            (
                item
                for item in raw_levels
                if isinstance(item, dict) and str(item.get("ilvl")) == level_key
            ),
            None,
        )
    if not isinstance(raw_level, dict):
        return None

    result = {
        key: str(value)
        for key, value in raw_level.items()
        if key in {"ilvl", "numFmt", "lvlText", "start", "lvlRestart"}
        and value is not None
    }
    result.setdefault("ilvl", level_key)

    raw_overrides = num.get("levelOverrides", num.get("overrides", {}))
    override: object = None
    if isinstance(raw_overrides, dict):
        override = raw_overrides.get(level_key)
    elif isinstance(raw_overrides, list):
        override = next(
            (
                item
                for item in raw_overrides
                if isinstance(item, dict) and str(item.get("ilvl")) == level_key
            ),
            None,
        )
    if isinstance(override, dict):
        for key in ("numFmt", "lvlText", "startOverride", "lvlRestart"):
            value = override.get(key)
            if value is not None:
                result[key] = str(value)
    return result


def _is_decimal_pattern(
    level: Optional[Dict[str, str]],
    pattern: str,
    *,
    flags: int = 0,
) -> bool:
    return bool(
        isinstance(level, dict)
        and str(level.get("numFmt") or "").lower() in {"decimal", "decimalzero"}
        and re.fullmatch(pattern, str(level.get("lvlText") or ""), flags=flags)
    )


def role_from_numbering_catalog(
    numbering_catalog: object,
    num_id: object,
    ilvl: object,
) -> Optional[str]:
    """Infer a role with sibling-level context from one Word list instance.

    Repeated PageFormat markers such as ``.1`` cannot be classified from one
    level alone.  This resolver first proves the surrounding PART/article
    ladder and only then assigns an absolute PageFormat hierarchy role.
    """

    if not isinstance(numbering_catalog, dict) or num_id is None:
        return None
    try:
        level_number = int(ilvl)
    except (TypeError, ValueError):
        return None
    if level_number not in PAGEFORMAT_LEVEL_ROLES:
        return None

    current = _level_from_catalog(numbering_catalog, num_id, level_number)
    if current is None:
        return None

    # Conventional CSI signatures with distinctive punctuation do not need
    # list context, but resolving them here gives callers one common entrypoint.
    direct = role_from_numbering_signature(
        current.get("numFmt"), current.get("lvlText"), level_number
    )
    if direct is not None:
        return direct

    level_zero = _level_from_catalog(numbering_catalog, num_id, 0)
    level_one = _level_from_catalog(numbering_catalog, num_id, 1)
    level_two = _level_from_catalog(numbering_catalog, num_id, 2)
    part_anchor = _is_decimal_pattern(
        level_zero,
        r"\s*(?:PART\s+)?%1\s*\.?\s*",
        flags=re.IGNORECASE,
    )
    article_anchor = _is_decimal_pattern(
        level_one,
        r"\s*%1\s*\.\s*%2\s*",
    )
    first_paragraph_anchor = _is_decimal_pattern(
        level_two,
        r"\s*\.\s*%3\s*",
    )
    if not (part_anchor and article_anchor and first_paragraph_anchor):
        return None

    # Every numeric level through the requested one must continue the same
    # PageFormat ladder.  This prevents an unrelated deep list from being
    # assigned a structural role merely because it happens to use ilvl=5.
    if level_number >= 2:
        for candidate_level in range(2, level_number + 1):
            candidate = _level_from_catalog(
                numbering_catalog, num_id, candidate_level
            )
            if not _is_decimal_pattern(
                candidate,
                rf"\s*\.\s*%{candidate_level + 1}\s*",
            ):
                return None
    return PAGEFORMAT_LEVEL_ROLES[level_number]


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
    # The deeper CSI levels are distinguished by punctuation even when Word
    # stores only the marker placeholder in lvlText.
    own_marker = rf"%{level + 1}" if level is not None else r"%\d+"
    if re.fullmatch(rf"\s*{own_marker}\s*\)\s*", marker):
        # ``1)`` at ilvl 2 is also a long-standing numeric subparagraph
        # pattern.  It becomes the fifth semantic level only after at least
        # three subordinate list levels have been established.
        if fmt in {"decimal", "decimalzero"} and level is not None and level >= 3:
            return "SUBPARAGRAPH_LEVEL_5"
        if fmt in {"lowerletter", "loweralpha"} and level is not None and level >= 4:
            return "SUBPARAGRAPH_LEVEL_6"
    if re.fullmatch(rf"\s*\(\s*{own_marker}\s*\)\s*", marker):
        if fmt in {"decimal", "decimalzero"} and level is not None and level >= 5:
            return "SUBPARAGRAPH_LEVEL_7"
        if fmt in {"lowerletter", "loweralpha"} and level is not None and level >= 6:
            return "SUBPARAGRAPH_LEVEL_8"

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


__all__ = [
    "PAGEFORMAT_LEVEL_ROLES",
    "role_from_numbering_catalog",
    "role_from_numbering_signature",
]
