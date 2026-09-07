"""One grammar for CSI MasterFormat section numbers.

The classifier, the target token extractor, and the header/footer token
patcher used to carry three incompatible section-number regexes plus a fourth
canonicaliser. ``SECTION 23 0500`` was invisible to the classifier but visible
to the token extractor and patcher, so the two halves of the header/footer
substitution could disagree about whether a target even had a section number.
Every consumer now shares this module.

Grammar
-------

A section number is three two-digit pairs with optional spaces, tabs, or
non-breaking spaces between them, optionally followed by a MasterFormat
level-4 suffix of a period and two digits::

    230500      23 05 00      23 0500      23 05 00.13      230500.13

The canonical form is the digits with every separator removed and the level-4
suffix rendered with its period: ``230500`` or ``230500.13``. Consumers
compare canonical forms and render replacements in the shape of the text they
replace, so ``SECTION 23 05 00`` patched to section ``260513`` becomes
``SECTION 26 05 13``.
"""

from __future__ import annotations

import re
from typing import Optional

_PAIR_SEPARATOR = r"[ \t ]*"
_SPACE_RUN_RE = re.compile(r"[ \t ]+")

#: Unanchored, group-free pattern text for embedding in larger expressions.
SECTION_NUMBER_PATTERN = (
    rf"\d{{2}}(?:{_PAIR_SEPARATOR}\d{{2}}){{2}}"  # three two-digit pairs
    r"(?:\.\d{2})?"  # optional MasterFormat level-4 suffix
)

#: Boundary that stops a section number from being read out of a longer digit
#: run, a letter-adjacent identifier, or a partial level-4 suffix such as
#: ``230500.1``. Underscores are allowed neighbours so a footer filename such
#: as ``233100_Metal Ducts.docx`` still exposes its number.
SECTION_NUMBER_BOUNDARY = r"(?![A-Za-z0-9])(?!\.\d)"
_SECTION_NUMBER_LEFT_BOUNDARY = r"(?<![A-Za-z0-9.])"

#: A section number anywhere in text, exposed as the ``number`` group.
SECTION_NUMBER_RE = re.compile(
    rf"{_SECTION_NUMBER_LEFT_BOUNDARY}(?P<number>{SECTION_NUMBER_PATTERN})"
    rf"{SECTION_NUMBER_BOUNDARY}"
)

#: ``SECTION <number>`` at the start of a paragraph, exposed as ``number``.
SECTION_HEADING_RE = re.compile(
    rf"^\s*SECTION\s+(?P<number>{SECTION_NUMBER_PATTERN}){SECTION_NUMBER_BOUNDARY}",
    re.IGNORECASE,
)

#: A bare section number at the start of text (``23 05 00 - PIPING``).
_LEADING_SECTION_NUMBER_RE = re.compile(
    rf"^\s*(?P<number>{SECTION_NUMBER_PATTERN}){SECTION_NUMBER_BOUNDARY}"
)

#: ``SECTION <number>`` anywhere in text (headers, footers, cross-references).
LABELED_SECTION_RE = re.compile(
    rf"\bSECTION\s+(?P<number>{SECTION_NUMBER_PATTERN}){SECTION_NUMBER_BOUNDARY}",
    re.IGNORECASE,
)


def canonical_section_number(value: Optional[str]) -> str:
    """Return the canonical ``230500`` / ``230500.13`` form, or ``""``.

    The whole value must be one section number (surrounding whitespace is
    ignored); anything else, including a bare six-digit run embedded in a
    longer number, yields the empty string so callers fail closed.
    """

    text = (value or "").strip()
    if not text or re.fullmatch(SECTION_NUMBER_PATTERN, text) is None:
        return ""
    digits = re.sub(r"\D", "", text)
    if len(digits) == 6:
        return digits
    if len(digits) == 8:
        return f"{digits[:6]}.{digits[6:]}"
    return ""


def section_number_display_form(value: Optional[str]) -> str:
    """Return the section number as written, with whitespace runs collapsed.

    Accepts a bare number, a number followed by other text, or a
    ``SECTION <number>`` label, and returns the number in its source shape
    (``23 05 00``, ``230500.13``). Returns ``""`` when the value does not
    begin with a recognisable section number.
    """

    text = (value or "").strip()
    if not text:
        return ""
    match = SECTION_HEADING_RE.match(text) or _LEADING_SECTION_NUMBER_RE.match(text)
    if match is None:
        return ""
    return _SPACE_RUN_RE.sub(" ", match.group("number"))


def render_section_number_like(source_form: str, target_canonical: str) -> str:
    """Render ``target_canonical`` in the digit grouping of ``source_form``.

    ``("23 05 00", "260513")`` gives ``"26 05 13"``; when the digit counts do
    not line up (a level-3 source and a level-4 target, or vice versa) the
    canonical target is returned unchanged.
    """

    target_digits = re.sub(r"\D", "", target_canonical or "")
    digit_groups = list(re.finditer(r"\d+", source_form or ""))
    if (
        not target_digits
        or not digit_groups
        or sum(len(group.group(0)) for group in digit_groups) != len(target_digits)
    ):
        return target_canonical
    pieces = []
    source_cursor = 0
    target_cursor = 0
    for group in digit_groups:
        pieces.append(source_form[source_cursor:group.start()])
        length = len(group.group(0))
        pieces.append(target_digits[target_cursor:target_cursor + length])
        source_cursor = group.end()
        target_cursor += length
    pieces.append(source_form[source_cursor:])
    return "".join(pieces)
