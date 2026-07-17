from __future__ import annotations

import re
from typing import Any, Dict, List, Optional, Set, Tuple

# CSI headings occur in substantially more forms than the canonical examples in
# the prompt. In particular, Word templates commonly use non-breaking spaces,
# compact six-digit section numbers, and a combined "SECTION ... - TITLE" line.
_WS = r"[\s\u00a0]"
RE_SECTION_ID = re.compile(
    rf"^SECTION{_WS}*\d{{2}}(?:{_WS}*[-\u2010-\u2015]?{_WS}*\d{{2}}){{2}}\b",
    re.IGNORECASE,
)
RE_SECTION_WITH_TITLE = re.compile(
    rf"^(SECTION{_WS}*\d{{2}}(?:{_WS}*[-\u2010-\u2015]?{_WS}*\d{{2}}){{2}})"
    rf"{_WS}*(?:[-\u2010-\u2015:]{_WS}*)?(.+)$",
    re.IGNORECASE,
)
RE_PART = re.compile(
    rf"^PART{_WS}+(?:\d+|[IVXLCDM]+)(?:{_WS}*[-\u2010-\u2015:.]{_WS}*|{_WS}+)(?=\S)",
    re.IGNORECASE,
)
RE_ARTICLE = re.compile(r"^\d{1,2}\.\d{2,3}(?:\s+|$)")
RE_ALPHA_PARA = re.compile(r"^[A-Z](?:\.|\))\s+")
RE_NUMERIC_SUB = re.compile(r"^\d+(?:\.|\))\s+")
RE_LOWER_SUBSUB = re.compile(r"^[a-z](?:\.|\))\s+")
RE_END_OF_SECTION = re.compile(
    rf"^END{_WS}+OF{_WS}+(?:SECTION|DIVISION)"
    rf"(?:{_WS}+\d{{2}}(?:{_WS}*[-\u2010-\u2015]?{_WS}*\d{{2}}){{2}})?"
    rf"{_WS}*[.\-\u2010-\u2015]*{_WS}*$",
    re.IGNORECASE,
)

RE_COPYRIGHT_NOTICE = re.compile(r"^Copyright\s+\d{4}\s+by\s+", re.IGNORECASE)
RE_DISTRIBUTION_NOTICE = re.compile(r"^Exclusively published and distributed by\s+", re.IGNORECASE)

RE_SPECIFIER_RETAIN = re.compile(r"^Retain\s+(first|both|one|option|Sections?|definitions?|paragraph|sub)\b", re.IGNORECASE)
RE_SPECIFIER_REVISE = re.compile(r"^Revise\s+(this|reference|first|sub)\b", re.IGNORECASE)
RE_SPECIFIER_VERIFY_SECTION = re.compile(r"^Verify that Section\b", re.IGNORECASE)
RE_SPECIFIER_EDITING_INSTRUCTION = re.compile(r"^See Editing Instruction\b", re.IGNORECASE)
RE_SPECIFIER_SPECIFY_PARTS = re.compile(r"^Specify parts in\s+(first|sub)\b", re.IGNORECASE)
RE_SPECIFIER_PARAGRAPH_DEFINED = re.compile(r"^Paragraph below is defined in Section\b", re.IGNORECASE)
RE_SPECIFIER_OPTION = re.compile(r"^Option:\s*", re.IGNORECASE)
RE_SPECIFIER_HIGH_COMPRESSIVE = re.compile(r"^High-compressive-strength inserts may permit\b", re.IGNORECASE)
RE_SPECIFIER_VERIFY_SUITABILITY = re.compile(r"^Verify suitability of\b", re.IGNORECASE)

def is_editor_note(raw_text: str) -> bool:
    txt = (raw_text or "").strip()
    if not txt:
        return False
    if txt.startswith("[") and txt.endswith("]"):
        return True
    return bool(
        re.match(
            r"^(?:SPEC(?:IFICATION)?\s+)?(?:EDITOR|SPECIFIER)(?:\s+(?:NOTE|INSTRUCTION))?\s*:",
            txt,
            re.IGNORECASE,
        )
        or re.match(
            r"^SPEC(?:IFICATION)?\s+(?:NOTE|INSTRUCTION)\s*:",
            txt,
            re.IGNORECASE,
        )
        or re.match(
            r"^HIDDEN\s+(?:EDITOR|SPECIFIER)\s+(?:NOTE|INSTRUCTION)\b",
            txt,
            re.IGNORECASE,
        )
    )


def is_copyright_notice(raw_text: str) -> bool:
    txt = (raw_text or "").strip()
    if not txt:
        return False
    return bool(RE_COPYRIGHT_NOTICE.match(txt) or RE_DISTRIBUTION_NOTICE.match(txt))


def is_specifier_note(raw_text: str) -> bool:
    txt = (raw_text or "").strip()
    if not txt:
        return False

    if (
        RE_SPECIFIER_RETAIN.match(txt)
        or RE_SPECIFIER_REVISE.match(txt)
        or RE_SPECIFIER_VERIFY_SECTION.match(txt)
        or RE_SPECIFIER_EDITING_INSTRUCTION.match(txt)
        or RE_SPECIFIER_SPECIFY_PARTS.match(txt)
        or RE_SPECIFIER_PARAGRAPH_DEFINED.match(txt)
        or RE_SPECIFIER_VERIFY_SUITABILITY.match(txt)
    ):
        return True

    low = txt.lower()
    has_para_ref = any(token in low for token in ["paragraph below", "paragraphs below", "subparagraph below", "subparagraphs below"])
    has_edit_verb = any(token in low for token in ["retain", "revise", "delete", "insert"])
    if has_para_ref and has_edit_verb:
        return True

    if "requires calculating and detailing at each use" in low:
        return True

    if RE_SPECIFIER_OPTION.match(txt) and "may be used" in low:
        return True

    if "catalogs indicate" in low:
        return True

    if RE_SPECIFIER_HIGH_COMPRESSIVE.match(txt):
        return True

    return False


def compute_skip_reason(raw_text: str, contains_sectpr: bool, in_table: bool) -> Optional[str]:
    text = (raw_text or "").strip()
    if not text:
        return "empty"
    if contains_sectpr:
        return "sectPr"
    if in_table:
        return "in_table"
    if is_editor_note(text):
        return "editor_note"
    if is_copyright_notice(text):
        return "copyright_notice"
    if is_specifier_note(text):
        return "specifier_note"
    return None


def is_classifiable_paragraph(paragraph: Dict[str, Any]) -> bool:
    """Return whether a paragraph needs an explicit styled/ignored disposition.

    Editorial/copyright content is auditable input, not invisible input. Only
    structural exclusions are outside the classification universe.
    """
    if "skip_reason" in paragraph:
        return paragraph.get("skip_reason") not in {"empty", "sectPr", "in_table"}
    text = (paragraph.get("text") or "").strip()
    skip_reason = compute_skip_reason(
        text,
        bool(paragraph.get("contains_sectPr", False)),
        bool(paragraph.get("in_table", False)),
    )
    return skip_reason not in {"empty", "sectPr", "in_table"}


def is_role_candidate_paragraph(paragraph: Dict[str, Any]) -> bool:
    """Return whether a paragraph may serve as CSI content or a role exemplar."""
    if "skip_reason" in paragraph:
        return paragraph.get("skip_reason") is None
    text = (paragraph.get("text") or "").strip()
    return compute_skip_reason(
        text,
        bool(paragraph.get("contains_sectPr", False)),
        bool(paragraph.get("in_table", False)),
    ) is None


def detect_role_signal(text: str, *, numeric_is_strong: bool, lower_is_strong: bool) -> Optional[str]:
    txt = (text or "").strip()
    if not txt:
        return None
    if RE_SECTION_ID.match(txt):
        return "SectionID"
    if RE_END_OF_SECTION.match(txt):
        return "END_OF_SECTION"
    if RE_PART.match(txt):
        return "PART"
    if RE_ARTICLE.match(txt):
        return "ARTICLE"
    if RE_ALPHA_PARA.match(txt):
        return "PARAGRAPH"
    if numeric_is_strong and RE_NUMERIC_SUB.match(txt):
        return "SUBPARAGRAPH"
    if lower_is_strong and RE_LOWER_SUBSUB.match(txt):
        return "SUBSUBPARAGRAPH"
    return None


def detect_numbering_role(
    paragraph: Dict[str, Any],
    numbering_catalog: Optional[Dict[str, Any]],
) -> Optional[str]:
    """Infer a CSI role from the effective Word numbering pattern.

    This is the critical path for templates where Word renders every marker and
    the stored paragraph text is only ``GENERAL``/``SUMMARY``/``Scope``.
    Ambiguous plain decimal lists deliberately return ``None``.
    """
    if not isinstance(numbering_catalog, dict):
        return None
    numpr = paragraph.get("effective_numPr") or paragraph.get("numPr")
    if not isinstance(numpr, dict):
        return None
    num_id = numpr.get("numId")
    ilvl = numpr.get("ilvl")
    if num_id is None:
        return None

    nums = numbering_catalog.get("nums", {})
    abstracts = numbering_catalog.get("abstracts", {})
    num = nums.get(str(num_id), {}) if isinstance(nums, dict) else {}
    abstract_id = num.get("abstractNumId") if isinstance(num, dict) else None
    abstract = abstracts.get(str(abstract_id), {}) if isinstance(abstracts, dict) else {}
    levels = abstract.get("levels", []) if isinstance(abstract, dict) else []
    level = next((item for item in levels if str(item.get("ilvl")) == str(ilvl)), None)
    if not isinstance(level, dict):
        return None

    num_fmt = level.get("numFmt")
    lvl_text = str(level.get("lvlText") or "")
    for override in num.get("levelOverrides", []) if isinstance(num, dict) else []:
        if str(override.get("ilvl")) == str(ilvl):
            num_fmt = override.get("numFmt", num_fmt)
            lvl_text = str(override.get("lvlText", lvl_text) or "")
            break

    fmt = str(num_fmt or "").lower()
    marker = lvl_text.upper()
    placeholder_count = len(re.findall(r"%\d+", lvl_text))
    try:
        level_number = int(ilvl) if ilvl is not None else None
    except (TypeError, ValueError):
        level_number = None

    if "PART" in marker:
        return "PART"
    if placeholder_count >= 2 and fmt in {"decimal", "decimalzero"} and level_number in {0, 1}:
        return "ARTICLE"
    if fmt in {"upperletter", "upperalpha"}:
        return "PARAGRAPH"
    if fmt in {"lowerletter", "loweralpha"}:
        return "SUBSUBPARAGRAPH"
    if fmt in {"decimal", "decimalzero"} and level_number is not None and level_number >= 2:
        return "SUBPARAGRAPH"
    return None


def infer_expected_roles(
    paragraphs: List[Dict[str, Any]],
    *,
    numbering_catalog: Optional[Dict[str, Any]] = None,
) -> Tuple[Set[str], Dict[str, List[int]]]:
    classifiable = [p for p in paragraphs if is_role_candidate_paragraph(p)]

    expected: Set[str] = set()
    strong_hits: Dict[str, List[int]] = {
        "SectionID": [],
        "PART": [],
        "ARTICLE": [],
        "PARAGRAPH": [],
        "SUBPARAGRAPH": [],
        "SUBSUBPARAGRAPH": [],
        "END_OF_SECTION": [],
    }

    section_indices: List[int] = []
    # Numeric/lowercase markers are only strong within an established CSI
    # hierarchy. A single "A." elsewhere must not turn every "1." or "a."
    # in a title page or option list into a structural role.
    alpha_context = False
    numeric_context = False
    for p in classifiable:
        idx = int(p["paragraph_index"])
        text = (p.get("text") or "").strip()
        signal = detect_numbering_role(p, numbering_catalog) or detect_role_signal(
            text,
            numeric_is_strong=alpha_context,
            lower_is_strong=numeric_context,
        )
        if signal:
            expected.add(signal)
            strong_hits[signal].append(idx)
            if signal == "SectionID":
                section_indices.append(idx)
                combined = RE_SECTION_WITH_TITLE.match(text)
                if combined and combined.group(2).strip():
                    expected.add("SectionTitle")
                    strong_hits.setdefault("SectionTitle", []).append(idx)

            if signal in {"SectionID", "SectionTitle", "PART", "ARTICLE", "END_OF_SECTION"}:
                alpha_context = False
                numeric_context = False
            elif signal == "PARAGRAPH":
                alpha_context = True
                numeric_context = False
            elif signal == "SUBPARAGRAPH":
                numeric_context = True

    for idx in section_indices:
        nxt = next((p for p in classifiable if int(p["paragraph_index"]) == idx + 1), None)
        if not nxt:
            continue
        if detect_role_signal(
            (nxt.get("text") or "").strip(),
            numeric_is_strong=False,
            lower_is_strong=False,
        ) is None:
            expected.add("SectionTitle")
            strong_hits.setdefault("SectionTitle", []).append(int(nxt["paragraph_index"]))
            break

    return expected, {k: v for k, v in strong_hits.items() if v}
