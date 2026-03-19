from __future__ import annotations

import re
from typing import Any, Dict, List, Optional, Set, Tuple

RE_SECTION_ID = re.compile(r"^SECTION\s+\d{2}\s+\d{2}\s+\d{2}\b", re.IGNORECASE)
RE_PART = re.compile(r"^PART\s+\d+\s*[-–—]\s+", re.IGNORECASE)
RE_ARTICLE = re.compile(r"^\d+\.\d{2,}\s+")
RE_ALPHA_PARA = re.compile(r"^[A-Z]\.\s+")
RE_NUMERIC_SUB = re.compile(r"^\d+\.\s+")
RE_LOWER_SUBSUB = re.compile(r"^[a-z]\.\s+")

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


def is_editor_note(raw_text: str) -> bool:
    txt = (raw_text or "").strip()
    return bool(txt) and txt.startswith("[") and txt.endswith("]")


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
    if text.upper() == "END OF SECTION":
        return "end_of_section"
    if is_editor_note(text):
        return "editor_note"
    if is_copyright_notice(text):
        return "copyright_notice"
    if is_specifier_note(text):
        return "specifier_note"
    return None


def is_classifiable_paragraph(paragraph: Dict[str, Any]) -> bool:
    if "skip_reason" in paragraph:
        return paragraph.get("skip_reason") is None
    text = (paragraph.get("text") or "").strip()
    skip_reason = compute_skip_reason(
        text,
        bool(paragraph.get("contains_sectPr", False)),
        bool(paragraph.get("in_table", False)),
    )
    return skip_reason is None


def detect_role_signal(text: str, *, numeric_is_strong: bool, lower_is_strong: bool) -> Optional[str]:
    txt = (text or "").strip()
    if not txt:
        return None
    if RE_SECTION_ID.match(txt):
        return "SectionID"
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


def infer_expected_roles(paragraphs: List[Dict[str, Any]]) -> Tuple[Set[str], Dict[str, List[int]]]:
    classifiable = [p for p in paragraphs if is_classifiable_paragraph(p)]
    has_alpha = any(RE_ALPHA_PARA.match((p.get("text") or "").strip()) for p in classifiable)
    has_numeric = any(RE_NUMERIC_SUB.match((p.get("text") or "").strip()) for p in classifiable)
    numeric_is_strong = has_alpha
    lower_is_strong = has_numeric

    expected: Set[str] = set()
    strong_hits: Dict[str, List[int]] = {
        "SectionID": [],
        "PART": [],
        "ARTICLE": [],
        "PARAGRAPH": [],
        "SUBPARAGRAPH": [],
        "SUBSUBPARAGRAPH": [],
    }

    section_indices: List[int] = []
    for p in classifiable:
        idx = int(p["paragraph_index"])
        text = (p.get("text") or "").strip()
        signal = detect_role_signal(text, numeric_is_strong=numeric_is_strong, lower_is_strong=lower_is_strong)
        if signal:
            expected.add(signal)
            strong_hits[signal].append(idx)
            if signal == "SectionID":
                section_indices.append(idx)

    for idx in section_indices:
        nxt = next((p for p in classifiable if int(p["paragraph_index"]) == idx + 1), None)
        if not nxt:
            continue
        if detect_role_signal((nxt.get("text") or "").strip(), numeric_is_strong=numeric_is_strong, lower_is_strong=lower_is_strong) is None:
            expected.add("SectionTitle")
            break

    return expected, {k: v for k, v in strong_hits.items() if v}
