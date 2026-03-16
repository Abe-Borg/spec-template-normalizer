from __future__ import annotations

import re
from typing import Any, Dict, List, Optional, Set, Tuple

RE_SECTION_ID = re.compile(r"^SECTION\s+\d{2}\s+\d{2}\s+\d{2}\b", re.IGNORECASE)
RE_PART = re.compile(r"^PART\s+\d+\s*[-–—]\s+", re.IGNORECASE)
RE_ARTICLE = re.compile(r"^\d+\.\d{2,}\s+")
RE_ALPHA_PARA = re.compile(r"^[A-Z]\.\s+")
RE_NUMERIC_SUB = re.compile(r"^\d+\.\s+")
RE_LOWER_SUBSUB = re.compile(r"^[a-z]\.\s+")


def is_editor_note(raw_text: str) -> bool:
    txt = (raw_text or "").strip()
    return bool(txt) and txt.startswith("[") and txt.endswith("]")


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
