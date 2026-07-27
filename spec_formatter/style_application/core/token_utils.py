from __future__ import annotations

import re
from pathlib import Path
from typing import Any, Dict

from .ooxml_text import read_xml_text
from .xml_helpers import iter_paragraph_xml_blocks, paragraph_text_from_block


PRESERVE_ACRONYMS = {
    "HVAC", "DDC", "VAV", "BAS", "BACNET", "AHU", "RTU", "VFD",
    "MAU", "ERV", "HRV", "FCU", "WSHP", "DX", "CHW", "HHW", "CW",
    "TAB", "MEP", "ASHRAE", "SMACNA", "NFPA", "UL", "ASTM",
    "PVC", "CPVC", "HDPE", "FRP", "BTU", "CFM", "GPM", "PSI",
}


_COMBINED_SECTION_HEADING_RE = re.compile(
    r"^\s*SECTION\s+"
    r"(?P<section>\d{2}(?:[ \t\u00a0]*\d{2}){2})"
    r"\s*[-\u2010\u2011\u2012\u2013\u2014\u2015:]\s*"
    r"(?P<title>\S(?:.*\S)?)\s*$",
    flags=re.IGNORECASE,
)


def _title_word(word: str) -> str:
    if "-" in word:
        return "-".join(_title_word(part) for part in word.split("-"))
    clean = re.sub(r"[^A-Za-z]", "", word or "")
    if clean and clean.upper() in PRESERVE_ACRONYMS:
        return word.upper() if word.isupper() else word
    return word.title() if word.isupper() else word


def smart_title_case(text: str) -> str:
    return " ".join(_title_word(word) for word in text.split())


def detect_case_pattern(text: str) -> str:
    """Detect whether text is UPPERCASE, Title Case, or mixed."""
    stripped = (text or "").strip()
    if not stripped:
        return "unknown"
    if stripped == stripped.upper():
        return "upper"
    words = stripped.split()
    if not words:
        return "unknown"
    capitalized = sum(1 for w in words if w[:1].isupper())
    if capitalized >= len(words) * 0.7:
        return "title"
    return "mixed"


def apply_case_pattern(text: str, pattern: str) -> str:
    """Apply a detected case pattern to text."""
    if pattern == "upper":
        return text.upper()
    if pattern == "title":
        return smart_title_case(text)
    return text


def extract_target_tokens(extract_dir: Path, classifications: Dict[str, Any]) -> Dict[str, str]:
    doc_path = extract_dir / "word" / "document.xml"
    if not doc_path.exists():
        return {}

    doc_xml = read_xml_text(doc_path)
    para_blocks = list(iter_paragraph_xml_blocks(doc_xml))
    tokens: Dict[str, str] = {}

    for item in classifications.get("classifications", []):
        if not isinstance(item, dict):
            continue
        idx = item.get("paragraph_index")
        role = item.get("csi_role")
        if not isinstance(idx, int) or idx < 0 or idx >= len(para_blocks):
            continue
        text = paragraph_text_from_block(para_blocks[idx][2]).strip()
        if role == "SectionID" and "SectionID" not in tokens:
            tokens["SectionID"] = text
            m = re.match(r"SECTION\s+([\d\s]+)", text, flags=re.IGNORECASE)
            if m:
                tokens["SectionID_numeric"] = re.sub(r"\s+", " ", m.group(1)).strip()
        if role == "SectionTitle" and "SectionTitle" not in tokens:
            tokens["SectionTitle"] = text
            tokens["SectionTitle_display"] = smart_title_case(text)
        if "SectionID" in tokens and "SectionTitle" in tokens:
            break

    # Some valid architect profiles intentionally expose neither SectionID nor
    # SectionTitle as body roles.  In that contract, a document's combined
    # top-level heading is retained as a deterministic ignored disposition.
    # Recover tokens only from that exact structural disposition and an
    # anchored ``SECTION 012900 - PAYMENT PROCEDURES`` shape.  This must not
    # become a general scan for section cross-references elsewhere in the body.
    if "SectionID" not in tokens or "SectionTitle" not in tokens:
        combined_headings: set[tuple[str, str]] = set()
        for item in classifications.get("ignored_paragraphs", []):
            if not isinstance(item, dict) or item.get("reason") != "section_header_no_role":
                continue
            idx = item.get("paragraph_index")
            if not isinstance(idx, int) or idx < 0 or idx >= len(para_blocks):
                continue
            text = paragraph_text_from_block(para_blocks[idx][2]).strip()
            match = _COMBINED_SECTION_HEADING_RE.fullmatch(text)
            if match is None:
                continue
            section_numeric = re.sub(r"\D", "", match.group("section"))
            title = match.group("title").strip()
            if len(section_numeric) != 6 or not title:
                continue
            combined_headings.add((section_numeric, title))

        # A target package represents one specification section.  Multiple
        # distinct combined headings are ambiguous, so fail closed rather than
        # choosing the first apparent match.
        if len(combined_headings) == 1:
            section_numeric, title = next(iter(combined_headings))
            compatible = True
            if "SectionID" in tokens:
                existing_id = (
                    tokens.get("SectionID_numeric")
                    or tokens.get("SectionID", "")
                )
                compatible = re.sub(r"\D", "", existing_id) == section_numeric
            if compatible and "SectionTitle" in tokens:
                compatible = (
                    re.sub(r"\s+", " ", tokens["SectionTitle"]).strip().casefold()
                    == re.sub(r"\s+", " ", title).strip().casefold()
                )
            if compatible:
                if "SectionID" not in tokens:
                    tokens["SectionID"] = f"SECTION {section_numeric}"
                    tokens["SectionID_numeric"] = section_numeric
                if "SectionTitle" not in tokens:
                    tokens["SectionTitle"] = title
                    tokens["SectionTitle_display"] = smart_title_case(title)

    return tokens
