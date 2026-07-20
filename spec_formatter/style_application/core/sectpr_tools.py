from __future__ import annotations

import re
from typing import Dict, Generator, List, Optional, Tuple

from .xml_helpers import iter_element_xml_blocks


CANONICAL_SECTPR_ORDER = [
    "headerReference", "footerReference", "type", "pgSz", "pgMar", "paperSrc",
    "pgBorders", "lnNumType", "pgNumType", "cols", "formProt", "vAlign",
    "noEndnote", "titlePg", "textDirection", "bidi", "rtlGutter", "docGrid",
    "printerSettings", "sectPrChange",
]


_OUT_OF_SCOPE_SECTION_CONTAINERS = (
    "w:tbl",
    "w:drawing",
    "w:pict",
    "w:object",
    "w:txbxContent",
    "v:textbox",
    "wps:txbx",
)


def _out_of_scope_ranges(document_xml: str) -> List[Tuple[int, int]]:
    ranges: List[Tuple[int, int]] = []
    for name in _OUT_OF_SCOPE_SECTION_CONTAINERS:
        ranges.extend(
            (start, end)
            for start, end, _block in iter_element_xml_blocks(document_xml, name)
        )
    merged: List[Tuple[int, int]] = []
    for start, end in sorted(ranges):
        if merged and start <= merged[-1][1]:
            merged[-1] = (merged[-1][0], max(merged[-1][1], end))
        else:
            merged.append((start, end))
    return merged


def iter_document_sectpr_blocks(
    document_xml: str,
) -> Generator[Tuple[int, int, str], None, None]:
    """Yield section properties outside table/drawing/text-box subtrees."""
    excluded = _out_of_scope_ranges(document_xml)
    for start, end, block in iter_element_xml_blocks(document_xml, "w:sectPr"):
        if any(range_start <= start < range_end for range_start, range_end in excluded):
            continue
        yield start, end, block


def extract_all_sectpr_blocks(document_xml: str) -> List[str]:
    return [block for _start, _end, block in iter_document_sectpr_blocks(document_xml)]


def has_body_level_sectpr(document_xml: str) -> bool:
    paragraph_ranges = [
        (start, end)
        for start, end, _block in iter_element_xml_blocks(document_xml, "w:p")
    ]
    return any(
        not any(p_start < start and end < p_end for p_start, p_end in paragraph_ranges)
        for start, end, _block in iter_document_sectpr_blocks(document_xml)
    )


def extract_tag_block(xml: str, tag: str) -> Optional[str]:
    self_closing = re.search(rf'(<w:{tag}\b[^>]*/>)', xml)
    if self_closing:
        return self_closing.group(1)
    paired = re.search(rf'(<w:{tag}\b[^>]*>[\s\S]*?</w:{tag}>)', xml, flags=re.S)
    return paired.group(1) if paired else None


def strip_tag_block(xml: str, tag: str) -> str:
    xml = re.sub(rf'<w:{tag}\b[^>]*/>', '', xml)
    return re.sub(rf'<w:{tag}\b[^>]*>[\s\S]*?</w:{tag}>', '', xml, flags=re.S)


def child_tag_name(child_xml: str) -> Optional[str]:
    m = re.match(r'<w:([A-Za-z0-9]+)\b', child_xml)
    return m.group(1) if m else None


def extract_sectpr_children(inner: str) -> List[str]:
    """Return every direct child element as an exact raw XML block.

    Section properties can contain extension elements such as
    ``mc:AlternateContent`` and ``w14:*``.  Scanning only ``w:`` children would
    silently delete those nodes when callers rebuild the section.  This small
    lexical scanner accepts qualified and unqualified element names, observes
    quoted attribute values, and fails closed on malformed nesting.
    """
    children: List[str] = []
    stack: List[str] = []
    child_start: Optional[int] = None

    def iter_tags():
        cursor = 0
        while True:
            start = inner.find("<", cursor)
            if start < 0:
                return
            if inner.startswith("<!--", start):
                end = inner.find("-->", start + 4)
                if end < 0:
                    raise ValueError("Malformed sectPr child XML: unterminated comment")
                cursor = end + 3
                continue
            if inner.startswith("<![CDATA[", start):
                end = inner.find("]]>", start + 9)
                if end < 0:
                    raise ValueError("Malformed sectPr child XML: unterminated CDATA")
                cursor = end + 3
                continue
            if inner.startswith("<?", start):
                end = inner.find("?>", start + 2)
                if end < 0:
                    raise ValueError("Malformed sectPr child XML: unterminated processing instruction")
                cursor = end + 2
                continue
            if inner.startswith("<!", start):
                raise ValueError("Malformed sectPr child XML: unsupported declaration")

            quote: Optional[str] = None
            end = start + 1
            while end < len(inner):
                char = inner[end]
                if quote:
                    if char == quote:
                        quote = None
                elif char in {'"', "'"}:
                    quote = char
                elif char == ">":
                    break
                end += 1
            if end >= len(inner):
                raise ValueError("Malformed sectPr child XML: unterminated tag")

            raw = inner[start:end + 1]
            name_match = re.match(
                r"<\s*(?P<close>/)?\s*(?P<name>[A-Za-z_][\w.:-]*)\b",
                raw,
            )
            if not name_match:
                raise ValueError(f"Malformed sectPr child XML tag: {raw[:80]!r}")
            is_close = bool(name_match.group("close"))
            is_self_closing = not is_close and bool(re.search(r"/\s*>$", raw))
            yield start, end + 1, name_match.group("name"), is_close, is_self_closing
            cursor = end + 1

    for start, end, name, is_close, is_self_closing in iter_tags():

        if is_close:
            if not stack:
                raise ValueError(
                    f"Malformed sectPr child XML: unexpected closing tag </{name}>"
                )
            if stack[-1] != name:
                raise ValueError(
                    "Malformed sectPr child XML: mismatched closing tag "
                    f"</{name}> for <{stack[-1]}>"
                )
            stack.pop()
            if not stack and child_start is not None:
                children.append(inner[child_start:end])
                child_start = None
            continue

        if not stack:
            if is_self_closing:
                children.append(inner[start:end])
                continue
            child_start = start
        if not is_self_closing:
            stack.append(name)

    if stack:
        raise ValueError(
            f"Malformed sectPr child XML: unclosed element <{stack[-1]}>"
        )

    return children


def replace_nth_sectpr_block(document_xml: str, idx: int, replacement: str) -> str:
    matches = list(iter_document_sectpr_blocks(document_xml))
    if idx < 0 or idx >= len(matches):
        return document_xml
    start, end, _block = matches[idx]
    return document_xml[:start] + replacement + document_xml[end:]


def canonical_sectpr_order_index() -> Dict[str, int]:
    return {tag: i for i, tag in enumerate(CANONICAL_SECTPR_ORDER)}

