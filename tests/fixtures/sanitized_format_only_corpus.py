"""Sanitized reproduction of the 154-paragraph Format-only failure shape.

The proprietary acceptance documents are deliberately not fixtures.  This
module builds a minimal OOXML architect/target pair whose target has the same
paragraph count and automatic-numbering index/level distribution as the
reported Payment Procedures case.  All visible strings are synthetic.
"""

from __future__ import annotations

from pathlib import Path

from spec_formatter.role_contract import ROLE_TO_ARCH_STYLE, ROLE_TO_STYLE_NAME
from tests.test_unified_roundtrip import (
    R_NS,
    W_NS,
    _rewrite_docx_parts,
    _section_xml,
    _write_docx,
)


NUMBERED_INDICES_BY_LEVEL: dict[int, tuple[int, ...]] = {
    0: (1, 151, 152),
    3: (2, 19, 23, 76),
    4: (3, 4, 22, 25, 36, 78, 80, 81, 85, 88, 93, 101, 105, 112, 132, 136),
    5: (
        6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18,
        26, 28, 30, 32, 34, 37, 46, 47, 58, 60, 63, 65, 67, 69,
        71, 73, 75, 83, 87, 89, 90, 91, 92, 94, 95, 96, 102, 106,
        107, 108, 109, 111, 113, 114, 115, 117, 118, 119, 120, 121,
        122, 123, 124, 125, 126, 127, 129, 130, 131, 133, 135, 138,
        139, 140, 141, 142, 143, 145, 146, 147, 148, 149, 150,
    ),
    6: (38, 39, 40, 41, 42, 43, 44, 48, 49, 50, 51, 52, 53, 54, 61, 97, 98, 99, 134),
    7: (55, 56, 57),
}

NUMBERED_LEVEL_BY_INDEX = {
    paragraph_index: level
    for level, paragraph_indices in NUMBERED_INDICES_BY_LEVEL.items()
    for paragraph_index in paragraph_indices
}

# These are the exact failure categories: three PARTs, four ARTICLEs, and the
# twelve automatically numbered sentence-form Section cross-references.
CRITICAL_MARKER_INDICES = frozenset(
    (*NUMBERED_INDICES_BY_LEVEL[0], *NUMBERED_INDICES_BY_LEVEL[3], *range(7, 19))
)

ROLE_BY_LEVEL = {
    0: "PART",
    3: "ARTICLE",
    4: "PARAGRAPH",
    5: "SUBPARAGRAPH",
    6: "SUBSUBPARAGRAPH",
    7: "SUBPARAGRAPH_LEVEL_5",
}

TARGET_STYLE_BY_LEVEL = {
    0: "PRT",
    3: "ART",
    4: "PR1",
    5: "PR2",
    6: "PR3",
    7: "PR4",
}

_NUMBERING_LEVELS = {
    0: ("decimal", "PART %1 -"),
    1: ("decimal", "%1.%2"),
    2: ("decimal", "%1.%2.%3"),
    3: ("decimal", "%1.%4"),
    4: ("upperLetter", "%5."),
    5: ("decimal", "%6."),
    6: ("lowerLetter", "%7."),
    7: ("decimal", "%8)"),
    8: ("lowerLetter", "%9)"),
}


def _numbering_xml(*, target: bool) -> str:
    abstract_id = "42" if target else "5"
    num_id = "17" if target else "5"
    levels = "".join(
        (
            f'<w:lvl w:ilvl="{level}"><w:start w:val="1"/>'
            f'<w:numFmt w:val="{number_format}"/>'
            f'<w:lvlText w:val="{level_text}"/>'
            f'<w:pPr><w:ind w:left="{720 + level * 180}" '
            'w:hanging="180"/></w:pPr></w:lvl>'
        )
        for level, (number_format, level_text) in _NUMBERING_LEVELS.items()
    )
    return (
        f'<w:numbering xmlns:w="{W_NS}">'
        f'<w:abstractNum w:abstractNumId="{abstract_id}">'
        f'<w:multiLevelType w:val="multilevel"/>{levels}</w:abstractNum>'
        f'<w:num w:numId="{num_id}">'
        f'<w:abstractNumId w:val="{abstract_id}"/></w:num>'
        '</w:numbering>'
    )


def _styles_xml(*, target: bool) -> str:
    styles = [
        '<w:style w:type="paragraph" w:default="1" w:styleId="Normal">'
        '<w:name w:val="Normal"/><w:qFormat/></w:style>'
    ]
    if target:
        styles.extend(
            [
                '<w:style w:type="paragraph" w:styleId="SID">'
                '<w:name w:val="Sanitized Section ID"/><w:basedOn w:val="Normal"/>'
                '</w:style>',
                '<w:style w:type="paragraph" w:styleId="CMT">'
                '<w:name w:val="Sanitized Non-CSI Note"/><w:basedOn w:val="Normal"/>'
                '<w:pPr><w:spacing w:before="37"/></w:pPr></w:style>',
                '<w:style w:type="paragraph" w:styleId="EOS">'
                '<w:name w:val="Sanitized End"/><w:basedOn w:val="Normal"/>'
                '</w:style>',
            ]
        )
        styles.extend(
            (
                f'<w:style w:type="paragraph" w:styleId="{style_id}">'
                f'<w:name w:val="Sanitized List Level {level}"/>'
                '<w:basedOn w:val="Normal"/><w:pPr><w:numPr>'
                f'<w:ilvl w:val="{level}"/><w:numId w:val="17"/>'
                '</w:numPr></w:pPr></w:style>'
            )
            for level, style_id in TARGET_STYLE_BY_LEVEL.items()
        )
    return (
        f'<w:styles xmlns:w="{W_NS}"><w:docDefaults>'
        '<w:rPrDefault><w:rPr/></w:rPrDefault>'
        '<w:pPrDefault><w:pPr/></w:pPrDefault></w:docDefaults>'
        f'{"".join(styles)}</w:styles>'
    )


def _sanitized_text(paragraph_index: int) -> str:
    critical_text = {
        1: "GENERAL",
        2: "SUMMARY",
        19: "REFERENCES",
        23: "SUBMITTALS",
        76: "QUALITY CONTROL",
        151: "PRODUCTS",
        152: "EXECUTION",
        153: "END OF SECTION",
    }
    if paragraph_index in critical_text:
        return critical_text[paragraph_index]
    if 7 <= paragraph_index <= 18:
        return f"Section 000000 Sanitized Reference {paragraph_index:02d} applies."
    level = NUMBERED_LEVEL_BY_INDEX.get(paragraph_index)
    if level is not None:
        return f"Sanitized numbered requirement {paragraph_index:03d}."
    if paragraph_index == 0:
        return "SECTION 000000 - SANITIZED PAYMENT PROCEDURES"
    return f"Sanitized non-CSI note {paragraph_index:03d}."


def _target_document_xml() -> str:
    paragraphs: list[str] = []
    for paragraph_index in range(154):
        level = NUMBERED_LEVEL_BY_INDEX.get(paragraph_index)
        if paragraph_index == 0:
            style_id = "SID"
        elif paragraph_index == 153:
            style_id = "EOS"
        else:
            style_id = TARGET_STYLE_BY_LEVEL[level] if level is not None else "CMT"

        # Direct formatting on ignored paragraphs makes byte-stability visible.
        # Level-4 paragraphs reproduce the production regression: their direct
        # "not bold" override must be removed because the architect PARAGRAPH
        # style supplies bold, while the unrelated color remains direct.
        if style_id == "CMT":
            direct_run_properties = (
                '<w:rPr><w:b/><w:color w:val="445566"/></w:rPr>'
            )
        elif level == 4:
            direct_run_properties = (
                '<w:rPr><w:b w:val="0"/><w:color w:val="556677"/></w:rPr>'
            )
        else:
            direct_run_properties = ""
        paragraphs.append(
            '<w:p><w:pPr>'
            f'<w:pStyle w:val="{style_id}"/>'
            '</w:pPr><w:r>'
            f'{direct_run_properties}<w:t>{_sanitized_text(paragraph_index)}</w:t>'
            '</w:r></w:p>'
        )
    return (
        '<?xml version="1.0" encoding="UTF-8"?>'
        f'<w:document xmlns:w="{W_NS}" xmlns:r="{R_NS}"><w:body>'
        f'{"".join(paragraphs)}{_section_xml(architect=False, first=False)}'
        '</w:body></w:document>'
    )


def _architect_document_xml() -> str:
    role_paragraphs = [
        '<w:p><w:r><w:t>SECTION 000000</w:t></w:r></w:p>',
        '<w:p><w:r><w:t>SANITIZED ARCHITECT SHELL</w:t></w:r></w:p>',
    ]
    for level, role in ROLE_BY_LEVEL.items():
        run_properties = '<w:rPr><w:b/></w:rPr>' if role == "PARAGRAPH" else ""
        role_paragraphs.append(
            '<w:p><w:pPr><w:numPr>'
            f'<w:ilvl w:val="{level}"/><w:numId w:val="5"/>'
            '</w:numPr></w:pPr><w:r>'
            f'{run_properties}<w:t>Sanitized architect exemplar for {role}</w:t>'
            '</w:r></w:p>'
        )
    return (
        '<?xml version="1.0" encoding="UTF-8"?>'
        f'<w:document xmlns:w="{W_NS}" xmlns:r="{R_NS}"><w:body>'
        f'{"".join(role_paragraphs)}{_section_xml(architect=True, first=True)}'
        '</w:body></w:document>'
    )


def write_sanitized_format_only_pair(
    architect_path: Path,
    target_path: Path,
) -> None:
    """Create the synthetic architect and 154-paragraph target packages."""

    _write_docx(architect_path, architect=True)
    _write_docx(target_path, architect=False)
    _rewrite_docx_parts(
        architect_path,
        {
            "word/document.xml": _architect_document_xml(),
            "word/styles.xml": _styles_xml(target=False),
            "word/numbering.xml": _numbering_xml(target=False),
        },
    )
    _rewrite_docx_parts(
        target_path,
        {
            "word/document.xml": _target_document_xml(),
            "word/styles.xml": _styles_xml(target=True),
            "word/numbering.xml": _numbering_xml(target=True),
        },
    )


def sanitized_template_classifier(**kwargs):
    """Classify the synthetic architect without any model or network call."""

    classifiable = [
        item
        for item in kwargs["slim_bundle"].get("paragraphs", [])
        if item.get("skip_reason") is None
    ]
    expected_roles = ("SectionID", "SectionTitle", *ROLE_BY_LEVEL.values())
    if len(classifiable) != len(expected_roles):
        raise AssertionError(
            f"expected {len(expected_roles)} sanitized architect exemplars, "
            f"got {len(classifiable)}"
        )

    role_rows = list(zip(classifiable, expected_roles))
    return {
        "create_styles": [
            {
                "styleId": ROLE_TO_ARCH_STYLE[role],
                "name": ROLE_TO_STYLE_NAME[role],
                "type": "paragraph",
                "derive_from_paragraph_index": paragraph["paragraph_index"],
                "basedOn": "Normal",
                "role": role,
            }
            for paragraph, role in role_rows
        ],
        "apply_pStyle": [
            {
                "paragraph_index": paragraph["paragraph_index"],
                "styleId": ROLE_TO_ARCH_STYLE[role],
            }
            for paragraph, role in role_rows
        ],
        "ignored_paragraphs": [],
        "roles": {
            role: {
                "styleId": ROLE_TO_ARCH_STYLE[role],
                "exemplar_paragraph_index": paragraph["paragraph_index"],
            }
            for paragraph, role in role_rows
        },
        "notes": ["generated sanitized Format-only regression corpus"],
    }


assert len(NUMBERED_LEVEL_BY_INDEX) == 121
assert len(CRITICAL_MARKER_INDICES) == 19
