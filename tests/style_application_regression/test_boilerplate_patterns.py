from pathlib import Path

import pytest

from spec_formatter.style_application.core.classification import build_phase2_slim_bundle, strip_boilerplate_with_report


def _write_document_xml(tmp_path: Path, paragraph_texts: list[str]) -> Path:
    word = tmp_path / "word"
    word.mkdir(parents=True)
    body = "".join(f"<w:p><w:r><w:t>{text}</w:t></w:r></w:p>" for text in paragraph_texts)
    doc_xml = (
        '<?xml version="1.0" encoding="UTF-8"?>\n'
        '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
        f"<w:body>{body}</w:body></w:document>"
    )
    (word / "document.xml").write_text(doc_xml, encoding="utf-8")
    return tmp_path


BOILERPLATE_INPUTS = [
    'Retain definitions remaining after this Section has been edited.',
    'Paragraph below is defined in Section 013300 "Submittal Procedures" as a "Delegated-Design Submittal."',
    "Trapeze pipe hanger in paragraph below requires calculating and detailing at each use.",
    "Metal framing system in first paragraph below requires calculating and detailing at each use.",
    "Equipment support in first paragraph below requires calculating and detailing at each use.",
    "Verify suitability of fasteners in this article for use in lightweight concrete...",
    "Verify suitability of fasteners in two subparagraphs below...",
    "Specify parts in first three subparagraphs below as galvanized or painted, as required.",
    "Option:  Thermal-hanger shield inserts may be used.  Include steel weight-distribution plate...",
    "Manufacturers' catalogs indicate that copper pipe hangers are small, typically NPS 4...",
    "High-compressive-strength inserts may permit use of shorter shields or shields with less arc span.",
]


REAL_CONTENT_INPUTS = [
    "Retain existing pipe supports where indicated.",
    "Verify all dimensions in the field before fabrication.",
    "Option: Provide seismic bracing where required by code.",
    "Section Includes:",
    "Adjustable Steel Clevis Hangers: (MSS Type 1.) B-Line B 3100",
    "Install hangers and supports to allow controlled thermal and seismic movement of piping systems.",
    "Structural Steel Welding Qualifications: Qualify procedures and personnel according to AWS D1.1",
    "Metal Pipe-Hanger Installation: Comply with MSS SP-69 and MSS SP-89.",
    "Hanger Adjustments: Adjust hangers to distribute loads equally on attachments.",
]


@pytest.mark.parametrize("text", BOILERPLATE_INPUTS)
def test_new_masterspec_patterns_strip_entire_paragraph(text: str):
    cleaned, tags = strip_boilerplate_with_report(text)
    assert cleaned == ""
    assert "masterspec_instruction" in tags


@pytest.mark.parametrize("text", REAL_CONTENT_INPUTS)
def test_new_masterspec_patterns_do_not_strip_real_content(text: str):
    cleaned, tags = strip_boilerplate_with_report(text)
    assert cleaned == text
    assert tags == []


def test_build_phase2_bundle_reports_expected_removed_paragraphs(tmp_path: Path):
    input_texts = [
        "SECTION 22 05 29",
        "HANGERS AND SUPPORTS FOR PLUMBING PIPING AND EQUIPMENT",
        *BOILERPLATE_INPUTS,
        *REAL_CONTENT_INPUTS,
        "PART 1 GENERAL",
        "1.01 SUMMARY",
    ]
    extract_dir = _write_document_xml(tmp_path, input_texts)

    bundle = build_phase2_slim_bundle(
        extract_dir,
        available_roles=[
            "SectionID",
            "SectionTitle",
            "PART",
            "ARTICLE",
            "PARAGRAPH",
            "SUBPARAGRAPH",
            "SUBSUBPARAGRAPH",
        ],
    )

    removed = bundle["filter_report"]["paragraphs_removed_entirely"]
    removed_indices = {entry["paragraph_index"] for entry in removed}
    expected_removed_indices = set(range(2, 2 + len(BOILERPLATE_INPUTS)))
    assert expected_removed_indices.issubset(removed_indices)

    for idx in [0, 1, 2 + len(BOILERPLATE_INPUTS), len(input_texts) - 1]:
        assert idx not in removed_indices

