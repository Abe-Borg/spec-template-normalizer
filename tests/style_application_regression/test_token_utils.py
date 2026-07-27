from spec_formatter.style_application.core.token_utils import extract_target_tokens


def _write_document(tmp_path, *paragraphs):
    word_dir = tmp_path / "word"
    word_dir.mkdir(parents=True)
    body = "".join(
        f"<w:p><w:r><w:t>{text}</w:t></w:r></w:p>"
        for text in paragraphs
    )
    (word_dir / "document.xml").write_text(
        '<w:document xmlns:w="http://schemas.openxmlformats.org/'
        'wordprocessingml/2006/main"><w:body>'
        f"{body}</w:body></w:document>",
        encoding="utf-8",
    )


def test_combined_ignored_section_heading_supplies_target_tokens(tmp_path):
    _write_document(
        tmp_path,
        "SECTION 012900 - PAYMENT PROCEDURES",
        "See SECTION 055000 - MISCELLANEOUS METALS for related work.",
    )

    tokens = extract_target_tokens(
        tmp_path,
        {
            "classifications": [],
            "ignored_paragraphs": [
                {"paragraph_index": 0, "reason": "section_header_no_role"},
                {"paragraph_index": 1, "reason": "non_csi_content"},
            ],
        },
    )

    assert tokens == {
        "SectionID": "SECTION 012900",
        "SectionID_numeric": "012900",
        "SectionTitle": "PAYMENT PROCEDURES",
        "SectionTitle_display": "Payment Procedures",
    }


def test_combined_heading_fallback_fails_closed_on_ambiguity(tmp_path):
    _write_document(
        tmp_path,
        "SECTION 012900 - PAYMENT PROCEDURES",
        "SECTION 055000 - MISCELLANEOUS METALS",
    )

    tokens = extract_target_tokens(
        tmp_path,
        {
            "classifications": [],
            "ignored_paragraphs": [
                {"paragraph_index": 0, "reason": "section_header_no_role"},
                {"paragraph_index": 1, "reason": "section_header_no_role"},
            ],
        },
    )

    assert tokens == {}


def test_combined_heading_fallback_does_not_mix_conflicting_explicit_section_id(
    tmp_path,
):
    _write_document(
        tmp_path,
        "SECTION 012900",
        "SECTION 055000 - MISCELLANEOUS METALS",
    )

    tokens = extract_target_tokens(
        tmp_path,
        {
            "classifications": [
                {"paragraph_index": 0, "csi_role": "SectionID"},
            ],
            "ignored_paragraphs": [
                {"paragraph_index": 1, "reason": "section_header_no_role"},
            ],
        },
    )

    assert tokens == {
        "SectionID": "SECTION 012900",
        "SectionID_numeric": "012900",
    }


def test_combined_heading_fallback_does_not_mix_conflicting_explicit_title(
    tmp_path,
):
    _write_document(
        tmp_path,
        "PAYMENT PROCEDURES",
        "SECTION 055000 - MISCELLANEOUS METALS",
    )

    tokens = extract_target_tokens(
        tmp_path,
        {
            "classifications": [
                {"paragraph_index": 0, "csi_role": "SectionTitle"},
            ],
            "ignored_paragraphs": [
                {"paragraph_index": 1, "reason": "section_header_no_role"},
            ],
        },
    )

    assert tokens == {
        "SectionTitle": "PAYMENT PROCEDURES",
        "SectionTitle_display": "Payment Procedures",
    }
