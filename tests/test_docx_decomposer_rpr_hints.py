from __future__ import annotations

from docx_decomposer import _read_on_off_tag, paragraph_rpr_hints_from_block


def test_read_on_off_tag_boolean_cases() -> None:
    assert _read_on_off_tag("<w:rPr><w:b/></w:rPr>", "b") is True
    assert _read_on_off_tag('<w:rPr><w:b w:val="true"/></w:rPr>', "b") is True
    assert _read_on_off_tag('<w:rPr><w:b w:val="1"/></w:rPr>', "b") is True
    assert _read_on_off_tag('<w:rPr><w:b w:val="false"/></w:rPr>', "b") is False
    assert _read_on_off_tag('<w:rPr><w:b w:val="0"/></w:rPr>', "b") is False
    assert _read_on_off_tag('<w:rPr><w:i w:val="false"/></w:rPr>', "i") is False
    assert _read_on_off_tag('<w:rPr><w:caps w:val="false"/></w:rPr>', "caps") is False
    assert _read_on_off_tag('<w:rPr><w:u w:val="single"/></w:rPr>', "u") is True
    assert _read_on_off_tag('<w:rPr><w:u w:val="none"/></w:rPr>', "u") is False
    assert _read_on_off_tag("<w:rPr><w:sz w:val=\"18\"/></w:rPr>", "b") is None


def test_paragraph_rpr_hints_respects_explicit_false_values() -> None:
    p_xml = (
        "<w:p>"
        "<w:r>"
        "<w:rPr>"
        '<w:b w:val="false"/>'
        '<w:i w:val="false"/>'
        '<w:caps w:val="false"/>'
        '<w:sz w:val="18"/>'
        '<w:rFonts w:ascii="Garamond"/>'
        "</w:rPr>"
        "<w:t>Example</w:t>"
        "</w:r>"
        "</w:p>"
    )

    assert paragraph_rpr_hints_from_block(p_xml) == {
        "bold": False,
        "italic": False,
        "caps": False,
        "sz": "18",
        "font": "Garamond",
    }


def test_paragraph_rpr_hints_omits_absent_boolean_keys() -> None:
    p_xml = "<w:p><w:r><w:rPr><w:sz w:val=\"22\"/></w:rPr><w:t>Example</w:t></w:r></w:p>"

    assert paragraph_rpr_hints_from_block(p_xml) == {"sz": "22"}
