import codecs
from pathlib import Path

import pytest

from arch_env_extractor import _read_xml_part
from docx_decomposer import build_slim_bundle
from ooxml_text import decode_xml_bytes, prepare_xml_text_for_utf8


W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"


@pytest.mark.parametrize(
    ("payload", "expected"),
    [
        (
            codecs.BOM_UTF8
            + '<?xml version="1.0" encoding="UTF-8"?><root>\u2713</root>'.encode("utf-8"),
            "\u2713",
        ),
        (
            codecs.BOM_UTF16_LE
            + '<?xml version="1.0" encoding="UTF-16"?><root>\u2713</root>'.encode("utf-16-le"),
            "\u2713",
        ),
        (
            codecs.BOM_UTF16_BE
            + '<?xml version="1.0" encoding="UTF-16"?><root>\u2713</root>'.encode("utf-16-be"),
            "\u2713",
        ),
        (
            '<?xml version="1.0" encoding="iso-8859-1"?><root>caf\u00e9</root>'.encode(
                "iso-8859-1"
            ),
            "caf\u00e9",
        ),
    ],
    ids=["utf8-bom", "utf16-le-bom", "utf16-be-bom", "declared-latin1"],
)
def test_decode_xml_bytes_encoding_matrix(payload: bytes, expected: str):
    assert expected in decode_xml_bytes(payload)
    assert 'encoding="UTF-8"' in prepare_xml_text_for_utf8(
        '<?xml version="1.0" encoding="UTF-16"?><root/>'
    )


def test_phase1_analysis_accepts_utf16_wordprocessingml_parts(tmp_path: Path):
    word = tmp_path / "word"
    word.mkdir()
    document = (
        '<?xml version="1.0" encoding="UTF-16"?>'
        f'<w:document xmlns:w="{W_NS}"><w:body>'
        '<w:p><w:r><w:t>PART 1 - GENERAL</w:t></w:r></w:p>'
        '</w:body></w:document>'
    )
    styles = (
        '<?xml version="1.0" encoding="UTF-16"?>'
        f'<w:styles xmlns:w="{W_NS}">'
        '<w:style w:type="paragraph" w:default="1" w:styleId="Normal">'
        '<w:name w:val="Normal"/></w:style></w:styles>'
    )
    (word / "document.xml").write_bytes(document.encode("utf-16"))
    (word / "styles.xml").write_bytes(styles.encode("utf-16"))

    bundle = build_slim_bundle(tmp_path)

    assert bundle["paragraphs"][0]["text"] == "PART 1 - GENERAL"
    assert "Normal" in bundle["style_catalog"]
    assert "PART 1 - GENERAL" in _read_xml_part(tmp_path, "word/document.xml")
