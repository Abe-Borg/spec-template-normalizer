from pathlib import Path

from arch_env_extractor import _read_xml_part
from docx_decomposer import build_slim_bundle
from ooxml_text import decode_xml_bytes, prepare_xml_text_for_utf8


W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"


def test_decode_xml_bytes_honors_utf16_bom_and_declared_legacy_encoding():
    utf16 = '<?xml version="1.0" encoding="UTF-16"?><root>✓</root>'.encode("utf-16")
    assert "✓" in decode_xml_bytes(utf16)

    latin1 = b'<?xml version="1.0" encoding="iso-8859-1"?><root>caf\xe9</root>'
    assert "café" in decode_xml_bytes(latin1)
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
