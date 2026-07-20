import zipfile
from pathlib import Path

import pytest

from spec_formatter.style_application.arch_env_applier import apply_page_layout
from spec_formatter.style_application.phase2_invariants import verify_phase2_invariants


DOC_XML_SINGLE = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <w:body>
    <w:p><w:r><w:t>Hi</w:t></w:r></w:p>
    <w:sectPr>
      <w:headerReference w:type="default" r:id="rId7"/>
      <w:footerReference w:type="default" r:id="rId8"/>
      <w:pgMar w:top="1080" w:right="1080" w:bottom="1080" w:left="1080" w:header="540" w:footer="540"/>
      <w:pgNumType w:start="1"/>
    </w:sectPr>
  </w:body>
</w:document>'''


ARCH_SECTPR = (
    '<w:sectPr>'
    '<w:pgSz w:w="12240" w:h="15840"/>'
    '<w:pgMar w:top="1800" w:right="1080" w:bottom="1440" w:left="2160" w:header="900" w:footer="720"/>'
    '<w:cols w:space="720"/>'
    '<w:docGrid w:linePitch="360"/>'
    '</w:sectPr>'
)


RELS_XML = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="document.xml"/>
  <Relationship Id="rId7" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header1.xml"/>
  <Relationship Id="rId8" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer" Target="footer1.xml"/>
</Relationships>'''


def _write_minimal_docx(path: Path, document_xml: str, rels_xml: str = RELS_XML):
    with zipfile.ZipFile(path, "w") as z:
        z.writestr("word/document.xml", document_xml)
        z.writestr("word/_rels/document.xml.rels", rels_xml)
        z.writestr("word/header1.xml", "<w:hdr xmlns:w='http://schemas.openxmlformats.org/wordprocessingml/2006/main'><w:p/></w:hdr>")
        z.writestr("word/footer1.xml", "<w:ftr xmlns:w='http://schemas.openxmlformats.org/wordprocessingml/2006/main'><w:p/></w:ftr>")


def test_apply_page_layout_syncs_managed_tags_and_preserves_hf_refs(tmp_path):
    extract = tmp_path / "extract"
    (extract / "word").mkdir(parents=True)
    (extract / "word" / "document.xml").write_text(DOC_XML_SINGLE, encoding="utf-8")

    registry = {
        "page_layout": {
            "default_section": {"sectPr": ARCH_SECTPR},
            "section_chain": [],
        }
    }

    log = []
    apply_page_layout(extract, registry, log)

    out = (extract / "word" / "document.xml").read_text(encoding="utf-8")
    assert 'w:top="1800"' in out
    assert 'w:left="2160"' in out
    assert '<w:pgSz w:w="12240" w:h="15840"/>' in out
    assert '<w:headerReference w:type="default" r:id="rId7"/>' in out
    assert '<w:footerReference w:type="default" r:id="rId8"/>' in out
    assert '<w:pgNumType w:start="1"/>' in out


def test_apply_page_layout_expands_self_closing_sectpr(tmp_path):
    extract = tmp_path / "extract"
    (extract / "word").mkdir(parents=True)
    document = (
        '<w:document xmlns:w="http://schemas.openxmlformats.org/'
        'wordprocessingml/2006/main"><w:body><w:p/>'
        '<w:sectPr w:rsidR="A1"/></w:body></w:document>'
    )
    (extract / "word" / "document.xml").write_text(document, encoding="utf-8")
    registry = {
        "page_layout": {
            "default_section": {"sectPr": ARCH_SECTPR},
            "section_chain": [],
        }
    }

    apply_page_layout(extract, registry, [])

    out = (extract / "word" / "document.xml").read_text(encoding="utf-8")
    assert '<w:sectPr w:rsidR="A1">' in out
    assert "</w:sectPr>" in out
    assert '<w:pgSz w:w="12240" w:h="15840"/>' in out
    assert 'w:top="1800"' in out


def test_apply_page_layout_creates_missing_body_level_sectpr(tmp_path):
    extract = tmp_path / "extract"
    (extract / "word").mkdir(parents=True)
    document = (
        '<w:document xmlns:w="http://schemas.openxmlformats.org/'
        'wordprocessingml/2006/main"><w:body>'
        '<w:p><w:r><w:t>Body</w:t></w:r></w:p>'
        '</w:body></w:document>'
    )
    (extract / "word" / "document.xml").write_text(document, encoding="utf-8")
    registry = {
        "page_layout": {
            "default_section": {"sectPr": ARCH_SECTPR},
            "section_chain": [],
        }
    }
    log = []

    apply_page_layout(extract, registry, log)

    out = (extract / "word" / "document.xml").read_text(encoding="utf-8")
    assert out.index("<w:sectPr>") < out.index("</w:body>")
    assert '<w:pgMar w:top="1800"' in out
    assert any("Created missing final body-level sectPr" in item for item in log)


def test_apply_page_layout_adds_final_section_after_paragraph_section_break(tmp_path):
    extract = tmp_path / "extract"
    (extract / "word").mkdir(parents=True)
    document = (
        '<w:document xmlns:w="http://schemas.openxmlformats.org/'
        'wordprocessingml/2006/main"><w:body>'
        '<w:p><w:pPr><w:sectPr/></w:pPr><w:r><w:t>Break</w:t></w:r></w:p>'
        '</w:body></w:document>'
    )
    (extract / "word" / "document.xml").write_text(document, encoding="utf-8")
    registry = {
        "page_layout": {
            "default_section": {"sectPr": ARCH_SECTPR},
            "section_chain": [],
        }
    }

    apply_page_layout(extract, registry, [])

    out = (extract / "word" / "document.xml").read_text(encoding="utf-8")
    assert out.count("<w:sectPr") == 2
    assert out.count('w:top="1800"') == 2


def test_apply_page_layout_preserves_nested_textbox_section_properties(tmp_path):
    extract = tmp_path / "extract"
    (extract / "word").mkdir(parents=True)
    textbox = (
        '<w:p><w:r><w:drawing><w:txbxContent><w:p><w:pPr>'
        '<w:sectPr><w:pgSz w:w="999"/><w:pgBorders w:offsetFrom="text"/></w:sectPr>'
        '</w:pPr><w:r><w:t>Nested</w:t></w:r></w:p>'
        '</w:txbxContent></w:drawing></w:r></w:p>'
    )
    document = (
        '<w:document xmlns:w="http://schemas.openxmlformats.org/'
        f'wordprocessingml/2006/main"><w:body>{textbox}'
        '<w:sectPr><w:pgSz w:w="12000"/></w:sectPr></w:body></w:document>'
    )
    (extract / "word" / "document.xml").write_text(document, encoding="utf-8")
    registry = {
        "page_layout": {
            "default_section": {"sectPr": ARCH_SECTPR},
            "section_chain": [],
        }
    }

    apply_page_layout(extract, registry, [])

    out = (extract / "word" / "document.xml").read_text(encoding="utf-8")
    assert textbox in out
    assert out.count('<w:pgSz w:w="12240" w:h="15840"/>') == 1


def test_apply_page_layout_normalizes_utf16_target_to_truthful_utf8(tmp_path):
    extract = tmp_path / "extract"
    (extract / "word").mkdir(parents=True)
    document = (
        '<?xml version="1.0" encoding="UTF-16"?>'
        '<w:document xmlns:w="http://schemas.openxmlformats.org/'
        'wordprocessingml/2006/main"><w:body><w:sectPr/></w:body></w:document>'
    )
    path = extract / "word" / "document.xml"
    path.write_bytes(document.encode("utf-16"))
    registry = {
        "page_layout": {
            "default_section": {"sectPr": ARCH_SECTPR},
            "section_chain": [],
        }
    }

    apply_page_layout(extract, registry, [])

    payload = path.read_bytes()
    assert b'encoding="UTF-8"' in payload
    assert 'w:top="1800"' in payload.decode("utf-8")


def test_phase2_invariants_allow_layout_change_but_reject_hf_rel_change(tmp_path):
    src = tmp_path / "src.docx"
    out = tmp_path / "out.docx"

    _write_minimal_docx(src, DOC_XML_SINGLE)

    changed_doc = DOC_XML_SINGLE.replace('w:top="1080"', 'w:top="1800"')
    changed_rels = RELS_XML.replace('Target="header1.xml"', 'Target="header2.xml"')
    _write_minimal_docx(out, changed_doc, changed_rels)

    with pytest.raises(RuntimeError, match="relationship subset changed"):
        verify_phase2_invariants(
            src_docx=src,
            new_document_xml=changed_doc.encode("utf-8"),
            new_docx=out,
            arch_template_registry={"page_layout": {"default_section": {"sectPr": ARCH_SECTPR}}},
        )
