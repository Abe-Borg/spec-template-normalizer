import base64
import hashlib
import re
import xml.etree.ElementTree as ET
from pathlib import Path

import pytest

from spec_formatter.style_application.core.opc_paths import resolve_internal_relationship_target
from spec_formatter.style_application.core.xml_helpers import iter_paragraph_xml_blocks
from spec_formatter.style_application.header_footer_importer import (
    import_headers_footers,
    patch_footer_tokens,
    patch_header_footer_tokens,
    remap_header_footer_numids,
)


def _seed_extract(tmp_path: Path) -> Path:
    (tmp_path / "word" / "_rels").mkdir(parents=True, exist_ok=True)
    (tmp_path / "word" / "document.xml").write_text(
        '<?xml version="1.0" encoding="UTF-8"?>'
        '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
        'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
        '<w:body><w:p><w:r><w:t>x</w:t></w:r></w:p><w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr></w:body></w:document>',
        encoding="utf-8",
    )
    (tmp_path / "word" / "_rels" / "document.xml.rels").write_text(
        '<?xml version="1.0" encoding="UTF-8"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>'
        '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header9.xml"/>'
        '</Relationships>',
        encoding="utf-8",
    )
    (tmp_path / "[Content_Types].xml").write_text(
        '<?xml version="1.0" encoding="UTF-8"?>'
        '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
        '<Default Extension="xml" ContentType="application/xml"/>'
        '</Types>',
        encoding="utf-8",
    )
    (tmp_path / "word" / "header9.xml").write_text("<old/>", encoding="utf-8")
    return tmp_path


def test_import_headers_footers_replaces_parts_and_refs(tmp_path):
    extract = _seed_extract(tmp_path)
    ct_path = extract / "[Content_Types].xml"
    ct_path.write_text(
        ct_path.read_text(encoding="utf-8").replace(
            "</Types>",
            '<Override PartName="/word/header9.xml" '
            'ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/>'
            "</Types>",
        ),
        encoding="utf-8",
    )
    registry = {
        "headers_footers": {
            "headers": [
                {
                    "part_name": "word/header1.xml",
                    "rid": "rId10",
                    "xml": '<w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"/>',
                    "media": [
                        {
                            "path": "media/logo.png",
                            "content_base64": base64.b64encode(b"png").decode("ascii"),
                        }
                    ],
                }
            ],
            "footers": [
                {
                    "part_name": "word/footer1.xml",
                    "rid": "rId11",
                    "xml": '<w:ftr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"/>',
                }
            ],
        },
        "page_layout": {
            "section_chain": [
                {
                    "header_refs": {"default": "rId10"},
                    "footer_refs": {"default": "rId11"},
                }
            ],
            "default_section": {
                "header_refs": {"default": "rId10"},
                "footer_refs": {"default": "rId11"},
            },
        },
    }

    log = []
    import_headers_footers(extract, registry, log)

    assert not (extract / "word" / "header9.xml").exists()
    assert (extract / "word" / "header1.xml").exists()
    assert (extract / "word" / "footer1.xml").exists()
    assert any(p.read_bytes() == b"png" for p in (extract / "word" / "media").iterdir())

    rels_xml = (extract / "word" / "_rels" / "document.xml.rels").read_text(encoding="utf-8")
    assert "relationships/header" in rels_xml
    assert "relationships/footer" in rels_xml
    assert "header9.xml" not in rels_xml

    doc_xml = (extract / "word" / "document.xml").read_text(encoding="utf-8")
    assert "headerReference" in doc_xml
    assert "footerReference" in doc_xml
    assert "<ns0:" not in doc_xml
    assert "<w:p" in doc_xml
    assert len(list(iter_paragraph_xml_blocks(doc_xml))) == 1

    ct_xml = (extract / "[Content_Types].xml").read_text(encoding="utf-8")
    assert "/word/header1.xml" in ct_xml
    assert "/word/footer1.xml" in ct_xml
    assert "/word/header9.xml" not in ct_xml
    assert 'Extension="png"' in ct_xml

def test_hf_rewire_preserves_unknown_document_prefixes(tmp_path):
    extract = _seed_extract(tmp_path)
    doc = (
        '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
        'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" '
        'xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" '
        'xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" mc:Ignorable="w14">'
        '<w:body><w:p><w:r><w14:paraId w14:val="1234"/><w:t>x</w:t></w:r></w:p>'
        '<w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr></w:body></w:document>'
    )
    (extract / "word" / "document.xml").write_text(doc, encoding="utf-8")
    registry = {
        "headers_footers": {"headers": [{"part_name": "word/header1.xml", "rid": "rId10", "xml": '<w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"/>'}]},
        "page_layout": {"default_section": {"header_refs": {"default": "rId10"}}},
    }
    log = []
    import_headers_footers(extract, registry, log)
    out = (extract / "word" / "document.xml").read_text(encoding="utf-8")
    assert "xmlns:mc" in out and "xmlns:w14" in out
    assert 'mc:Ignorable="w14"' in out
    assert "<w14:paraId" in out


def test_hf_media_import_does_not_overwrite_existing_body_media(tmp_path):
    extract = _seed_extract(tmp_path)
    media_dir = extract / "word" / "media"
    media_dir.mkdir(parents=True, exist_ok=True)
    (media_dir / "image1.png").write_bytes(b"body")

    registry = {
        "headers_footers": {
            "headers": [{
                "part_name": "word/header1.xml",
                "rid": "rId10",
                "xml": '<w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"/>',
                "rels_xml": '<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/image1.png"/></Relationships>',
                "media": [{"path": "media/image1.png", "content_base64": base64.b64encode(b"header").decode("ascii")}],
            }]
        },
        "page_layout": {"default_section": {"header_refs": {"default": "rId10"}}},
    }

    result = import_headers_footers(extract, registry, [])
    assert (media_dir / "image1.png").read_bytes() == b"body"
    assert result.media_names
    assert all(name.startswith("word/media/hf_") for name in result.media_names)


def test_hf_media_allocation_is_case_insensitive(tmp_path):
    extract = _seed_extract(tmp_path)
    media_dir = extract / "word" / "media"
    media_dir.mkdir(parents=True, exist_ok=True)
    payload = b"header"
    digest = hashlib.sha1(payload).hexdigest()[:8]
    colliding_name = f"HF_HEADER1_01_{digest}.PNG"
    (media_dir / colliding_name).write_bytes(b"existing")
    registry = {
        "headers_footers": {
            "headers": [{
                "part_name": "word/header1.xml",
                "rid": "rId10",
                "xml": '<w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"/>',
                "media": [{
                    "path": "media/logo.png",
                    "content_base64": base64.b64encode(payload).decode("ascii"),
                }],
            }],
        },
        "page_layout": {"default_section": {"header_refs": {"default": "rId10"}}},
    }

    result = import_headers_footers(extract, registry, [])

    assert (media_dir / colliding_name).read_bytes() == b"existing"
    assert any(name.endswith(f"_{digest}_1.png") for name in result.media_names)


def test_conflicting_image_mime_uses_override_without_duplicate_default(tmp_path):
    extract = _seed_extract(tmp_path)
    ct_path = extract / "[Content_Types].xml"
    ct_path.write_text(
        ct_path.read_text(encoding="utf-8").replace(
            "</Types>",
            '<Default Extension="png" ContentType="image/png"/></Types>',
        ),
        encoding="utf-8",
    )
    registry = {
        "headers_footers": {
            "headers": [{
                "part_name": "word/header1.xml",
                "rid": "rId10",
                "xml": (
                    '<w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
                    'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" '
                    'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">'
                    '<w:p><w:r><w:drawing><a:blip r:embed="rIdImage"/>'
                    '</w:drawing></w:r></w:p></w:hdr>'
                ),
                "rels_xml": (
                    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
                    '<Relationship Id="rIdImage" '
                    'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" '
                    'Target="media/logo.png"/></Relationships>'
                ),
                "media": [{
                    "rel_id": "rIdImage",
                    "target": "media/logo.png",
                    "content_type": "image/x-custom-png",
                    "content_base64": base64.b64encode(b"custom").decode("ascii"),
                }],
            }],
        },
        "page_layout": {"default_section": {"header_refs": {"default": "rId10"}}},
    }

    result = import_headers_footers(extract, registry, [])

    ct_root = ET.fromstring(ct_path.read_bytes())
    defaults = [
        node for node in ct_root.findall("{*}Default")
        if node.attrib.get("Extension", "").casefold() == "png"
    ]
    assert len(defaults) == 1
    assert defaults[0].attrib["ContentType"] == "image/png"
    media_part = next(iter(result.media_names))
    overrides = {
        node.attrib["PartName"]: node.attrib["ContentType"]
        for node in ct_root.findall("{*}Override")
    }
    assert overrides[f"/{media_part}"] == "image/x-custom-png"


def test_registry_rels_declaration_is_normalized_before_utf8_parse_and_write(tmp_path):
    extract = _seed_extract(tmp_path)
    registry = {
        "headers_footers": {
            "headers": [{
                "part_name": "word/header1.xml",
                "rid": "rId10",
                "xml": (
                    '<w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
                    'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
                    '<w:p><w:hyperlink r:id="rIdLink"/></w:p></w:hdr>'
                ),
                "rels_xml": (
                    '<?xml version="1.0" encoding="UTF-16"?>'
                    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
                    '<Relationship Id="rIdLink" '
                    'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" '
                    'Target="https://example.invalid" TargetMode="External"/>'
                    '</Relationships>'
                ),
            }],
        },
        "page_layout": {"default_section": {"header_refs": {"default": "rId10"}}},
    }

    import_headers_footers(extract, registry, [])

    rels_bytes = (extract / "word" / "_rels" / "header1.xml.rels").read_bytes()
    assert b'encoding="UTF-8"' in rels_bytes
    ET.fromstring(rels_bytes)


def test_imports_relationship_driven_nested_custom_part_and_rels_paths(tmp_path):
    extract = _seed_extract(tmp_path)
    xml = (
        '<w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
        'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" '
        'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">'
        '<w:p><w:r><w:drawing><a:blip r:embed="rIdImage"/>'
        '</w:drawing></w:r></w:p></w:hdr>'
    )
    rels_xml = (
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rIdImage" '
        'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" '
        'Target="../media/brand/logo.png"/></Relationships>'
    )
    media = [{
        "rel_id": "rIdImage",
        "target": "media/brand/logo.png",
        "content_type": "image/png",
        "data_base64": base64.b64encode(b"nested-logo").decode("ascii"),
    }]
    common = {
        "part_name": "word/headers/default.xml",
        "rels_part_name": "word/headers/_rels/default.xml.rels",
        "xml": xml,
        "rels_xml": rels_xml,
        "media": media,
    }
    registry = {
        "headers_footers": {
            "headers": [
                {**common, "rel_id": "rIdArchDefault"},
                {**common, "rel_id": "rIdArchEven"},
            ],
        },
        "page_layout": {
            "default_section": {
                "header_refs": {
                    "default": "rIdArchDefault",
                    "even": "rIdArchEven",
                },
            },
        },
    }

    result = import_headers_footers(extract, registry, [])

    assert result.part_names == {"word/headers/default.xml"}
    assert result.rels_names == {"word/headers/_rels/default.xml.rels"}
    assert (extract / "word" / "headers" / "default.xml").is_file()
    assert (extract / "word" / "headers" / "_rels" / "default.xml.rels").is_file()
    header_rels = ET.fromstring(
        (extract / "word" / "headers" / "_rels" / "default.xml.rels").read_bytes()
    )
    image_rel = next(iter(header_rels))
    assert image_rel.attrib["Target"].startswith("../media/hf_default_")
    imported_media_part = resolve_internal_relationship_target(
        "word/headers/default.xml",
        image_rel.attrib["Target"],
    )
    assert imported_media_part in result.media_names
    assert (extract / imported_media_part).read_bytes() == b"nested-logo"
    document_rels = ET.fromstring(
        (extract / "word" / "_rels" / "document.xml.rels").read_bytes()
    )
    header_rel = next(
        rel for rel in document_rels
        if rel.attrib.get("Type", "").endswith("/header")
    )
    assert header_rel.attrib["Target"] == "headers/default.xml"
    document_xml = (extract / "word" / "document.xml").read_text(encoding="utf-8")
    refs = re.findall(r'<w:headerReference\b[^>]*r:id="([^"]+)"', document_xml)
    assert len(refs) == 2
    assert len(set(refs)) == 1


def test_patch_footer_tokens_handles_split_wt_nodes_and_case_mirroring(tmp_path):
    word_dir = tmp_path / "word"
    word_dir.mkdir(parents=True, exist_ok=True)
    footer_path = word_dir / "footer1.xml"
    footer_path.write_text(
        '<w:ftr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
        '<w:tbl><w:tr><w:tc><w:p>'
        '<w:r><w:t>Metal </w:t></w:r><w:r><w:t>Ducts</w:t></w:r>'
        '<w:r><w:t xml:space="preserve"> </w:t></w:r>'
        '<w:r><w:t>23 31 00</w:t></w:r>'
        '</w:p></w:tc></w:tr></w:tbl>'
        '</w:ftr>',
        encoding="utf-8",
    )
    log = []
    patch_footer_tokens(
        target_extract_dir=tmp_path,
        source_tokens={"SectionTitle": "METAL DUCTS", "SectionID": "SECTION 23 31 00"},
        target_tokens={"SectionTitle": "Direct-Digital Control System for HVAC", "SectionID": "SECTION 23 09 00"},
        log=log,
    )

    out = footer_path.read_text(encoding="utf-8")
    assert "Direct-Digital Control System for HVAC" in out
    assert "23 09 00" in out
    assert "Metal " not in out
    assert any("Patched tokens in footer1.xml" in line for line in log)


def test_import_reports_styles_referenced_by_header_footer_parts(tmp_path):
    extract = _seed_extract(tmp_path)
    registry = {
        "headers_footers": {
            "headers": [{
                "part_name": "word/header1.xml",
                "rid": "rId10",
                "xml": (
                    '<w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
                    '<w:p><w:pPr><w:pStyle w:val="ArchitectHeader"/></w:pPr>'
                    '<w:r><w:rPr><w:rStyle w:val="ArchitectHeaderChar"/></w:rPr>'
                    '<w:t>Header</w:t></w:r></w:p></w:hdr>'
                ),
            }],
        },
        "page_layout": {"default_section": {"header_refs": {"default": "rId10"}}},
    }

    result = import_headers_footers(extract, registry, [])
    assert result.style_ids == {"ArchitectHeader", "ArchitectHeaderChar"}


def test_token_patching_covers_headers_and_footers(tmp_path):
    word_dir = tmp_path / "word"
    word_dir.mkdir(parents=True)
    for name, root_tag in (("header1.xml", "hdr"), ("footer1.xml", "ftr")):
        (word_dir / name).write_text(
            f'<w:{root_tag} xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
            '<w:p><w:r><w:t>METAL DUCTS 23 31 00</w:t></w:r></w:p>'
            f'</w:{root_tag}>',
            encoding="utf-8",
        )

    patch_header_footer_tokens(
        target_extract_dir=tmp_path,
        source_tokens={"SectionTitle": "METAL DUCTS", "SectionID": "SECTION 23 31 00"},
        target_tokens={"SectionTitle": "Air Terminals", "SectionID": "SECTION 23 37 00"},
        log=[],
    )

    for name in ("header1.xml", "footer1.xml"):
        out = (word_dir / name).read_text(encoding="utf-8")
        assert "AIR TERMINALS" in out
        assert "23 37 00" in out
        assert "METAL DUCTS" not in out


def test_token_patching_preserves_nested_textbox_subtree_byte_for_byte(tmp_path):
    custom_part = tmp_path / "word" / "layout" / "custom.xml"
    custom_part.parent.mkdir(parents=True)
    textbox = (
        '<w:drawing><w:txbxContent><w:p><w:r><w:t>METAL DUCTS</w:t></w:r>'
        '</w:p></w:txbxContent></w:drawing>'
    )
    custom_part.write_text(
        '<w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
        f'<w:p><w:r>{textbox}<w:t>METAL DUCTS</w:t></w:r></w:p></w:hdr>',
        encoding="utf-8",
    )

    patch_header_footer_tokens(
        target_extract_dir=tmp_path,
        source_tokens={"SectionTitle": "METAL DUCTS"},
        target_tokens={"SectionTitle": "Air Terminals"},
        log=[],
        part_names=["word/layout/custom.xml"],
    )

    output = custom_part.read_text(encoding="utf-8")
    assert textbox in output
    assert output.count("METAL DUCTS") == 1
    assert "AIR TERMINALS" in output


def test_import_reports_and_remaps_direct_header_numbering(tmp_path):
    extract = _seed_extract(tmp_path)
    registry = {
        "headers_footers": {
            "headers": [{
                "part_name": "word/header1.xml",
                "rid": "rId10",
                "xml": (
                    '<w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
                    '<w:p><w:pPr><w:numPr><w:ilvl w:val="0"/>'
                    '<w:numId w:val="7"/></w:numPr></w:pPr></w:p></w:hdr>'
                ),
            }]
        },
        "page_layout": {"default_section": {"header_refs": {"default": "rId10"}}},
    }
    result = import_headers_footers(extract, registry, [])
    assert result.direct_num_ids == {7}

    remap_header_footer_numids(extract, list(result.part_names), {7: 42}, [])
    out = (extract / "word" / "header1.xml").read_text(encoding="utf-8")
    assert '<w:numId w:val="42"/>' in out


def test_import_rejects_header_part_path_traversal(tmp_path):
    extract = _seed_extract(tmp_path)
    registry = {
        "headers_footers": {
            "headers": [{
                "part_name": "../outside.xml",
                "rid": "rId10",
                "xml": '<w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"/>',
            }]
        },
        "page_layout": {"default_section": {"header_refs": {"default": "rId10"}}},
    }
    with pytest.raises(ValueError, match="Unsafe or invalid header part_name"):
        import_headers_footers(extract, registry, [])


def test_import_validates_single_quoted_relationship_references(tmp_path):
    extract = _seed_extract(tmp_path)
    registry = {
        "headers_footers": {
            "headers": [{
                "part_name": "word/header1.xml",
                "rid": "rId10",
                "xml": (
                    "<w:hdr xmlns:w='http://schemas.openxmlformats.org/wordprocessingml/2006/main' "
                    "xmlns:r='http://schemas.openxmlformats.org/officeDocument/2006/relationships'>"
                    "<w:p><w:hyperlink r:id='rIdMissing'/></w:p></w:hdr>"
                ),
            }]
        },
        "page_layout": {"default_section": {"header_refs": {"default": "rId10"}}},
    }

    with pytest.raises(ValueError, match="has no relationships part"):
        import_headers_footers(extract, registry, [])


def test_nested_header_image_target_is_remapped_by_relationship_id(tmp_path):
    extract = _seed_extract(tmp_path)
    registry = {
        "headers_footers": {
            "headers": [{
                "part_name": "word/header1.xml",
                "rid": "rId10",
                "xml": (
                    '<w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
                    'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" '
                    'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">'
                    '<w:p><w:r><w:drawing><a:blip r:embed="rId7"/></w:drawing></w:r></w:p>'
                    '</w:hdr>'
                ),
                "rels_xml": (
                    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
                    '<Relationship Id="rId7" '
                    'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" '
                    'Target="media/linked/image1.png"/>'
                    '</Relationships>'
                ),
                "media": [{
                    "rel_id": "rId7",
                    "target": "media/linked/image1.png",
                    "data_base64": base64.b64encode(b"nested-image").decode("ascii"),
                }],
            }],
        },
        "page_layout": {"default_section": {"header_refs": {"default": "rId10"}}},
    }

    result = import_headers_footers(extract, registry, [])

    rels_root = ET.fromstring((extract / "word" / "_rels" / "header1.xml.rels").read_bytes())
    relationship = next(iter(rels_root))
    rewritten_target = relationship.attrib["Target"]
    assert rewritten_target.startswith("media/hf_header1_")
    assert (extract / "word" / rewritten_target).read_bytes() == b"nested-image"
    assert f"word/{rewritten_target}" in result.media_names


def test_duplicate_image_target_is_disambiguated_by_relationship_id_and_external_is_preserved(tmp_path):
    extract = _seed_extract(tmp_path)
    registry = {
        "headers_footers": {
            "headers": [{
                "part_name": "word/header1.xml",
                "rid": "rId10",
                "xml": (
                    '<w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
                    'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" '
                    'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">'
                    '<w:p><w:r><w:drawing><a:blip r:embed="rId7"/>'
                    '<a:blip r:embed="rId8"/></w:drawing></w:r>'
                    '<w:hyperlink r:id="rId9"/></w:p></w:hdr>'
                ),
                "rels_xml": (
                    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
                    '<Relationship Id="rId7" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/shared.png"/>'
                    '<Relationship Id="rId8" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/shared.png"/>'
                    '<Relationship Id="rId9" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" '
                    'Target="https://example.invalid/spec" TargetMode="External"/>'
                    '</Relationships>'
                ),
                "media": [
                    {"rel_id": "rId7", "target": "media/shared.png", "content_type": "image/custom", "data_base64": base64.b64encode(b"one").decode("ascii")},
                    {"rel_id": "rId8", "target": "media/shared.png", "content_type": "image/custom", "data_base64": base64.b64encode(b"two").decode("ascii")},
                ],
            }],
        },
        "page_layout": {"default_section": {"header_refs": {"default": "rId10"}}},
    }

    import_headers_footers(extract, registry, [])

    rels_root = ET.fromstring((extract / "word" / "_rels" / "header1.xml.rels").read_bytes())
    rels = {node.attrib["Id"]: node.attrib for node in rels_root}
    assert rels["rId7"]["Target"] != rels["rId8"]["Target"]
    assert (extract / "word" / rels["rId7"]["Target"]).read_bytes() == b"one"
    assert (extract / "word" / rels["rId8"]["Target"]).read_bytes() == b"two"
    assert rels["rId9"]["Target"] == "https://example.invalid/spec"
    assert rels["rId9"]["TargetMode"] == "External"
    assert 'ContentType="image/custom"' in (extract / "[Content_Types].xml").read_text(encoding="utf-8")


def test_import_omits_unreferenced_architect_parts(tmp_path):
    extract = _seed_extract(tmp_path)
    registry = {
        "headers_footers": {
            "headers": [{
                "part_name": "word/header1.xml",
                "rid": "rId10",
                "xml": '<w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"/>',
            }],
            "footers": [{
                "part_name": "word/footer9.xml",
                "rid": "rId99",
                "xml": '<w:ftr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"/>',
            }],
        },
        "page_layout": {"default_section": {"header_refs": {"default": "rId10"}}},
    }

    log = []
    result = import_headers_footers(extract, registry, log)

    assert result.part_names == {"word/header1.xml"}
    assert not (extract / "word" / "footer9.xml").exists()
    assert any("unreferenced" in line for line in log)


@pytest.mark.parametrize("sectpr_xml", ["<w:sectPr/>", ""])
def test_import_rewires_self_closing_or_missing_body_sectpr(tmp_path, sectpr_xml):
    extract = _seed_extract(tmp_path)
    doc_path = extract / "word" / "document.xml"
    original = doc_path.read_text(encoding="utf-8")
    original = re.sub(r"<w:sectPr\b[\s\S]*?</w:sectPr>", sectpr_xml, original)
    doc_path.write_text(original, encoding="utf-8")
    registry = {
        "headers_footers": {
            "headers": [{
                "part_name": "word/header1.xml",
                "rid": "rId10",
                "xml": '<w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"/>',
            }],
        },
        "page_layout": {"default_section": {"header_refs": {"default": "rId10"}}},
    }

    import_headers_footers(extract, registry, [])

    out = doc_path.read_text(encoding="utf-8")
    assert out.count("<w:sectPr") == 1
    assert "headerReference" in out
