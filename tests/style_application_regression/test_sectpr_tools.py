import zipfile

import pytest

from spec_formatter.style_application.core.sectpr_tools import (
    extract_all_sectpr_blocks,
    extract_sectpr_children,
    replace_nth_sectpr_block,
)
from spec_formatter.style_application.phase2_invariants import verify_phase2_invariants


W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
PKG_REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"


def test_sectpr_scanner_supports_self_closing_and_paired_blocks():
    document = (
        '<w:document><w:body><w:p><w:pPr><w:sectPr/></w:pPr></w:p>'
        '<w:sectPr><w:pgBorders><w:top w:val="single"/></w:pgBorders></w:sectPr>'
        '</w:body></w:document>'
    )
    assert extract_all_sectpr_blocks(document) == [
        '<w:sectPr/>',
        '<w:sectPr><w:pgBorders><w:top w:val="single"/></w:pgBorders></w:sectPr>',
    ]


def test_document_sectpr_scanner_excludes_table_and_textbox_subtrees():
    table = '<w:tbl><w:tr><w:tc><w:p><w:pPr><w:sectPr/></w:pPr></w:p></w:tc></w:tr></w:tbl>'
    drawing = (
        '<w:p><w:r><w:drawing><w:txbxContent><w:p><w:pPr>'
        '<w:sectPr><w:pgSz w:w="999"/></w:sectPr>'
        '</w:pPr></w:p></w:txbxContent></w:drawing></w:r></w:p>'
    )
    body_section = '<w:sectPr><w:pgSz w:w="12240"/></w:sectPr>'
    document = f'<w:document><w:body>{table}{drawing}{body_section}</w:body></w:document>'

    assert extract_all_sectpr_blocks(document) == [body_section]
    assert replace_nth_sectpr_block(document, 0, '<w:sectPr><w:pgMar/></w:sectPr>') == (
        f'<w:document><w:body>{table}{drawing}<w:sectPr><w:pgMar/></w:sectPr>'
        '</w:body></w:document>'
    )


def test_replace_nth_sectpr_can_expand_self_closing_block():
    document = '<w:body><w:sectPr/><w:sectPr><w:pgSz/></w:sectPr></w:body>'
    result = replace_nth_sectpr_block(
        document,
        0,
        '<w:sectPr><w:pgMar w:top="1440"/></w:sectPr>',
    )
    assert result == (
        '<w:body><w:sectPr><w:pgMar w:top="1440"/></w:sectPr>'
        '<w:sectPr><w:pgSz/></w:sectPr></w:body>'
    )


def test_extract_sectpr_children_keeps_nested_child_whole():
    inner = (
        '<w:headerReference w:type="default" r:id="rId1"/>'
        '<w:pgBorders><w:top w:val="single"/><w:bottom w:val="double"/></w:pgBorders>'
        '<mc:AlternateContent><mc:Choice Requires="w14">'
        '<w14:extension w14:val="a>b"/></mc:Choice><mc:Fallback><custom/></mc:Fallback>'
        '</mc:AlternateContent>'
        '<w:docGrid w:linePitch="360"/>'
    )
    assert extract_sectpr_children(inner) == [
        '<w:headerReference w:type="default" r:id="rId1"/>',
        '<w:pgBorders><w:top w:val="single"/><w:bottom w:val="double"/></w:pgBorders>',
        '<mc:AlternateContent><mc:Choice Requires="w14">'
        '<w14:extension w14:val="a>b"/></mc:Choice><mc:Fallback><custom/></mc:Fallback>'
        '</mc:AlternateContent>',
        '<w:docGrid w:linePitch="360"/>',
    ]


def test_extract_sectpr_children_rejects_mismatched_extension_tags():
    with pytest.raises(ValueError, match="mismatched closing tag"):
        extract_sectpr_children('<mc:AlternateContent><w14:extension/></mc:Choice>')


def _write_docx(path, document_xml, rels_xml, *, include_header=False):
    with zipfile.ZipFile(path, "w") as archive:
        archive.writestr("word/document.xml", document_xml)
        archive.writestr("word/_rels/document.xml.rels", rels_xml)
        if include_header:
            archive.writestr(
                "word/header1.xml",
                f'<w:hdr xmlns:w="{W_NS}"><w:p/></w:hdr>',
            )


def test_invariants_allow_created_body_sectpr_and_require_expected_reference(tmp_path):
    source = tmp_path / "source.docx"
    output = tmp_path / "output.docx"
    source_document = f'<w:document xmlns:w="{W_NS}" xmlns:r="{R_NS}"><w:body><w:p/></w:body></w:document>'
    output_document = (
        f'<w:document xmlns:w="{W_NS}" xmlns:r="{R_NS}"><w:body><w:p/>'
        '<w:sectPr><w:headerReference w:type="default" r:id="rId9"/></w:sectPr>'
        '</w:body></w:document>'
    )
    source_rels = f'<Relationships xmlns="{PKG_REL_NS}"/>'
    output_rels = (
        f'<Relationships xmlns="{PKG_REL_NS}">'
        f'<Relationship Id="rId9" Type="{R_NS}/header" Target="header1.xml"/>'
        '</Relationships>'
    )
    _write_docx(source, source_document, source_rels)
    _write_docx(output, output_document, output_rels, include_header=True)
    registry = {
        "headers_footers": {
            "headers": [
                {"part_name": "word/header1.xml", "rid": "rIdArchitect"}
            ],
            "footers": [],
        },
        "page_layout": {
            "default_section": {
                "sectPr": '<w:sectPr><w:headerReference w:type="default" r:id="rIdArchitect"/></w:sectPr>',
                "header_refs": {"default": "rIdArchitect"},
                "footer_refs": {},
            },
            "section_chain": [],
        },
    }

    verify_phase2_invariants(
        src_docx=source,
        new_document_xml=output_document.encode("utf-8"),
        new_docx=output,
        arch_template_registry=registry,
    )

    missing_ref_document = output_document.replace(
        '<w:headerReference w:type="default" r:id="rId9"/>',
        "",
    )
    broken = tmp_path / "broken.docx"
    _write_docx(broken, missing_ref_document, output_rels, include_header=True)
    with pytest.raises(RuntimeError, match="references do not match architect mapping"):
        verify_phase2_invariants(
            src_docx=source,
            new_document_xml=missing_ref_document.encode("utf-8"),
            new_docx=broken,
            arch_template_registry=registry,
        )


def test_invariants_allow_body_sectpr_after_existing_paragraph_section(tmp_path):
    source = tmp_path / "paragraph-section.docx"
    paragraph_sectpr = '<w:sectPr><w:type w:val="continuous"/><w:pgSz w:w="12000"/></w:sectPr>'
    source_document = (
        f'<w:document xmlns:w="{W_NS}"><w:body><w:p><w:pPr>{paragraph_sectpr}'
        '</w:pPr><w:r><w:t>Visible</w:t></w:r></w:p></w:body></w:document>'
    )
    output_document = source_document.replace(
        '</w:body>',
        '<w:sectPr><w:pgSz w:w="12240"/><w:pgMar w:top="1440"/></w:sectPr></w:body>',
    )
    _write_docx(source, source_document, f'<Relationships xmlns="{PKG_REL_NS}"/>')

    verify_phase2_invariants(
        src_docx=source,
        new_document_xml=output_document.encode("utf-8"),
    )


def test_invariants_reject_loss_of_unmanaged_section_semantics(tmp_path):
    source = tmp_path / "unmanaged-section.docx"
    source_document = (
        f'<w:document xmlns:w="{W_NS}"><w:body><w:p/>'
        '<w:sectPr><w:type w:val="nextPage"/><w:pgNumType w:start="3"/>'
        '<w:vAlign w:val="center"/><w:pgSz w:w="12000"/></w:sectPr>'
        '</w:body></w:document>'
    )
    changed = source_document.replace('<w:pgNumType w:start="3"/>', '')
    _write_docx(source, source_document, f'<Relationships xmlns="{PKG_REL_NS}"/>')

    with pytest.raises(RuntimeError, match="non-layout sectPr semantics changed"):
        verify_phase2_invariants(
            src_docx=source,
            new_document_xml=changed.encode("utf-8"),
        )


def test_orphan_architect_hf_entries_preserve_target_parts_and_relationships(tmp_path):
    source = tmp_path / "orphan-source.docx"
    output = tmp_path / "orphan-output.docx"
    document = (
        f'<w:document xmlns:w="{W_NS}" xmlns:r="{R_NS}"><w:body><w:p/>'
        '<w:sectPr><w:headerReference w:type="default" r:id="rIdTarget"/></w:sectPr>'
        '</w:body></w:document>'
    )
    rels = (
        f'<Relationships xmlns="{PKG_REL_NS}">'
        f'<Relationship Id="rIdTarget" Type="{R_NS}/header" Target="header1.xml"/>'
        '</Relationships>'
    )
    _write_docx(source, document, rels, include_header=True)
    _write_docx(output, document, rels, include_header=True)
    registry = {
        "headers_footers": {
            "headers": [{"part_name": "word/headerOrphan.xml", "rid": "rIdOrphan"}],
            "footers": [],
        },
        "page_layout": {
            "default_section": {"sectPr": "<w:sectPr/>", "header_refs": {}, "footer_refs": {}},
            "section_chain": [],
        },
    }

    verify_phase2_invariants(
        src_docx=source,
        new_document_xml=document.encode("utf-8"),
        new_docx=output,
        arch_template_registry=registry,
    )
