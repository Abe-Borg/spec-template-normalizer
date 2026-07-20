import zipfile
from pathlib import Path

import pytest

from spec_formatter.style_application.batch_runner import _build_and_patch_output
from spec_formatter.style_application.phase2_invariants import validate_docx_package, verify_phase2_invariants


W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
PKG_REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"


def _parts(*, style_id="Body", num_id=1, document_target="word/document.xml"):
    return {
        "[Content_Types].xml": (
            '<?xml version="1.0" encoding="UTF-8"?>'
            '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
            '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
            '<Default Extension="xml" ContentType="application/xml"/>'
            '</Types>'
        ),
        "_rels/.rels": (
            f'<Relationships xmlns="{PKG_REL_NS}">'
            f'<Relationship Id="rId1" Type="{R_NS}/officeDocument" Target="{document_target}"/>'
            '</Relationships>'
        ),
        "word/document.xml": (
            f'<w:document xmlns:w="{W_NS}" xmlns:r="{R_NS}"><w:body>'
            f'<w:p><w:pPr><w:pStyle w:val="{style_id}"/>'
            f'<w:numPr><w:ilvl w:val="0"/><w:numId w:val="{num_id}"/></w:numPr>'
            '</w:pPr><w:r><w:t>Text</w:t></w:r></w:p><w:sectPr/></w:body></w:document>'
        ),
        "word/_rels/document.xml.rels": (
            f'<Relationships xmlns="{PKG_REL_NS}">'
            f'<Relationship Id="rId1" Type="{R_NS}/styles" Target="styles.xml"/>'
            f'<Relationship Id="rId2" Type="{R_NS}/numbering" Target="numbering.xml"/>'
            '</Relationships>'
        ),
        "word/styles.xml": (
            f'<w:styles xmlns:w="{W_NS}">'
            '<w:style w:type="paragraph" w:styleId="Body"><w:name w:val="Body"/></w:style>'
            '</w:styles>'
        ),
        "word/numbering.xml": (
            f'<w:numbering xmlns:w="{W_NS}">'
            '<w:abstractNum w:abstractNumId="2"><w:lvl w:ilvl="0"/></w:abstractNum>'
            '<w:num w:numId="1"><w:abstractNumId w:val="2"/></w:num>'
            '</w:numbering>'
        ),
    }


def _write_docx(path: Path, parts):
    with zipfile.ZipFile(path, "w") as zf:
        for name, value in parts.items():
            zf.writestr(name, value)


def test_valid_package_passes_final_validation(tmp_path):
    docx = tmp_path / "valid.docx"
    _write_docx(docx, _parts())
    validate_docx_package(docx)


@pytest.mark.parametrize(
    ("kwargs", "message"),
    [
        ({"document_target": "word/missing.xml"}, "targets missing part"),
        ({"style_id": "MissingStyle"}, "unresolved style reference"),
        ({"num_id": 99}, "unresolved numId reference"),
    ],
)
def test_final_validation_rejects_broken_references(tmp_path, kwargs, message):
    docx = tmp_path / "broken.docx"
    _write_docx(docx, _parts(**kwargs))
    with pytest.raises(RuntimeError, match=message):
        validate_docx_package(docx)


def test_final_validation_rejects_part_without_content_type(tmp_path):
    docx = tmp_path / "broken.docx"
    parts = _parts()
    parts["word/media/logo.unknown"] = b"data"
    _write_docx(docx, parts)
    with pytest.raises(RuntimeError, match="has no content type"):
        validate_docx_package(docx)


def test_final_validation_rejects_case_insensitive_duplicate_members(tmp_path):
    docx = tmp_path / "broken.docx"
    parts = _parts()
    _write_docx(docx, parts)
    with zipfile.ZipFile(docx, "a") as archive:
        archive.writestr("WORD/DOCUMENT.XML", parts["word/document.xml"])

    with pytest.raises(RuntimeError, match="case-insensitive duplicate ZIP members"):
        validate_docx_package(docx)


def test_final_validation_rejects_relationship_without_type(tmp_path):
    docx = tmp_path / "missing-rel-type.docx"
    parts = _parts()
    parts["_rels/.rels"] = parts["_rels/.rels"].replace(
        f' Type="{R_NS}/officeDocument"',
        "",
    )
    _write_docx(docx, parts)

    with pytest.raises(RuntimeError, match="missing Id, Type, or Target"):
        validate_docx_package(docx)


def test_final_validation_rejects_invalid_relationship_target_mode(tmp_path):
    docx = tmp_path / "bad-target-mode.docx"
    parts = _parts()
    parts["_rels/.rels"] = parts["_rels/.rels"].replace(
        ' Target="word/document.xml"',
        ' Target="word/document.xml" TargetMode="external"',
    )
    _write_docx(docx, parts)

    with pytest.raises(RuntimeError, match="invalid TargetMode"):
        validate_docx_package(docx)


def test_final_validation_rejects_case_insensitive_duplicate_overrides(tmp_path):
    docx = tmp_path / "duplicate-overrides.docx"
    parts = _parts()
    parts["[Content_Types].xml"] = parts["[Content_Types].xml"].replace(
        "</Types>",
        '<Override PartName="/word/document.xml" ContentType="application/xml"/>'
        '<Override PartName="/WORD/DOCUMENT.XML" ContentType="application/xml"/>'
        "</Types>",
    )
    _write_docx(docx, parts)

    with pytest.raises(RuntimeError, match="duplicate content-type Override"):
        validate_docx_package(docx)


@pytest.mark.parametrize("relationship_attribute", ["id", "embed", "link"])
def test_final_validation_rejects_unresolved_relationship_attributes(
    tmp_path,
    relationship_attribute,
):
    docx = tmp_path / "broken-header-ref.docx"
    parts = _parts()
    parts["word/header1.xml"] = (
        f'<w:hdr xmlns:w="{W_NS}" xmlns:r="{R_NS}" '
        'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">'
        f'<w:p><w:r><w:drawing><a:blip r:{relationship_attribute}="rIdMissing"/>'
        '</w:drawing></w:r></w:p></w:hdr>'
    )
    _write_docx(docx, parts)
    with pytest.raises(RuntimeError, match="relationship references exist"):
        validate_docx_package(docx)


@pytest.mark.parametrize("relationship_attribute", ["id", "embed", "link"])
def test_final_validation_rejects_empty_relationship_attributes(
    tmp_path,
    relationship_attribute,
):
    docx = tmp_path / "empty-header-ref.docx"
    parts = _parts()
    parts["word/header1.xml"] = (
        f'<w:hdr xmlns:w="{W_NS}" xmlns:r="{R_NS}" '
        'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">'
        f'<w:p><w:r><w:drawing><a:blip r:{relationship_attribute}="  "/>'
        '</w:drawing></w:r></w:p></w:hdr>'
    )
    _write_docx(docx, parts)

    with pytest.raises(RuntimeError, match=rf"empty r:{relationship_attribute}"):
        validate_docx_package(docx)


def test_output_build_preserves_source_header_when_architect_has_none(tmp_path):
    source = tmp_path / "source.docx"
    parts = _parts()
    parts["word/header1.xml"] = f'<w:hdr xmlns:w="{W_NS}"><w:p/></w:hdr>'
    header_rels = (
        f'<Relationships xmlns="{PKG_REL_NS}">'
        f'<Relationship Id="rId1" Type="{R_NS}/image" Target="media/header-logo.png"/>'
        "</Relationships>"
    )
    header_media = b"unchanged-header-media"
    parts["word/_rels/header1.xml.rels"] = header_rels
    parts["word/media/header-logo.png"] = header_media
    parts["[Content_Types].xml"] = parts["[Content_Types].xml"].replace(
        "</Types>",
        '<Default Extension="png" ContentType="image/png"/></Types>',
    )
    parts["word/_rels/document.xml.rels"] = parts["word/_rels/document.xml.rels"].replace(
        "</Relationships>",
        f'<Relationship Id="rId3" Type="{R_NS}/header" Target="header1.xml"/>'
        "</Relationships>",
    )
    parts["word/document.xml"] = parts["word/document.xml"].replace(
        "<w:sectPr/>",
        '<w:sectPr><w:headerReference w:type="default" r:id="rId3"/></w:sectPr>',
    )
    _write_docx(source, parts)

    extract = tmp_path / "extract"
    (extract / "word").mkdir(parents=True)
    (extract / "word" / "document.xml").write_text(parts["word/document.xml"], encoding="utf-8")
    (extract / "word" / "styles.xml").write_text(parts["word/styles.xml"], encoding="utf-8")
    drifted_header = f'<w:hdr xmlns:w="{W_NS}"><w:p><w:r><w:t>DRIFTED</w:t></w:r></w:p></w:hdr>'
    (extract / "word" / "header1.xml").write_text(drifted_header, encoding="utf-8")

    output = _build_and_patch_output(
        source,
        extract,
        {"header_footer_import": {"part_names": set(), "rels_names": set(), "media_names": set()}},
        tmp_path / "out",
        arch_template_registry={},
    )

    with zipfile.ZipFile(output) as zf:
        assert "word/header1.xml" in zf.namelist()
        assert zf.read("word/header1.xml").decode("utf-8") == parts["word/header1.xml"]
        assert zf.read("word/_rels/header1.xml.rels").decode("utf-8") == header_rels
        assert zf.read("word/media/header-logo.png") == header_media


def test_preservation_invariant_follows_custom_header_relationship_target(tmp_path):
    source = tmp_path / "source-custom-header.docx"
    output = tmp_path / "output-custom-header.docx"
    parts = _parts()
    custom_part = "word/layout/footer-looking.xml"
    original_header = f'<w:hdr xmlns:w="{W_NS}"><w:p/></w:hdr>'
    parts[custom_part] = original_header
    parts["word/_rels/document.xml.rels"] = parts[
        "word/_rels/document.xml.rels"
    ].replace(
        "</Relationships>",
        f'<Relationship Id="rId3" Type="{R_NS}/header" '
        'Target="layout/footer-looking.xml"/></Relationships>',
    )
    parts["word/document.xml"] = parts["word/document.xml"].replace(
        "<w:sectPr/>",
        '<w:sectPr><w:headerReference w:type="default" r:id="rId3"/></w:sectPr>',
    )
    _write_docx(source, parts)

    output_parts = dict(parts)
    output_parts[custom_part] = original_header.replace("<w:p/>", "<w:p><w:r/></w:p>")
    _write_docx(output, output_parts)

    with pytest.raises(RuntimeError, match="target header/footer part changed"):
        verify_phase2_invariants(
            source,
            output_parts["word/document.xml"].encode("utf-8"),
            new_docx=output,
            arch_template_registry={},
        )
