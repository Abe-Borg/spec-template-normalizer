from __future__ import annotations

import codecs
import hashlib
import json
import re
import zipfile
from pathlib import Path
import xml.etree.ElementTree as ET

from spec_formatter.pipeline import CSI_TO_CANADIAN, format_specifications
from spec_formatter.style_application.phase2_invariants import validate_docx_package


W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"
REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
CT_NS = "http://schemas.openxmlformats.org/package/2006/content-types"
W14_NS = "http://schemas.microsoft.com/office/word/2010/wordml"
MC_NS = "http://schemas.openxmlformats.org/markup-compatibility/2006"

PNG_1X1 = bytes.fromhex(
    "89504e470d0a1a0a0000000d49484452000000010000000108060000001f15c489"
    "0000000d49444154789c6360f8cfc00000040101005f8d3a0000000049454e44ae426082"
)

TABLE_BLOCK = (
    "<w:tbl><w:tr><w:tc>"
    "<w:p><w:r><w:t>Outer cell</w:t></w:r></w:p>"
    "<w:tbl><w:tr><w:tc><w:p><w:r><w:t>Inner cell</w:t></w:r></w:p>"
    "</w:tc></w:tr></w:tbl>"
    "</w:tc></w:tr></w:tbl>"
)
TEXTBOX_BLOCK = (
    "<w:p><w:r><w:drawing><w:txbxContent>"
    "<w:p><w:r><w:t>Text box text</w:t></w:r></w:p>"
    "</w:txbxContent></w:drawing></w:r></w:p>"
)


def _utf16_xml(body: str, *, little_endian: bool) -> bytes:
    text = f'<?xml version="1.0" encoding="UTF-16"?>{body}'
    if little_endian:
        return codecs.BOM_UTF16_LE + text.encode("utf-16-le")
    return codecs.BOM_UTF16_BE + text.encode("utf-16-be")


def _content_types(header_names: list[str], footer_names: list[str]) -> str:
    overrides = [
        ("/word/document.xml", "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"),
        ("/word/styles.xml", "application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"),
        ("/word/settings.xml", "application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"),
        ("/word/numbering.xml", "application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"),
    ]
    overrides.extend(
        (f"/word/{name}", "application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml")
        for name in header_names
    )
    overrides.extend(
        (f"/word/{name}", "application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml")
        for name in footer_names
    )
    return (
        f'<Types xmlns="{CT_NS}">'
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
        '<Default Extension="xml" ContentType="application/xml"/>'
        '<Default Extension="png" ContentType="image/png"/>'
        + "".join(
            f'<Override PartName="{part}" ContentType="{content_type}"/>'
            for part, content_type in overrides
        )
        + "</Types>"
    )


def _root_relationships() -> str:
    return (
        f'<Relationships xmlns="{REL_NS}">'
        '<Relationship Id="rId1" '
        'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" '
        'Target="word/document.xml"/>'
        "</Relationships>"
    )


def _styles_xml(*, target: bool) -> str:
    inherited = (
        '<w:style w:type="paragraph" w:styleId="TargetLevel2">'
        '<w:name w:val="Target Level 2"/><w:basedOn w:val="Normal"/>'
        '<w:pPr><w:numPr><w:ilvl w:val="2"/><w:numId w:val="17"/>'
        '</w:numPr></w:pPr></w:style>'
        if target
        else ""
    )
    return (
        f'<w:styles xmlns:w="{W_NS}">'
        '<w:docDefaults><w:rPrDefault><w:rPr/></w:rPrDefault>'
        '<w:pPrDefault><w:pPr/></w:pPrDefault></w:docDefaults>'
        '<w:style w:type="paragraph" w:default="1" w:styleId="Normal">'
        '<w:name w:val="Normal"/><w:qFormat/></w:style>'
        f"{inherited}</w:styles>"
    )


def _numbering_xml(*, target: bool) -> str:
    abstract_id = "42" if target else "5"
    num_id = "17" if target else "5"
    return (
        f'<w:numbering xmlns:w="{W_NS}">'
        f'<w:abstractNum w:abstractNumId="{abstract_id}">'
        '<w:multiLevelType w:val="multilevel"/>'
        '<w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="upperLetter"/>'
        '<w:lvlText w:val="%1."/><w:pPr><w:ind w:left="720" w:hanging="360"/></w:pPr></w:lvl>'
        '<w:lvl w:ilvl="2"><w:start w:val="1"/><w:numFmt w:val="decimal"/>'
        '<w:lvlText w:val="%3)"/><w:pPr><w:ind w:left="1440" w:hanging="360"/></w:pPr></w:lvl>'
        '</w:abstractNum>'
        f'<w:num w:numId="{num_id}"><w:abstractNumId w:val="{abstract_id}"/></w:num>'
        '</w:numbering>'
    )


def _section_xml(*, architect: bool, first: bool) -> str:
    if architect:
        refs = (
            '<w:headerReference w:type="default" r:id="rIdHdrDefault"/>'
            '<w:headerReference w:type="first" r:id="rIdHdrFirst"/>'
            '<w:headerReference w:type="even" r:id="rIdHdrEven"/>'
            '<w:footerReference w:type="default" r:id="rIdFtrDefault"/>'
        )
        size = '<w:pgSz w:w="10000" w:h="15000"/>'
        borders = '<w:pgBorders w:offsetFrom="edge"/>'
    else:
        refs = '<w:headerReference w:type="default" r:id="rIdOldHeader"/>'
        size = '<w:pgSz w:w="12000" w:h="16000"/>'
        borders = (
            '<w:pgBorders w:offsetFrom="page"/>'
            if first
            else '<w:pgBorders w:offsetFrom="text"/>'
        )
    return (
        f'<w:sectPr>{refs}{size}'
        '<w:pgMar w:top="720" w:right="720" w:bottom="720" w:left="720"/>'
        f'{borders}<w:cols w:num="1"/><w:docGrid w:linePitch="360"/>'
        + ('<w:titlePg/>' if first else '')
        + '</w:sectPr>'
    )


def _document_xml(*, architect: bool) -> str:
    first_num_id = "5" if architect else "17"
    first_ppr = (
        '<w:pPr><w:numPr><w:ilvl w:val="0"/>'
        f'<w:numId w:val="{first_num_id}"/></w:numPr></w:pPr>'
    )
    if architect:
        second_ppr = (
            '<w:pPr><w:numPr><w:ilvl w:val="2"/><w:numId w:val="5"/>'
            '</w:numPr></w:pPr>'
        )
    else:
        second_ppr = '<w:pPr><w:pStyle w:val="TargetLevel2"/></w:pPr>'
    return (
        '<?xml version="1.0" encoding="UTF-8"?>'
        f'<w:document xmlns:w="{W_NS}" xmlns:r="{R_NS}"><w:body>'
        f'<w:p>{first_ppr}<w:r><w:t>Architect paragraph one</w:t></w:r></w:p>'
        f'<w:p><w:pPr>{_section_xml(architect=architect, first=True)}</w:pPr></w:p>'
        f'<w:p>{second_ppr}<w:r><w:t>Architect paragraph two</w:t></w:r></w:p>'
        f'{TABLE_BLOCK}{TEXTBOX_BLOCK}'
        f'{_section_xml(architect=architect, first=False)}'
        '</w:body></w:document>'
    )


def _base_document_relationships(*, architect: bool) -> str:
    relationships = [
        '<Relationship Id="rIdStyles" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>',
        '<Relationship Id="rIdSettings" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/>',
        '<Relationship Id="rIdNumbering" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering" Target="numbering.xml"/>',
    ]
    if architect:
        relationships.extend(
            [
                '<Relationship Id="rIdHdrDefault" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header1.xml"/>',
                '<Relationship Id="rIdHdrFirst" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header2.xml"/>',
                '<Relationship Id="rIdHdrEven" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header3.xml"/>',
                '<Relationship Id="rIdFtrDefault" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer" Target="footer1.xml"/>',
            ]
        )
    else:
        relationships.append(
            '<Relationship Id="rIdOldHeader" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header9.xml"/>'
        )
    return f'<Relationships xmlns="{REL_NS}">{"".join(relationships)}</Relationships>'


def _architect_header_parts() -> dict[str, bytes | str]:
    header1 = (
        f'<w:hdr xmlns:w="{W_NS}" xmlns:r="{R_NS}" xmlns:a="{A_NS}">'
        '<w:p><w:r><w:drawing><a:blip r:embed="rIdImage"/></w:drawing></w:r>'
        '<w:hyperlink r:id="rIdLink"><w:r><w:t>Architect link</w:t></w:r>'
        '</w:hyperlink></w:p></w:hdr>'
    )
    header1_rels = (
        f'<Relationships xmlns="{REL_NS}">'
        '<Relationship Id="rIdImage" '
        'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" '
        'Target="media/nested/logo.png"/>'
        '<Relationship Id="rIdLink" '
        'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" '
        'Target="https://example.com/spec" TargetMode="External"/>'
        '</Relationships>'
    )
    return {
        "word/header1.xml": header1,
        "word/_rels/header1.xml.rels": header1_rels,
        "word/header2.xml": f'<w:hdr xmlns:w="{W_NS}"><w:p><w:r><w:t>First header</w:t></w:r></w:p></w:hdr>',
        "word/header3.xml": f'<w:hdr xmlns:w="{W_NS}"><w:p><w:r><w:t>Even header</w:t></w:r></w:p></w:hdr>',
        "word/footer1.xml": f'<w:ftr xmlns:w="{W_NS}"><w:p><w:r><w:t>Default footer</w:t></w:r></w:p></w:ftr>',
        "word/media/nested/logo.png": PNG_1X1,
    }


def _write_docx(path: Path, *, architect: bool) -> None:
    if architect:
        headers = ["header1.xml", "header2.xml", "header3.xml"]
        footers = ["footer1.xml"]
        settings = _utf16_xml(
            f'<w:settings xmlns:w="{W_NS}"><w:evenAndOddHeaders/>'
            '<w:zoom w:percent="95"/><w:compat>'
            '<w:compatSetting w:name="compatibilityMode" '
            'w:uri="http://schemas.microsoft.com/office/word" w:val="15"/>'
            '</w:compat></w:settings>',
            little_endian=True,
        )
        extra_parts = _architect_header_parts()
    else:
        headers = ["header9.xml"]
        footers = []
        settings = _utf16_xml(
            f'<w:settings xmlns:w="{W_NS}"><w:zoom w:percent="110"/></w:settings>',
            little_endian=False,
        )
        extra_parts = {
            "word/header9.xml": (
                f'<w:hdr xmlns:w="{W_NS}"><w:p><w:r><w:t>Old target header</w:t>'
                '</w:r></w:p></w:hdr>'
            )
        }

    parts: dict[str, bytes | str] = {
        "[Content_Types].xml": _content_types(headers, footers),
        "_rels/.rels": _root_relationships(),
        "word/document.xml": _document_xml(architect=architect),
        "word/_rels/document.xml.rels": _base_document_relationships(architect=architect),
        "word/styles.xml": _styles_xml(target=not architect),
        "word/settings.xml": settings,
        "word/numbering.xml": _numbering_xml(target=not architect),
        **extra_parts,
    }
    with zipfile.ZipFile(path, "w", compression=zipfile.ZIP_DEFLATED) as package:
        for name, payload in parts.items():
            package.writestr(name, payload)


def _rewrite_docx_parts(path: Path, replacements: dict[str, str]) -> None:
    with zipfile.ZipFile(path, "r") as package:
        parts = {name: package.read(name) for name in package.namelist()}
    for name, text in replacements.items():
        parts[name] = text.encode("utf-8")
    with zipfile.ZipFile(path, "w", compression=zipfile.ZIP_DEFLATED) as package:
        for name, payload in parts.items():
            package.writestr(name, payload)


def _deterministic_classifier(**kwargs):
    classifiable = [
        item for item in kwargs["slim_bundle"].get("paragraphs", [])
        if item.get("skip_reason") is None
    ]
    if len(classifiable) != 2:
        raise AssertionError(
            "expected two classifiable architect paragraphs, got "
            + json.dumps(classifiable, sort_keys=True)
        )
    role_rows = [
        (classifiable[0], "PARAGRAPH", "CSI_Paragraph__ARCH", "CSI Paragraph"),
        (classifiable[1], "SUBPARAGRAPH", "CSI_Subparagraph__ARCH", "CSI Subparagraph"),
    ]
    return {
        "create_styles": [
            {
                "styleId": style_id,
                "name": name,
                "type": "paragraph",
                "derive_from_paragraph_index": paragraph["paragraph_index"],
                "basedOn": "Normal",
                "role": role,
            }
            for paragraph, role, style_id, name in role_rows
        ],
        "apply_pStyle": [
            {"paragraph_index": paragraph["paragraph_index"], "styleId": style_id}
            for paragraph, _role, style_id, _name in role_rows
        ],
        "ignored_paragraphs": [],
        "roles": {
            role: {
                "styleId": style_id,
                "exemplar_paragraph_index": paragraph["paragraph_index"],
            }
            for paragraph, role, style_id, _name in role_rows
        },
        "notes": ["deterministic cross-repository fixture"],
    }


def _sha256(path: Path) -> str:
    return hashlib.sha256(path.read_bytes()).hexdigest()


def _xml_text_sequence(payload: bytes) -> list[str]:
    root = ET.fromstring(payload)
    return [node.text or "" for node in root.iter(f"{{{W_NS}}}t")]


def _relationship_map(payload: bytes) -> dict[str, ET.Element]:
    root = ET.fromstring(payload)
    return {
        rel.attrib["Id"]: rel
        for rel in root.findall(f"{{{REL_NS}}}Relationship")
    }


def test_unified_formatter_round_trips_without_api(tmp_path: Path) -> None:
    architect = tmp_path / "architect.docx"
    target = tmp_path / "target.docx"
    _write_docx(architect, architect=True)
    _write_docx(target, architect=False)

    architect_sha = _sha256(architect)
    target_sha = _sha256(target)
    with zipfile.ZipFile(target) as package:
        target_document_before = package.read("word/document.xml")

    run = format_specifications(
        architect_template=architect,
        target_specs=[target],
        output_dir=tmp_path / "formatted",
        cache_dir=tmp_path / "template-cache",
        api_key="",
        max_workers=1,
        template_model="unified-roundtrip-fixture",
        template_classifier=_deterministic_classifier,
    )
    assert run.success
    assert run.template_profile.reused is False
    bundle = run.template_profile.bundle_dir
    assert _sha256(architect) == architect_sha

    manifest = json.loads((bundle / "phase1_bundle_manifest.json").read_text(encoding="utf-8"))
    registry = json.loads((bundle / "arch_template_registry.json").read_text(encoding="utf-8"))
    style_registry = json.loads((bundle / "arch_style_registry.json").read_text(encoding="utf-8"))
    assert manifest["source"]["sha256"] == architect_sha
    assert (bundle / "source_settings.xml").read_bytes().startswith(codecs.BOM_UTF16_LE)
    assert '<w:settings xmlns:w="' in registry["settings"]["settings_xml"]
    assert 'encoding="UTF-8"' in registry["settings"]["settings_xml"]
    assert set(style_registry["roles"]) == {"PARAGRAPH", "SUBPARAGRAPH"}

    assert registry["headers_footers"]["header_footer_media"] == [
        "media/nested/logo.png"
    ]
    header1 = next(
        item for item in registry["headers_footers"]["headers"]
        if item["part_name"] == "word/header1.xml"
    )
    assert any(
        item.get("target") == "media/nested/logo.png"
        and item.get("rel_id") == "rIdImage"
        for item in header1["media"]
    )
    assert 'Target="https://example.com/spec" TargetMode="External"' in header1["rels_xml"]

    result = run.targets[0]
    assert result.success, "\n".join(result.log)
    assert result.output_path is not None and result.output_path.is_file()
    assert result.output_path.name == "target_FORMATTED.docx"
    assert any("Anthropic request skipped" in line for line in result.log)
    assert any("Application coverage: 2/2 (100.0%)" in line for line in result.log)
    assert _sha256(target) == target_sha

    output = result.output_path
    validate_docx_package(output)
    with zipfile.ZipFile(output) as package:
        names = set(package.namelist())
        output_document = package.read("word/document.xml")
        output_document_text = output_document.decode("utf-8")
        output_settings = package.read("word/settings.xml")
        output_numbering = package.read("word/numbering.xml").decode("utf-8")
        document_rels = _relationship_map(package.read("word/_rels/document.xml.rels"))
        header1_rels = _relationship_map(package.read("word/_rels/header1.xml.rels"))

        assert "word/header9.xml" not in names
        assert {
            "word/header1.xml",
            "word/header2.xml",
            "word/header3.xml",
            "word/footer1.xml",
        } <= names
        imported_image_part = "word/" + header1_rels["rIdImage"].attrib["Target"]
        assert imported_image_part in names
        assert package.read(imported_image_part) == PNG_1X1

    assert _xml_text_sequence(output_document) == _xml_text_sequence(target_document_before)
    assert TABLE_BLOCK in target_document_before.decode("utf-8")
    assert TABLE_BLOCK in output_document_text
    assert TEXTBOX_BLOCK in target_document_before.decode("utf-8")
    assert TEXTBOX_BLOCK in output_document_text

    applied_style_ids = re.findall(r'<w:pStyle w:val="([^"]+)"', output_document_text)
    assert sum(
        style_id.startswith("SF_") and "_BODY_CSI_Paragraph__ARCH_" in style_id
        for style_id in applied_style_ids
    ) == 1
    assert sum(
        style_id.startswith("SF_") and "_BODY_CSI_Subparagraph__ARCH_" in style_id
        for style_id in applied_style_ids
    ) == 1
    assert output_settings.decode("utf-8").startswith('<?xml version="1.0" encoding="UTF-8"?>')
    assert 'w:percent="110"' in output_settings.decode("utf-8")
    assert 'w:name="compatibilityMode"' in output_settings.decode("utf-8")
    assert '<w:numFmt w:val="upperLetter"/>' in output_numbering
    assert '<w:lvlText w:val="%1."/>' in output_numbering
    assert '<w:numFmt w:val="decimal"/>' in output_numbering
    assert '<w:lvlText w:val="%3)"/>' in output_numbering

    image_rel = header1_rels["rIdImage"]
    assert image_rel.attrib["Target"].startswith("media/hf_header1_")
    link_rel = header1_rels["rIdLink"]
    assert link_rel.attrib["Target"] == "https://example.com/spec"
    assert link_rel.attrib["TargetMode"] == "External"

    section_blocks = re.findall(r'<w:sectPr\b[^>]*>[\s\S]*?</w:sectPr>', output_document_text)
    assert len(section_blocks) == 2
    for index, section in enumerate(section_blocks):
        assert '<w:pgSz w:w="10000" w:h="15000"/>' in section
        expected_border = "page" if index == 0 else "text"
        assert f'<w:pgBorders w:offsetFrom="{expected_border}"/>' in section
        refs = re.findall(
            r'<w:(header|footer)Reference\b[^>]*w:type="([^"]+)"[^>]*r:id="([^"]+)"',
            section,
        )
        assert {(kind, ref_type) for kind, ref_type, _rid in refs} == {
            ("header", "default"),
            ("header", "first"),
            ("header", "even"),
            ("footer", "default"),
        }
        resolved_parts = {
            (kind, ref_type): document_rels[rid].attrib["Target"]
            for kind, ref_type, rid in refs
        }
        assert resolved_parts == {
            ("header", "default"): "header1.xml",
            ("header", "first"): "header2.xml",
            ("header", "even"): "header3.xml",
            ("footer", "default"): "footer1.xml",
        }


def test_unified_formatter_keeps_valid_output_when_another_target_is_corrupt(
    tmp_path: Path,
) -> None:
    architect = tmp_path / "architect.docx"
    valid_target = tmp_path / "valid-target.docx"
    corrupt_target = tmp_path / "corrupt-target.docx"
    output_dir = tmp_path / "formatted"
    _write_docx(architect, architect=True)
    _write_docx(valid_target, architect=False)
    corrupt_target.write_bytes(b"this is not an OOXML package")

    input_hashes = {
        path: _sha256(path)
        for path in (architect, valid_target, corrupt_target)
    }

    run = format_specifications(
        architect_template=architect,
        target_specs=[valid_target, corrupt_target],
        output_dir=output_dir,
        cache_dir=tmp_path / "template-cache",
        api_key="",
        max_workers=2,
        template_model="unified-partial-success-fixture",
        template_classifier=_deterministic_classifier,
    )

    assert run.success is False
    assert run.succeeded == 1
    assert run.failed == 1
    results = {result.source_path: result for result in run.targets}

    valid_result = results[valid_target.resolve()]
    assert valid_result.success, "\n".join(valid_result.log)
    assert valid_result.error is None
    assert run.run_dir is not None
    assert valid_result.output_path == run.run_dir / "valid-target_FORMATTED.docx"
    assert valid_result.output_path.is_file()
    validate_docx_package(valid_result.output_path)

    corrupt_result = results[corrupt_target.resolve()]
    assert corrupt_result.success is False
    assert corrupt_result.output_path is None
    assert corrupt_result.error
    assert not (run.run_dir / "corrupt-target_FORMATTED.docx").exists()
    assert list(run.run_dir.glob("*_FORMATTED.docx")) == [valid_result.output_path]

    assert {
        path: _sha256(path)
        for path in (architect, valid_target, corrupt_target)
    } == input_hashes


def test_format_only_preserves_target_automatic_numbering_for_typed_architect_role(
    tmp_path: Path,
) -> None:
    architect = tmp_path / "typed-numbering-architect.docx"
    target = tmp_path / "automatic-numbering-target.docx"
    _write_docx(architect, architect=True)
    _write_docx(target, architect=False)

    with zipfile.ZipFile(architect) as package:
        architect_document = package.read("word/document.xml").decode("utf-8")
    architect_document = re.sub(
        r'(<w:body><w:p>)<w:pPr><w:numPr>[\s\S]*?</w:numPr></w:pPr>',
        r"\1",
        architect_document,
        count=1,
    ).replace(
        "Architect paragraph one",
        "A. Architect paragraph one",
        1,
    )
    _rewrite_docx_parts(
        architect,
        {"word/document.xml": architect_document},
    )

    with zipfile.ZipFile(target) as package:
        target_document_before = package.read("word/document.xml")
    target_sha = _sha256(target)

    run = format_specifications(
        architect_template=architect,
        target_specs=[target],
        output_dir=tmp_path / "formatted",
        cache_dir=tmp_path / "template-cache",
        api_key="",
        max_workers=1,
        template_model="typed-numbering-format-only-fixture",
        template_classifier=_deterministic_classifier,
    )

    assert run.success, "\n".join(run.targets[0].log)
    result = run.targets[0]
    assert result.output_path is not None
    assert _sha256(target) == target_sha
    assert any("Preserved source Word numbering" in line for line in result.log)

    role_registry = json.loads(
        (run.template_profile.bundle_dir / "arch_style_registry.json").read_text(
            encoding="utf-8"
        )
    )
    assert role_registry["roles"]["PARAGRAPH"]["numbering_provenance"] == "text_literal"

    validate_docx_package(result.output_path)
    with zipfile.ZipFile(result.output_path) as package:
        output_document = package.read("word/document.xml")
        output_document_text = output_document.decode("utf-8")
        output_numbering = package.read("word/numbering.xml").decode("utf-8")

    assert _xml_text_sequence(output_document) == _xml_text_sequence(
        target_document_before
    )
    first_paragraph = output_document_text.split("</w:p>", 1)[0]
    assert re.search(
        r'<w:pStyle w:val="SF_[^"]+_BODY_CSI_Paragraph__ARCH_[^"]+"/>',
        first_paragraph,
    )
    assert '<w:numId w:val="17"/>' in first_paragraph
    assert '<w:ilvl w:val="0"/>' in first_paragraph
    assert '<w:num w:numId="17">' in output_numbering


def _write_canadian_pair(tmp_path: Path) -> tuple[Path, Path]:
    """An architect with a two-level Canadian list and a typed-CSI target."""

    architect = tmp_path / "canadian-architect.docx"
    target = tmp_path / "csi-target.docx"
    _write_docx(architect, architect=True)
    _write_docx(target, architect=False)

    with zipfile.ZipFile(architect) as package:
        architect_document = package.read("word/document.xml").decode("utf-8")
        architect_numbering = package.read("word/numbering.xml").decode("utf-8")
    architect_document = architect_document.replace(
        '<w:ilvl w:val="2"/><w:numId w:val="5"/>',
        '<w:ilvl w:val="1"/><w:numId w:val="5"/>',
        1,
    )
    architect_numbering = architect_numbering.replace(
        '<w:numFmt w:val="upperLetter"/><w:lvlText w:val="%1."/>',
        '<w:numFmt w:val="decimal"/><w:lvlText w:val=".%1"/>',
        1,
    ).replace(
        '<w:lvl w:ilvl="2">',
        '<w:lvl w:ilvl="1">',
        1,
    ).replace(
        '<w:lvlText w:val="%3)"/>',
        '<w:lvlText w:val=".%2"/>',
        1,
    )
    _rewrite_docx_parts(
        architect,
        {
            "word/document.xml": architect_document,
            "word/numbering.xml": architect_numbering,
        },
    )

    with zipfile.ZipFile(target) as package:
        target_document = package.read("word/document.xml").decode("utf-8")
    target_document = re.sub(
        r'(<w:body><w:p>)<w:pPr><w:numPr>[\s\S]*?</w:numPr></w:pPr>',
        r"\1",
        target_document,
        count=1,
    )
    target_document = target_document.replace(
        '<w:pPr><w:pStyle w:val="TargetLevel2"/></w:pPr>',
        "",
        1,
    ).replace(
        "Architect paragraph one",
        "A. Work Included",
        1,
    ).replace(
        "Architect paragraph two",
        "1. Pumps",
        1,
    )
    _rewrite_docx_parts(target, {"word/document.xml": target_document})
    return architect, target


def test_unified_canadian_mode_converts_typed_csi_markers_end_to_end(
    tmp_path: Path,
) -> None:
    architect, target = _write_canadian_pair(tmp_path)
    target_sha = _sha256(target)

    run = format_specifications(
        architect_template=architect,
        target_specs=[target],
        output_dir=tmp_path / "formatted",
        cache_dir=tmp_path / "template-cache",
        api_key="",
        max_workers=1,
        conversion_mode=CSI_TO_CANADIAN,
        template_model="canadian-roundtrip-fixture",
        template_classifier=_deterministic_classifier,
    )

    assert run.success, "\n".join(run.targets[0].log)
    result = run.targets[0]
    assert result.output_path is not None
    assert result.output_path.name == "csi-target_CANADIAN_FORMATTED.docx"
    assert result.conversion_report is not None
    assert result.conversion_report.literal_markers_removed == 2
    assert _sha256(target) == target_sha

    validate_docx_package(result.output_path)
    with zipfile.ZipFile(result.output_path) as package:
        output_document = package.read("word/document.xml")
        output_numbering = package.read("word/numbering.xml").decode("utf-8")
    assert _xml_text_sequence(output_document)[:2] == ["Work Included", "Pumps"]
    assert '<w:lvlText w:val=".%1"/>' in output_numbering
    assert '<w:lvlText w:val=".%2"/>' in output_numbering
    assert b'<w:numId w:val="17"/>' not in output_document


def _extension_namespace_architect_styles() -> str:
    """An architect stylesheet shaped the way current Word writes one.

    ``w14`` is declared on the root and marked ignorable, the document
    defaults carry ligatures, and ``Normal`` overrides them, so a clone that
    resolves its inheritance correctly carries Normal's value, not the
    defaults'.
    """

    return (
        f'<w:styles xmlns:w="{W_NS}" xmlns:w14="{W14_NS}" xmlns:mc="{MC_NS}" '
        'mc:Ignorable="w14">'
        "<w:docDefaults><w:rPrDefault><w:rPr>"
        '<w:rFonts w:ascii="Aptos" w:hAnsi="Aptos"/><w:sz w:val="24"/>'
        '<w14:ligatures w14:val="standardContextual"/>'
        "</w:rPr></w:rPrDefault>"
        "<w:pPrDefault><w:pPr/></w:pPrDefault></w:docDefaults>"
        '<w:style w:type="paragraph" w:default="1" w:styleId="Normal">'
        '<w:name w:val="Normal"/><w:qFormat/>'
        '<w:rPr><w14:ligatures w14:val="standard"/></w:rPr></w:style>'
        "</w:styles>"
    )


def _styles_root_tag(styles_xml: str) -> str:
    match = re.search(r"<w:styles\b[^>]*>", styles_xml)
    assert match is not None, "output stylesheet has no w:styles root"
    return match.group(0)


def _assert_declares_ignorable_w14(styles_xml: str) -> ET.Element:
    root_tag = _styles_root_tag(styles_xml)
    assert f'xmlns:w14="{W14_NS}"' in root_tag
    assert f'xmlns:mc="{MC_NS}"' in root_tag
    assert 'mc:Ignorable="w14"' in root_tag
    root = ET.fromstring(styles_xml.encode("utf-8"))
    defaults = root.find(
        f"{{{W_NS}}}docDefaults/{{{W_NS}}}rPrDefault/{{{W_NS}}}rPr/{{{W14_NS}}}ligatures"
    )
    assert defaults is not None
    assert defaults.get(f"{{{W14_NS}}}val") == "standardContextual"
    return root


def _assert_ligatures_last(style: ET.Element, value: str) -> None:
    rpr = style.find(f"{{{W_NS}}}rPr")
    assert rpr is not None
    tags = [child.tag for child in rpr]
    assert tags[-1] == f"{{{W14_NS}}}ligatures", tags
    assert all(tag.startswith(f"{{{W_NS}}}") for tag in tags[:-1]), tags
    assert rpr[-1].get(f"{{{W14_NS}}}val") == value


def test_format_only_carries_extension_namespace_styles_into_a_bare_target_stylesheet(
    tmp_path: Path,
) -> None:
    # Every template current Word saves has w14:ligatures in its document
    # defaults. Format-only used to fail on it at style import ("unbound
    # prefix"), because materialization parsed each run-property block with
    # only the w namespace declared.
    architect = tmp_path / "w14-architect.docx"
    target = tmp_path / "bare-target.docx"
    _write_docx(architect, architect=True)
    _write_docx(target, architect=False)
    _rewrite_docx_parts(architect, {"word/styles.xml": _extension_namespace_architect_styles()})
    with zipfile.ZipFile(target) as package:
        target_styles = package.read("word/styles.xml").decode("utf-8")
        target_document_before = package.read("word/document.xml")
    assert _styles_root_tag(target_styles) == f'<w:styles xmlns:w="{W_NS}">'

    run = format_specifications(
        architect_template=architect,
        target_specs=[target],
        output_dir=tmp_path / "formatted",
        cache_dir=tmp_path / "template-cache",
        api_key="",
        max_workers=1,
        template_model="w14-format-only-fixture",
        template_classifier=_deterministic_classifier,
    )

    assert run.success, "\n".join(run.targets[0].log)
    result = run.targets[0]
    assert result.output_path is not None
    validate_docx_package(result.output_path)
    with zipfile.ZipFile(result.output_path) as package:
        output_styles = package.read("word/styles.xml").decode("utf-8")
        output_document = package.read("word/document.xml")

    root = _assert_declares_ignorable_w14(output_styles)
    # The defaults arrive first and declare w14; the imported styles then find
    # it declared. Both steps say what they did, in the log and as counts.
    run_log = (run.run_dir / "run.log").read_text(encoding="utf-8")
    assert (
        "Added XML namespace declarations to the target styles.xml root "
        "for the architect docDefaults: mc, w14, mc:Ignorable=w14"
    ) in run_log
    events = {
        event["event"]: event["fields"]
        for event in map(json.loads, run.diagnostics_path.read_text(encoding="utf-8").splitlines())
        if event.get("event") in {"apply_environment", "style_import"}
    }
    assert events["apply_environment"]["styles_namespace_additions"] == 3
    assert events["style_import"]["styles_namespace_additions"] == 0
    body_clones = [
        style
        for style in root.findall(f"{{{W_NS}}}style")
        if "_BODY_CSI_" in (style.get(f"{{{W_NS}}}styleId") or "")
    ]
    assert len(body_clones) == 2
    for clone in body_clones:
        # Detached from Normal, so Normal's override is materialized into it.
        _assert_ligatures_last(clone, "standard")
    assert _xml_text_sequence(output_document) == _xml_text_sequence(target_document_before)


def test_canadian_mode_carries_extension_namespace_styles_into_a_bare_target_stylesheet(
    tmp_path: Path,
) -> None:
    # The Canadian path does not materialize run properties, so it never hit
    # the parse error. It cloned Normal with its w14 child intact and wrote it
    # into a stylesheet whose root did not declare w14.
    architect, target = _write_canadian_pair(tmp_path)
    _rewrite_docx_parts(architect, {"word/styles.xml": _extension_namespace_architect_styles()})

    run = format_specifications(
        architect_template=architect,
        target_specs=[target],
        output_dir=tmp_path / "formatted",
        cache_dir=tmp_path / "template-cache",
        api_key="",
        max_workers=1,
        conversion_mode=CSI_TO_CANADIAN,
        template_model="w14-canadian-fixture",
        template_classifier=_deterministic_classifier,
    )

    assert run.success, "\n".join(run.targets[0].log)
    result = run.targets[0]
    assert result.output_path is not None
    validate_docx_package(result.output_path)
    with zipfile.ZipFile(result.output_path) as package:
        output_styles = package.read("word/styles.xml").decode("utf-8")

    root = _assert_declares_ignorable_w14(output_styles)
    normal_clones = [
        style
        for style in root.findall(f"{{{W_NS}}}style")
        if "_SHELL_Normal_" in (style.get(f"{{{W_NS}}}styleId") or "")
    ]
    assert len(normal_clones) == 1
    _assert_ligatures_last(normal_clones[0], "standard")
