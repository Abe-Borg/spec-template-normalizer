from __future__ import annotations

import codecs
import hashlib
import json
import os
import re
import subprocess
import sys
import zipfile
from pathlib import Path
import xml.etree.ElementTree as ET

import pytest

from spec_formatter.style_application.batch_runner import load_and_validate_shared_config, process_single_file
from spec_formatter.style_application.phase2_invariants import validate_docx_package


W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"
REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
CT_NS = "http://schemas.openxmlformats.org/package/2006/content-types"

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


def _normalizer_repo() -> Path:
    candidates = []
    configured = os.environ.get("SPEC_TEMPLATE_NORMALIZER_REPO")
    if configured:
        candidates.append(Path(configured))
    candidates.extend(
        [
            Path.cwd().parent / "spec-template-normalizer",
            Path(__file__).resolve().parents[2] / "spec-template-normalizer",
            Path(r"C:\Github-Repos\spec-template-normalizer"),
        ]
    )
    for candidate in candidates:
        if (candidate / "phase1_pipeline.py").is_file():
            return candidate.resolve()
    pytest.skip(
        "Cross-repository fixture requires spec-template-normalizer; set "
        "SPEC_TEMPLATE_NORMALIZER_REPO to its checkout"
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
    first_section_break = (
        f'<w:p><w:pPr>{_section_xml(architect=architect, first=True)}</w:pPr></w:p>'
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
        f'{first_section_break}'
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


def _run_phase1(architect: Path, output_root: Path, normalizer: Path) -> Path:
    script = r'''
import json
import sys
from pathlib import Path

from phase1_bundle import validate_bundle_directory
from phase1_pipeline import run_phase1


def deterministic_classifier(**kwargs):
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


result = run_phase1(
    Path(sys.argv[1]),
    Path(sys.argv[2]),
    api_key="",
    model="cross-repository-fixture",
    classifier=deterministic_classifier,
)
validate_bundle_directory(result.bundle_dir, expected_source_sha256=result.source_sha256)
print("PHASE1_BUNDLE=" + str(result.bundle_dir))
'''
    env = os.environ.copy()
    search_paths = [str(normalizer)]
    normalizer_venv = normalizer / "venv" / "Lib" / "site-packages"
    if normalizer_venv.is_dir():
        search_paths.append(str(normalizer_venv))
    if env.get("PYTHONPATH"):
        search_paths.append(env["PYTHONPATH"])
    env["PYTHONPATH"] = os.pathsep.join(search_paths)
    completed = subprocess.run(
        [sys.executable, "-c", script, str(architect), str(output_root)],
        cwd=normalizer,
        env=env,
        capture_output=True,
        text=True,
        timeout=60,
        check=False,
    )
    if completed.returncode != 0:
        raise AssertionError(
            "Phase 1 subprocess failed\nSTDOUT:\n"
            f"{completed.stdout}\nSTDERR:\n{completed.stderr}"
        )
    marker = next(
        (line for line in completed.stdout.splitlines() if line.startswith("PHASE1_BUNDLE=")),
        None,
    )
    assert marker is not None, completed.stdout
    bundle = Path(marker.split("=", 1)[1])
    assert bundle.is_dir()
    return bundle


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


def test_generated_phase1_bundle_round_trips_through_phase2_without_api(tmp_path: Path) -> None:
    normalizer = _normalizer_repo()
    architect = tmp_path / "architect.docx"
    target = tmp_path / "target.docx"
    _write_docx(architect, architect=True)
    _write_docx(target, architect=False)

    architect_sha = _sha256(architect)
    target_sha = _sha256(target)
    with zipfile.ZipFile(target) as package:
        target_document_before = package.read("word/document.xml")

    bundle = _run_phase1(architect, tmp_path / "phase1-output", normalizer)
    assert _sha256(architect) == architect_sha

    manifest = json.loads((bundle / "phase1_bundle_manifest.json").read_text(encoding="utf-8"))
    registry = json.loads((bundle / "arch_template_registry.json").read_text(encoding="utf-8"))
    style_registry = json.loads((bundle / "arch_style_registry.json").read_text(encoding="utf-8"))
    assert manifest["source"]["sha256"] == architect_sha
    assert (bundle / "source_settings.xml").read_bytes().startswith(codecs.BOM_UTF16_LE)
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

    shared = load_and_validate_shared_config(bundle)
    result = process_single_file(
        docx_path=target,
        arch_registry=shared.arch_registry,
        env_registry=shared.env_registry,
        arch_styles_xml=shared.arch_styles_xml,
        available_roles=shared.available_roles,
        api_key="",
        output_dir=tmp_path / "phase2-output",
        source_tokens=shared.source_tokens,
        arch_root=shared.arch_root,
        role_specs=shared.role_specs,
    )
    assert result.success, "\n".join(result.log)
    assert result.output_path is not None and result.output_path.is_file()
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

    assert output_document_text.count('w:val="CSI_Paragraph__ARCH"') == 1
    assert output_document_text.count('w:val="CSI_Subparagraph__ARCH"') == 1
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
