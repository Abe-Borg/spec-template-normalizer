from __future__ import annotations

import codecs
import json
import os
import stat
import zipfile
from pathlib import Path

import pytest

import docx_decomposer
import phase1_pipeline
from docx_decomposer import extract_docx
from phase1_bundle import validate_bundle_directory
from phase1_pipeline import run_phase1


W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
CONTENT_TYPES_NS = "http://schemas.openxmlformats.org/package/2006/content-types"

SOURCE_STYLES = (
    f'<w:styles xmlns:w="{W_NS}">'
    '<w:docDefaults><w:rPrDefault><w:rPr/></w:rPrDefault>'
    '<w:pPrDefault><w:pPr/></w:pPrDefault></w:docDefaults>'
    '<w:style w:type="paragraph" w:default="1" w:styleId="Normal">'
    '<w:name w:val="Normal"/><w:qFormat/>'
    "</w:style>"
    "</w:styles>"
).encode("utf-8")


def _encode_declared_xml(body: str, encoding_case: str) -> bytes:
    declarations = {
        "utf8-bom": "UTF-8",
        "utf16-le": "UTF-16",
        "utf16-be": "UTF-16",
        "latin1": "iso-8859-1",
    }
    text = f'<?xml version="1.0" encoding="{declarations[encoding_case]}"?>{body}'
    if encoding_case == "utf8-bom":
        return codecs.BOM_UTF8 + text.encode("utf-8")
    if encoding_case == "utf16-le":
        return codecs.BOM_UTF16_LE + text.encode("utf-16-le")
    if encoding_case == "utf16-be":
        return codecs.BOM_UTF16_BE + text.encode("utf-16-be")
    return text.encode("iso-8859-1")


def _write_minimal_docx(
    path: Path,
    *,
    styles_bytes: bytes = SOURCE_STYLES,
    settings_bytes: bytes | None = None,
) -> None:
    content_types = (
        f'<Types xmlns="{CONTENT_TYPES_NS}">'
        '<Default Extension="rels" '
        'ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
        '<Default Extension="xml" ContentType="application/xml"/>'
        '<Override PartName="/word/document.xml" '
        'ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
        '<Override PartName="/word/styles.xml" '
        'ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>'
        "</Types>"
    )
    root_rels = (
        f'<Relationships xmlns="{REL_NS}">'
        '<Relationship Id="rId1" '
        'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" '
        'Target="word/document.xml"/>'
        "</Relationships>"
    )
    document = (
        f'<w:document xmlns:w="{W_NS}"><w:body>'
        '<w:p><w:pPr><w:pStyle w:val="Normal"/></w:pPr>'
        '<w:r><w:t>END OF SECTION</w:t></w:r></w:p>'
        "<w:sectPr/>"
        "</w:body></w:document>"
    )
    with zipfile.ZipFile(path, "w", compression=zipfile.ZIP_DEFLATED) as package:
        package.writestr("[Content_Types].xml", content_types)
        package.writestr("_rels/.rels", root_rels)
        package.writestr("word/document.xml", document)
        package.writestr("word/styles.xml", styles_bytes)
        if settings_bytes is not None:
            package.writestr("word/settings.xml", settings_bytes)


def _write_prompts(prompt_dir: Path) -> None:
    prompt_dir.mkdir()
    (prompt_dir / "master_prompt.txt").write_text("deterministic master", encoding="utf-8")
    (prompt_dir / "run_instruction_prompt.txt").write_text(
        "deterministic run instruction",
        encoding="utf-8",
    )


def _end_of_section_instructions() -> dict:
    return {
        "create_styles": [
            {
                "styleId": "CSI_EndOfSection__ARCH",
                "name": "CSI End Of Section",
                "type": "paragraph",
                "derive_from_paragraph_index": 0,
                "basedOn": "Normal",
                "role": "END_OF_SECTION",
            }
        ],
        "apply_pStyle": [
            {"paragraph_index": 0, "styleId": "CSI_EndOfSection__ARCH"},
        ],
        "ignored_paragraphs": [],
        "roles": {
            "END_OF_SECTION": {
                "styleId": "CSI_EndOfSection__ARCH",
                "exemplar_paragraph_index": 0,
            }
        },
        "notes": [],
    }


def _end_of_section_wire_instructions() -> dict:
    instructions = _end_of_section_instructions()
    instructions["roles"] = [
        {
            "role": "END_OF_SECTION",
            "styleId": "CSI_EndOfSection__ARCH",
            "exemplar_paragraph_index": 0,
        }
    ]
    return instructions


def test_injected_classifier_runs_pipeline_and_publishes_valid_bundle(tmp_path: Path) -> None:
    source = tmp_path / "architect.docx"
    output_root = tmp_path / "output"
    prompt_dir = tmp_path / "prompts"
    _write_minimal_docx(source)
    _write_prompts(prompt_dir)
    original_docx = source.read_bytes()
    classifier_calls = []

    def deterministic_classifier(**kwargs):
        classifier_calls.append(kwargs)
        return _end_of_section_wire_instructions()

    result = run_phase1(
        source,
        output_root,
        api_key="",
        model="deterministic-test-model",
        prompt_dir=prompt_dir,
        classifier=deterministic_classifier,
    )
    manifest = validate_bundle_directory(
        result.bundle_dir,
        expected_source_sha256=result.source_sha256,
    )

    assert len(classifier_calls) == 1
    assert classifier_calls[0]["api_key"] == ""
    assert classifier_calls[0]["model"] == "deterministic-test-model"
    assert classifier_calls[0]["slim_bundle"]["paragraphs"][0]["text"] == "END OF SECTION"
    assert result.coverage == 1.0
    assert result.manifest_path.is_file()
    assert manifest.producer["classifier"] == {
        "provider": "injected",
        "model": "deterministic_classifier",
    }
    assert (result.bundle_dir / "source_styles.xml").read_bytes() == SOURCE_STYLES
    assert source.read_bytes() == original_docx

    # Only the immutable published bundle may survive. The snapshot, extracted
    # package, staging tree, and generated intermediates remain private.
    assert list(output_root.iterdir()) == [result.bundle_dir]
    assert not any(path.name == "extracted" for path in output_root.rglob("*"))
    assert not any(path.name.startswith(".phase1-work-") for path in output_root.rglob("*"))
    assert not any(path.name.startswith(".phase1-bundle-staging-") for path in output_root.rglob("*"))


@pytest.mark.parametrize(
    "encoding_case",
    ["utf8-bom", "utf16-le", "utf16-be", "latin1"],
)
def test_registry_xml_is_utf8_normalized_while_source_artifacts_remain_exact(
    tmp_path: Path,
    encoding_case: str,
) -> None:
    source = tmp_path / f"architect-{encoding_case}.docx"
    output_root = tmp_path / "output"
    prompt_dir = tmp_path / "prompts"
    styles_bytes = _encode_declared_xml(SOURCE_STYLES.decode("utf-8"), encoding_case)
    settings_body = (
        f'<w:settings xmlns:w="{W_NS}"><w:docVar w:name="caf\u00e9" w:val="1"/>'
        '<w:zoom w:percent="100"/></w:settings>'
    )
    settings_bytes = _encode_declared_xml(settings_body, encoding_case)
    _write_minimal_docx(
        source,
        styles_bytes=styles_bytes,
        settings_bytes=settings_bytes,
    )
    _write_prompts(prompt_dir)

    result = run_phase1(
        source,
        output_root,
        api_key="",
        model="deterministic-test-model",
        prompt_dir=prompt_dir,
        classifier=lambda **_kwargs: _end_of_section_instructions(),
    )

    registry = json.loads(
        (result.bundle_dir / "arch_template_registry.json").read_text(encoding="utf-8")
    )
    assert 'encoding="UTF-8"' in registry["settings"]["settings_xml"]
    assert "caf\u00e9" in registry["settings"]["settings_xml"]
    assert (result.bundle_dir / "source_styles.xml").read_bytes() == styles_bytes
    assert (result.bundle_dir / "source_settings.xml").read_bytes() == settings_bytes


@pytest.mark.parametrize(
    ("malicious_name", "message"),
    [
        ("../escaped.txt", "traversal"),
        ("word\\escaped.xml", "Unsafe DOCX package member name"),
        ("C:/escaped.xml", "absolute"),
    ],
)
def test_unsafe_zip_member_is_rejected_and_partial_extract_removed(
    tmp_path: Path,
    malicious_name: str,
    message: str,
) -> None:
    source = tmp_path / "malicious.docx"
    extract_dir = tmp_path / "extracted"
    archive_name = malicious_name.replace("\\", "/")
    with zipfile.ZipFile(source, "w") as package:
        package.writestr("word/partial.txt", b"partial")
        package.writestr(archive_name, b"secret")
    if "\\" in malicious_name:
        # ZipInfo normalizes backslashes when writing on Windows. Patch both
        # the local and central directory names so the archive itself contains
        # the hostile raw name; lengths are identical, so offsets stay valid.
        archive_bytes = source.read_bytes()
        assert archive_bytes.count(archive_name.encode("ascii")) == 2
        source.write_bytes(
            archive_bytes.replace(
                archive_name.encode("ascii"),
                malicious_name.encode("ascii"),
            )
        )

    with pytest.raises(ValueError, match=message):
        extract_docx(source, extract_dir)

    assert not extract_dir.exists()
    assert not (tmp_path / "escaped.txt").exists()


def test_zip_symlink_is_rejected_and_partial_extract_removed(tmp_path: Path) -> None:
    source = tmp_path / "symlink.docx"
    extract_dir = tmp_path / "extracted"
    symlink = zipfile.ZipInfo("word/media/image-link.png")
    symlink.create_system = 3
    symlink.external_attr = (stat.S_IFLNK | 0o777) << 16
    with zipfile.ZipFile(source, "w") as package:
        package.writestr("word/partial.txt", b"partial")
        package.writestr(symlink, "../../outside-secret")

    with pytest.raises(ValueError, match="symbolic link"):
        extract_docx(source, extract_dir)

    assert not extract_dir.exists()


def test_zip_part_size_metadata_limit_removes_partial_extract(
    tmp_path: Path,
    monkeypatch,
) -> None:
    source = tmp_path / "oversized.docx"
    extract_dir = tmp_path / "extracted"
    with zipfile.ZipFile(source, "w") as package:
        package.writestr("word/partial.txt", b"ok")
        package.writestr("word/oversized.bin", b"12345")
    monkeypatch.setattr(docx_decomposer, "MAX_PACKAGE_PART_BYTES", 4)

    with pytest.raises(ValueError, match="per-part limit is 4"):
        extract_docx(source, extract_dir)

    assert not extract_dir.exists()


def test_zip_compression_ratio_limit_removes_partial_extract(
    tmp_path: Path,
    monkeypatch,
) -> None:
    source = tmp_path / "compression-bomb.docx"
    extract_dir = tmp_path / "extracted"
    with zipfile.ZipFile(source, "w", compression=zipfile.ZIP_DEFLATED) as package:
        package.writestr("word/partial.txt", os.urandom(64))
        package.writestr("word/compression-bomb.bin", b"0" * 4096)
    monkeypatch.setattr(docx_decomposer, "MAX_COMPRESSION_RATIO", 2)

    with pytest.raises(ValueError, match="suspicious compression ratio"):
        extract_docx(source, extract_dir)

    assert not extract_dir.exists()


def test_source_snapshot_rejects_change_during_copy(tmp_path: Path, monkeypatch) -> None:
    source = tmp_path / "source.docx"
    destination = tmp_path / "snapshot" / "source.docx"
    source.write_bytes(b"stable source bytes")
    original_copy = phase1_pipeline.shutil.copyfileobj

    def copy_then_change_timestamp(reader, writer, length):
        original_copy(reader, writer, length=length)
        current = source.stat()
        os.utime(
            source,
            ns=(current.st_atime_ns, current.st_mtime_ns + 10_000_000),
        )

    monkeypatch.setattr(
        phase1_pipeline.shutil,
        "copyfileobj",
        copy_then_change_timestamp,
    )

    with pytest.raises(RuntimeError, match="changed while it was being snapshotted"):
        phase1_pipeline._snapshot_source(source, destination)


def test_source_snapshot_rejects_same_size_change_with_restored_timestamp(
    tmp_path: Path, monkeypatch
) -> None:
    source = tmp_path / "source.docx"
    destination = tmp_path / "snapshot" / "source.docx"
    source.write_bytes(b"original-content")
    original_stat = source.stat()
    original_copy = phase1_pipeline.shutil.copyfileobj

    def copy_then_rewrite(reader, writer, length):
        original_copy(reader, writer, length=length)
        source.write_bytes(b"modified-content")
        os.utime(
            source,
            ns=(original_stat.st_atime_ns, original_stat.st_mtime_ns),
        )

    monkeypatch.setattr(phase1_pipeline.shutil, "copyfileobj", copy_then_rewrite)

    with pytest.raises(RuntimeError, match="changed while it was being snapshotted"):
        phase1_pipeline._snapshot_source(source, destination)


def test_progress_callback_failure_is_nonfatal() -> None:
    def broken_progress(_message: str) -> None:
        raise RuntimeError("closed UI")

    with pytest.warns(RuntimeWarning, match="progress callback failed"):
        phase1_pipeline._emit(broken_progress, "published")
