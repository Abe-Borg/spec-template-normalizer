from __future__ import annotations

import hashlib
import json
import zipfile
from pathlib import Path

import pytest

from phase1_bundle import (
    MANIFEST_FILENAME,
    BundleArtifacts,
    ProducerIdentity,
    build_classification_audit,
    bundle_directory_name,
    discard_staged_bundle,
    publish_phase1_bundle,
    publish_staged_bundle,
    source_identity,
    stage_phase1_bundle,
    validate_classification_audit,
    validate_bundle_directory,
)


W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
SOURCE_STYLES = (
    f'<w:styles xmlns:w="{W_NS}">'
    '<w:style w:type="paragraph" w:styleId="Normal"><w:name w:val="Normal"/></w:style>'
    "</w:styles>"
).encode()
PORTABLE_STYLES = (
    f'<w:styles xmlns:w="{W_NS}">'
    '<w:style w:type="paragraph" w:styleId="Normal"><w:name w:val="Normal"/></w:style>'
    '<w:style w:type="paragraph" w:styleId="CSI_EndOfSection__ARCH">'
    '<w:name w:val="CSI End Of Section"/></w:style>'
    "</w:styles>"
).encode()
SOURCE_SETTINGS = f'<w:settings xmlns:w="{W_NS}"><w:zoom w:percent="100"/></w:settings>'.encode()


def _sha(path: Path) -> str:
    return hashlib.sha256(path.read_bytes()).hexdigest()


def _template_registry(source: Path) -> dict:
    normal_xml = f'<w:style xmlns:w="{W_NS}" w:type="paragraph" w:styleId="Normal"/>'
    with zipfile.ZipFile(source) as package:
        has_settings = "word/settings.xml" in package.namelist()
    return {
        "meta": {
            "schema_version": "1.0.0",
            "source_docx": {
                "filename": source.name,
                "sha256": _sha(source),
                "extracted_utc": "2026-07-17T00:00:00Z",
            },
        },
        "package_inventory": {
            "has_styles": True,
            "has_theme": False,
            "has_settings": has_settings,
        },
        "doc_defaults": {
            "default_run_props": {"rPr": None},
            "default_paragraph_props": {"pPr": None},
        },
        "styles": {
            "style_defs": [
                {"style_id": "Normal", "raw_style_xml": normal_xml},
            ],
            "latent_styles": {"latentStyles_xml": None},
        },
        "theme": {"theme1_xml": None},
        "settings": {"compat": {"compat_xml": None}},
        "page_layout": {"default_section": {"sectPr": None}, "section_chain": []},
        "headers_footers": {"headers": [], "footers": [], "header_footer_media": []},
        "numbering": {"abstract_nums": [], "nums": []},
        "fonts": {"font_table_xml": None},
        "custom_xml": {"relationships": [], "other_parts_passthrough": []},
        "capture_policy": {
            "store_raw_xml_blocks": False,
            "store_normalized_xml_blocks": True,
        },
    }


def _build_inputs(tmp_path: Path, *, with_settings: bool = True) -> tuple[Path, BundleArtifacts]:
    source = tmp_path / "Architect Template.docx"
    with zipfile.ZipFile(source, "w") as package:
        package.writestr("word/styles.xml", SOURCE_STYLES)
        if with_settings:
            package.writestr("word/settings.xml", SOURCE_SETTINGS)

    source_styles = tmp_path / "captured-source-styles.xml"
    source_styles.write_bytes(SOURCE_STYLES)
    portable_styles = tmp_path / "generated-portable-styles.xml"
    portable_styles.write_bytes(PORTABLE_STYLES)
    source_settings = None
    if with_settings:
        source_settings = tmp_path / "captured-source-settings.xml"
        source_settings.write_bytes(SOURCE_SETTINGS)

    style_registry = tmp_path / "style.json"
    style_registry.write_text(
        json.dumps(
            {
                "version": 2,
                "source_docx": source.name,
                "source_sha256": _sha(source),
                "source_tokens": {},
                "roles": {
                    "END_OF_SECTION": {
                        "style_id": "CSI_EndOfSection__ARCH",
                        "exemplar_paragraph_index": 0,
                        "numbering_provenance": "none",
                    }
                },
            }
        ),
        encoding="utf-8",
    )
    template_registry = tmp_path / "template.json"
    template_registry.write_text(json.dumps(_template_registry(source)), encoding="utf-8")
    audit_path = tmp_path / "classification.json"
    audit_instructions = {
        "create_styles": [
            {
                "styleId": "CSI_EndOfSection__ARCH",
                "derive_from_paragraph_index": 0,
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
    }
    audit = build_classification_audit(
        audit_instructions,
        {
            "paragraphs": [
                {
                    "paragraph_index": 0,
                    "text": "END OF SECTION",
                    "skip_reason": None,
                    "text_was_truncated": False,
                }
            ]
        },
        _sha(source),
        [0],
    )
    audit_path.write_text(json.dumps(audit), encoding="utf-8")
    return source, BundleArtifacts(
        style_registry=style_registry,
        template_registry=template_registry,
        classification_audit=audit_path,
        source_styles=source_styles,
        portable_styles=portable_styles,
        source_settings=source_settings,
    )


def _producer(run_id: str = "run-001") -> ProducerIdentity:
    return ProducerIdentity(
        name="spec-template-normalizer",
        version="2.0.0",
        run_id=run_id,
        classifier_provider="anthropic",
        classifier_model="test-model",
        master_prompt_sha256="1" * 64,
        run_instruction_sha256="2" * 64,
    )


def test_stage_publish_and_validate_complete_bundle(tmp_path: Path):
    source, artifacts = _build_inputs(tmp_path)
    output = tmp_path / "out"

    staged = stage_phase1_bundle(output, source, artifacts, _producer())
    assert staged.path.name.startswith(".phase1-bundle-staging-")
    assert not (output / staged.directory_name).exists()

    published = publish_staged_bundle(staged)
    manifest = validate_bundle_directory(published, expected_source_sha256=_sha(source))

    assert published.name.endswith(".phase1")
    assert manifest.source.filename == source.name
    assert {item.artifact_id for item in manifest.artifacts} == {
        "style_registry",
        "template_registry",
        "classification_audit",
        "source_styles",
        "portable_styles",
        "source_settings",
    }
    assert (published / "source_styles.xml").read_bytes() == SOURCE_STYLES
    assert (published / "portable_styles.xml").read_bytes() == PORTABLE_STYLES


def test_source_settings_is_omitted_only_when_source_has_no_settings_part(tmp_path: Path):
    source, artifacts = _build_inputs(tmp_path, with_settings=False)
    published = publish_phase1_bundle(tmp_path / "out", source, artifacts, _producer())
    manifest = validate_bundle_directory(published)
    assert "source_settings" not in {item.artifact_id for item in manifest.artifacts}
    assert not (published / "source_settings.xml").exists()


def test_source_settings_is_required_when_present_in_docx(tmp_path: Path):
    source, artifacts = _build_inputs(tmp_path)
    without_settings = BundleArtifacts(
        style_registry=artifacts.style_registry,
        template_registry=artifacts.template_registry,
        classification_audit=artifacts.classification_audit,
        source_styles=artifacts.source_styles,
        portable_styles=artifacts.portable_styles,
    )
    with pytest.raises(ValueError, match="source_settings must be supplied"):
        stage_phase1_bundle(tmp_path / "out", source, without_settings, _producer())


def test_rejects_source_styles_that_are_not_exact_original_bytes(tmp_path: Path):
    source, artifacts = _build_inputs(tmp_path)
    artifacts.source_styles.write_bytes(PORTABLE_STYLES)
    with pytest.raises(ValueError, match="not the exact pre-mutation"):
        stage_phase1_bundle(tmp_path / "out", source, artifacts, _producer())


def test_rejects_stale_registry_source_identity(tmp_path: Path):
    source, artifacts = _build_inputs(tmp_path)
    registry = json.loads(artifacts.style_registry.read_text(encoding="utf-8"))
    registry["source_sha256"] = "0" * 64
    artifacts.style_registry.write_text(json.dumps(registry), encoding="utf-8")
    with pytest.raises(ValueError, match="source_sha256 does not match"):
        stage_phase1_bundle(tmp_path / "out", source, artifacts, _producer())


def test_rejects_undeclared_extra_portable_style(tmp_path: Path):
    source, artifacts = _build_inputs(tmp_path)
    portable = artifacts.portable_styles.read_text(encoding="utf-8")
    artifacts.portable_styles.write_text(
        portable.replace(
            "</w:styles>",
            '<w:style w:type="paragraph" w:styleId="UndeclaredExtra"/>'
            "</w:styles>",
        ),
        encoding="utf-8",
    )
    with pytest.raises(ValueError, match="neither source styles nor declared role styles"):
        stage_phase1_bundle(tmp_path / "out", source, artifacts, _producer())


def test_classification_audit_preserves_exact_dispositions_and_fingerprints():
    instructions = {
        "apply_pStyle": [{"paragraph_index": 0, "styleId": "Normal"}],
        "ignored_paragraphs": [{"paragraph_index": 1, "reason": "NON_CSI"}],
        "roles": {},
    }
    audit = build_classification_audit(
        instructions,
        {
            "paragraphs": [
                {"paragraph_index": 0, "text": "A. Scope", "skip_reason": None},
                {"paragraph_index": 1, "text": "[NOTE]", "skip_reason": "editor_note"},
                {"paragraph_index": 2, "text": "", "skip_reason": "blank"},
            ]
        },
        "a" * 64,
        [0, 1],
    )
    validated = validate_classification_audit(audit, expected_source_sha256="a" * 64)
    assert [item["classification"] for item in validated["paragraphs"]] == [
        "styled",
        "ignored",
        "out_of_scope",
    ]
    assert validated["paragraphs"][1]["ignore_reason"] == "NON_CSI"
    assert validated["paragraphs"][1]["text"] == "[NOTE]"


def test_classification_audit_rejects_changed_paragraph_text():
    audit = build_classification_audit(
        {"apply_pStyle": [], "ignored_paragraphs": [], "roles": {}},
        {"paragraphs": [{"paragraph_index": 0, "text": "original", "skip_reason": "blank"}]},
        "a" * 64,
        [],
    )
    audit["paragraphs"][0]["text"] = "changed"
    with pytest.raises(ValueError, match="slim_text_sha256 mismatch"):
        validate_classification_audit(audit)


@pytest.mark.parametrize("mutation", ["tamper", "missing", "extra"])
def test_validation_rejects_tampered_missing_and_unlisted_artifacts(tmp_path: Path, mutation: str):
    source, artifacts = _build_inputs(tmp_path)
    published = publish_phase1_bundle(tmp_path / "out", source, artifacts, _producer())
    if mutation == "tamper":
        (published / "portable_styles.xml").write_bytes(b"changed")
        match = "SHA-256 mismatch|size mismatch"
    elif mutation == "missing":
        (published / "portable_styles.xml").unlink()
        match = "missing or not a regular file"
    else:
        (published / "arch_settings_raw.xml").write_bytes(b"stale")
        match = "unlisted/stale artifacts"
    with pytest.raises(ValueError, match=match):
        validate_bundle_directory(published)


def test_existing_destination_requires_explicit_overwrite(tmp_path: Path):
    source, artifacts = _build_inputs(tmp_path)
    output = tmp_path / "out"
    first = stage_phase1_bundle(
        output, source, artifacts, _producer("first"), directory_name="fixed.phase1"
    )
    publish_staged_bundle(first)

    replacement = stage_phase1_bundle(
        output, source, artifacts, _producer("second"), directory_name="fixed.phase1"
    )
    with pytest.raises(FileExistsError, match="overwrite=True"):
        publish_staged_bundle(replacement)
    published = publish_staged_bundle(replacement, overwrite=True)
    manifest = validate_bundle_directory(published)
    assert manifest.producer["run_id"] == "second"


def test_default_directory_name_is_source_scoped_and_run_unique(tmp_path: Path):
    source, artifacts = _build_inputs(tmp_path)
    identity = source_identity(source)
    assert bundle_directory_name(identity, "run-a") != bundle_directory_name(identity, "run-b")

    first = stage_phase1_bundle(tmp_path / "out", source, artifacts, _producer("run-a"))
    second = stage_phase1_bundle(tmp_path / "out", source, artifacts, _producer("run-b"))
    try:
        assert first.directory_name != second.directory_name
        assert identity.sha256[:12] in first.directory_name
    finally:
        discard_staged_bundle(first)
        discard_staged_bundle(second)


def test_manifest_schema_is_present_and_declares_v1_contract():
    schema_path = Path(__file__).parents[1] / "schemas" / "phase1_bundle_manifest.v1.schema.json"
    schema = json.loads(schema_path.read_text(encoding="utf-8"))
    assert schema["properties"]["bundle_format"]["const"] == "spec-template-normalizer.phase1"
    assert schema["properties"]["manifest_version"]["const"] == 1
    assert MANIFEST_FILENAME == "phase1_bundle_manifest.json"
