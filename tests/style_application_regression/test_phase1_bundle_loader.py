import hashlib
import json
from pathlib import Path

import pytest

from spec_formatter.style_application.batch_runner import load_and_validate_shared_config


SOURCE_SHA = "a" * 64
W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"


def _sha(path: Path) -> str:
    return hashlib.sha256(path.read_bytes()).hexdigest()


def _canonical_sha(value) -> str:
    payload = json.dumps(
        value,
        ensure_ascii=False,
        sort_keys=True,
        separators=(",", ":"),
        allow_nan=False,
    ).encode("utf-8")
    return hashlib.sha256(payload).hexdigest()


def _registry_payloads():
    style_registry = {
        "version": 2,
        "source_docx": "architect.docx",
        "source_sha256": SOURCE_SHA,
        "source_tokens": {"SectionTitle": "AIR TERMINALS"},
        "roles": {"PART": {
            "style_id": "CSI-Part",
            "exemplar_paragraph_index": 0,
            "numbering_provenance": "none",
        }},
    }
    template_registry = {
        "meta": {
            "schema_version": "1.0.0",
            "source_docx": {
                "filename": "architect.docx",
                "sha256": SOURCE_SHA,
            }
        },
        "theme": {},
        "settings": {},
        "fonts": {},
        "package_inventory": {"has_settings": False},
        "headers_footers": {"headers": [], "footers": [], "header_footer_media": []},
        "custom_xml": {"relationships": [], "other_parts_passthrough": []},
        "capture_policy": {
            "store_raw_xml_blocks": False,
            "store_normalized_xml_blocks": True,
        },
        "doc_defaults": {},
        "styles": {
            "style_defs": [{
                "style_id": "CSI-Part",
                "type": "paragraph",
                "name": "CSI Part",
            }]
        },
        "numbering": {"abstract_nums": [], "nums": []},
        "page_layout": {
            "default_section": {
                "sectPr": f'<w:sectPr xmlns:w="{W_NS}"/>'
            },
            "section_chain": [],
        },
    }
    return style_registry, template_registry


def _write_bundle(root: Path, *, include_manifest=True) -> Path:
    root.mkdir()
    style_registry, template_registry = _registry_payloads()
    (root / "arch_style_registry.json").write_text(
        json.dumps(style_registry), encoding="utf-8"
    )
    (root / "arch_template_registry.json").write_text(
        json.dumps(template_registry), encoding="utf-8"
    )
    instructions = {
        "create_styles": [],
        "apply_pStyle": [{"paragraph_index": 0, "styleId": "CSI-Part"}],
        "ignored_paragraphs": [],
        "roles": {
            "PART": {"styleId": "CSI-Part", "exemplar_paragraph_index": 0}
        },
    }
    text = "PART 1 - GENERAL"
    paragraphs = [{
        "paragraph_index": 0,
        "text": text,
        "slim_text_sha256": hashlib.sha256(text.encode("utf-8")).hexdigest(),
        "text_was_truncated": False,
        "skip_reason": None,
        "classification": "styled",
        "style_id": "CSI-Part",
    }]
    audit = {
        "audit_version": 1,
        "source_sha256": SOURCE_SHA,
        "instructions_sha256": _canonical_sha(instructions),
        "paragraphs_sha256": _canonical_sha(paragraphs),
        "expected_paragraph_indices": [0],
        "instructions": instructions,
        "paragraphs": paragraphs,
    }
    (root / "classification_audit.json").write_text(json.dumps(audit), encoding="utf-8")
    styles_xml = (
        f'<w:styles xmlns:w="{W_NS}">'
        '<w:style w:type="paragraph" w:styleId="CSI-Part">'
        '<w:name w:val="CSI Part"/></w:style></w:styles>'
    )
    (root / "source_styles.xml").write_text(styles_xml, encoding="utf-8")
    (root / "portable_styles.xml").write_text(styles_xml, encoding="utf-8")

    if include_manifest:
        specs = [
            ("style_registry", "arch_style_registry.json", "application/json", True, "generated"),
            ("template_registry", "arch_template_registry.json", "application/json", True, "generated"),
            ("classification_audit", "classification_audit.json", "application/json", True, "generated"),
            ("source_styles", "source_styles.xml", "application/xml", True, "source_exact"),
            ("portable_styles", "portable_styles.xml", "application/xml", True, "generated"),
        ]
        artifacts = []
        for artifact_id, filename, media_type, required, source_kind in specs:
            path = root / filename
            artifacts.append({
                "artifact_id": artifact_id,
                "path": filename,
                "media_type": media_type,
                "sha256": _sha(path),
                "size_bytes": path.stat().st_size,
                "required": required,
                "source_kind": source_kind,
            })
        manifest = {
            "bundle_format": "spec-template-normalizer.phase1",
            "manifest_version": 1,
            "bundle_id": "aaaaaaaaaaaa-bbbbbbbbbbbb",
            "created_utc": "2026-07-17T12:00:00Z",
            "producer": {"name": "spec-template-normalizer", "version": "2", "run_id": "run-1"},
            "source": {
                "filename": "architect.docx",
                "sha256": SOURCE_SHA,
                "size_bytes": 12345,
            },
            "required_artifacts": [
                "style_registry", "template_registry", "classification_audit",
                "source_styles", "portable_styles"
            ],
            "artifacts": artifacts,
        }
        (root / "phase1_bundle_manifest.json").write_text(
            json.dumps(manifest), encoding="utf-8"
        )
    return root


def _refresh_manifest_records(bundle: Path, *filenames: str) -> None:
    manifest_path = bundle / "phase1_bundle_manifest.json"
    manifest = json.loads(manifest_path.read_text(encoding="utf-8"))
    selected = set(filenames)
    for record in manifest["artifacts"]:
        if record["path"] not in selected:
            continue
        path = bundle / record["path"]
        record["sha256"] = _sha(path)
        record["size_bytes"] = path.stat().st_size
    manifest_path.write_text(json.dumps(manifest), encoding="utf-8")


def _rewrite_audit_style(bundle: Path, style_id: str) -> None:
    audit_path = bundle / "classification_audit.json"
    audit = json.loads(audit_path.read_text(encoding="utf-8"))
    audit["instructions"]["create_styles"] = [{
        "styleId": style_id,
        "derive_from_paragraph_index": 0,
        "role": "PART",
    }]
    audit["instructions"]["roles"]["PART"]["styleId"] = style_id
    audit["instructions"]["apply_pStyle"][0]["styleId"] = style_id
    audit["paragraphs"][0]["style_id"] = style_id
    audit["instructions_sha256"] = _canonical_sha(audit["instructions"])
    audit["paragraphs_sha256"] = _canonical_sha(audit["paragraphs"])
    audit_path.write_text(json.dumps(audit), encoding="utf-8")


def test_strict_loader_verifies_bundle_and_uses_portable_styles(tmp_path):
    bundle = _write_bundle(tmp_path / "architect.phase1")
    shared = load_and_validate_shared_config(bundle)
    assert shared.arch_registry == {"PART": "CSI-Part"}
    assert 'w:styleId="CSI-Part"' in shared.arch_styles_xml
    assert shared.source_tokens == {"SectionTitle": "AIR TERMINALS"}
    assert shared.bundle_manifest["manifest_version"] == 1
    assert shared.legacy_mode is False


def test_strict_loader_accepts_declared_generated_portable_role_style(tmp_path):
    bundle = _write_bundle(tmp_path / "architect.phase1")
    style_registry_path = bundle / "arch_style_registry.json"
    style_registry = json.loads(style_registry_path.read_text(encoding="utf-8"))
    style_registry["roles"]["PART"]["style_id"] = "CSI_Part__ARCH"
    style_registry_path.write_text(json.dumps(style_registry), encoding="utf-8")
    _rewrite_audit_style(bundle, "CSI_Part__ARCH")

    portable_path = bundle / "portable_styles.xml"
    portable = portable_path.read_text(encoding="utf-8")
    portable_path.write_text(
        portable.replace(
            "</w:styles>",
            '<w:style w:type="paragraph" w:styleId="CSI_Part__ARCH">'
            '<w:name w:val="Generated Part"/></w:style></w:styles>',
        ),
        encoding="utf-8",
    )
    _refresh_manifest_records(
        bundle,
        "arch_style_registry.json",
        "classification_audit.json",
        "portable_styles.xml",
    )

    shared = load_and_validate_shared_config(bundle)
    assert shared.arch_registry == {"PART": "CSI_Part__ARCH"}
    assert 'w:styleId="CSI_Part__ARCH"' in shared.arch_styles_xml


def test_strict_loader_rejects_undeclared_extra_portable_style(tmp_path):
    bundle = _write_bundle(tmp_path / "architect.phase1")
    portable_path = bundle / "portable_styles.xml"
    portable = portable_path.read_text(encoding="utf-8")
    portable_path.write_text(
        portable.replace(
            "</w:styles>",
            '<w:style w:type="paragraph" w:styleId="Unexpected"/>'
            "</w:styles>",
        ),
        encoding="utf-8",
    )
    _refresh_manifest_records(bundle, "portable_styles.xml")

    with pytest.raises(ValueError, match="exactly the source styles plus declared role styles"):
        load_and_validate_shared_config(bundle)


def test_strict_loader_rejects_tampered_artifact(tmp_path):
    bundle = _write_bundle(tmp_path / "architect.phase1")
    path = bundle / "portable_styles.xml"
    content = path.read_text(encoding="utf-8")
    path.write_text(content.replace("CSI Part", "BAD PART"), encoding="utf-8")
    with pytest.raises(ValueError, match="SHA-256 mismatch"):
        load_and_validate_shared_config(bundle)


def test_strict_loader_checksums_classification_audit(tmp_path):
    bundle = _write_bundle(tmp_path / "architect.phase1")
    audit_path = bundle / "classification_audit.json"
    audit_path.write_text('{"audit_version": 9}', encoding="utf-8")
    with pytest.raises(ValueError, match="classification_audit.json"):
        load_and_validate_shared_config(bundle)


def test_strict_loader_rejects_rehashed_audit_from_different_source(tmp_path):
    bundle = _write_bundle(tmp_path / "architect.phase1")
    audit_path = bundle / "classification_audit.json"
    audit = json.loads(audit_path.read_text(encoding="utf-8"))
    audit["source_sha256"] = "b" * 64
    audit_path.write_text(json.dumps(audit), encoding="utf-8")
    _refresh_manifest_records(bundle, "classification_audit.json")

    with pytest.raises(ValueError, match="source_sha256 does not match"):
        load_and_validate_shared_config(bundle)


def test_strict_loader_rejects_rehashed_audit_role_mismatch(tmp_path):
    bundle = _write_bundle(tmp_path / "architect.phase1")
    audit_path = bundle / "classification_audit.json"
    audit = json.loads(audit_path.read_text(encoding="utf-8"))
    audit["instructions"]["roles"]["PART"]["exemplar_paragraph_index"] = 7
    audit["instructions_sha256"] = _canonical_sha(audit["instructions"])
    audit_path.write_text(json.dumps(audit), encoding="utf-8")
    _refresh_manifest_records(bundle, "classification_audit.json")

    with pytest.raises(ValueError, match="disagree for role PART"):
        load_and_validate_shared_config(bundle)


def test_strict_loader_rejects_boolean_audit_role_exemplar(tmp_path):
    bundle = _write_bundle(tmp_path / "architect.phase1")

    registry_path = bundle / "arch_style_registry.json"
    registry = json.loads(registry_path.read_text(encoding="utf-8"))
    registry["roles"]["PART"]["exemplar_paragraph_index"] = 1
    registry_path.write_text(json.dumps(registry), encoding="utf-8")

    audit_path = bundle / "classification_audit.json"
    audit = json.loads(audit_path.read_text(encoding="utf-8"))
    # True == 1 in Python, so equality-only linkage validation accepted this.
    audit["instructions"]["roles"]["PART"]["exemplar_paragraph_index"] = True
    audit["instructions_sha256"] = _canonical_sha(audit["instructions"])
    audit_path.write_text(json.dumps(audit), encoding="utf-8")
    _refresh_manifest_records(
        bundle,
        "arch_style_registry.json",
        "classification_audit.json",
    )

    with pytest.raises(ValueError, match="exemplar_paragraph_index must be a non-negative integer"):
        load_and_validate_shared_config(bundle)


def test_strict_loader_rejects_missing_required_template_section(tmp_path):
    bundle = _write_bundle(tmp_path / "architect.phase1")
    registry_path = bundle / "arch_template_registry.json"
    registry = json.loads(registry_path.read_text(encoding="utf-8"))
    del registry["headers_footers"]
    registry_path.write_text(json.dumps(registry), encoding="utf-8")
    _refresh_manifest_records(bundle, "arch_template_registry.json")

    with pytest.raises(ValueError, match="missing required sections"):
        load_and_validate_shared_config(bundle)


def test_strict_loader_rejects_settings_inventory_mismatch(tmp_path):
    bundle = _write_bundle(tmp_path / "architect.phase1")
    registry_path = bundle / "arch_template_registry.json"
    registry = json.loads(registry_path.read_text(encoding="utf-8"))
    registry["package_inventory"]["has_settings"] = True
    registry_path.write_text(json.dumps(registry), encoding="utf-8")
    _refresh_manifest_records(bundle, "arch_template_registry.json")

    with pytest.raises(ValueError, match="settings inventory disagrees"):
        load_and_validate_shared_config(bundle)


def test_strict_loader_rejects_unlisted_stale_file(tmp_path):
    bundle = _write_bundle(tmp_path / "architect.phase1")
    (bundle / "arch_settings_raw.xml").write_text("<legacy/>", encoding="utf-8")
    with pytest.raises(ValueError, match="unlisted/stale artifacts"):
        load_and_validate_shared_config(bundle)


def test_strict_loader_rejects_unsupported_manifest_version(tmp_path):
    bundle = _write_bundle(tmp_path / "architect.phase1")
    manifest_path = bundle / "phase1_bundle_manifest.json"
    manifest = json.loads(manifest_path.read_text(encoding="utf-8"))
    manifest["manifest_version"] = 2
    manifest_path.write_text(json.dumps(manifest), encoding="utf-8")
    with pytest.raises(ValueError, match="Unsupported manifest_version"):
        load_and_validate_shared_config(bundle)


def test_legacy_bundle_requires_explicit_opt_in(tmp_path):
    bundle = _write_bundle(tmp_path / "legacy", include_manifest=False)
    (bundle / "arch_styles_raw.xml").write_bytes((bundle / "portable_styles.xml").read_bytes())
    (bundle / "source_styles.xml").unlink()
    (bundle / "portable_styles.xml").unlink()

    with pytest.raises(FileNotFoundError, match="Strict Phase 1 bundle required"):
        load_and_validate_shared_config(bundle)

    shared = load_and_validate_shared_config(bundle, allow_legacy_bundle=True)
    assert shared.legacy_mode is True
    assert shared.bundle_manifest is None
    assert shared.role_specs is None
