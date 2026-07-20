"""
Registry loading and preflight reporting for Phase 2.

Handles loading the architect style registry and generating
preflight reports before classification application.
"""

import base64
import binascii
import re
import json
import hashlib
from datetime import datetime
import xml.etree.ElementTree as _ET
from pathlib import Path, PurePosixPath, PureWindowsPath
from typing import Dict, Any, List, Optional, Set, Tuple
from urllib.parse import unquote, urlsplit
from xml.sax.saxutils import escape as _sax_escape

from spec_formatter.role_contract import ALLOWED_ROLES

from .ooxml_namespaces import W_NS
from .ooxml_text import prepare_xml_text_for_utf8
from .opc_paths import (
    is_safe_header_footer_part_name,
    relationship_part_name_for_owner,
    resolve_internal_relationship_target as _resolve_safe_internal_relationship_target,
)


PHASE1_BUNDLE_FORMAT = "spec-template-normalizer.phase1"
PHASE1_MANIFEST_VERSION = 1
PHASE1_MANIFEST_FILENAME = "phase1_bundle_manifest.json"
PHASE1_REQUIRED_ARTIFACT_IDS = (
    "style_registry",
    "template_registry",
    "classification_audit",
    "source_styles",
    "portable_styles",
)
MAX_PHASE1_MANIFEST_BYTES = 2 * 1024 * 1024
MAX_PHASE1_ARTIFACT_BYTES = 256 * 1024 * 1024
MAX_CLASSIFICATION_AUDIT_RECORDS = 200_000
MAX_HEADER_FOOTER_MEDIA_BYTES = 16 * 1024 * 1024
MAX_HEADER_FOOTER_MEDIA_TOTAL_BYTES = 64 * 1024 * 1024
_PHASE1_ARTIFACT_SPECS = {
    "style_registry": (
        "arch_style_registry.json", "application/json", True, "generated"
    ),
    "template_registry": (
        "arch_template_registry.json", "application/json", True, "generated"
    ),
    "classification_audit": (
        "classification_audit.json", "application/json", True, "generated"
    ),
    "source_styles": (
        "source_styles.xml", "application/xml", True, "source_exact"
    ),
    "portable_styles": (
        "portable_styles.xml", "application/xml", True, "generated"
    ),
    "source_settings": (
        "source_settings.xml", "application/xml", False, "source_exact"
    ),
}
_SHA256_RE = re.compile(r"^[0-9a-f]{64}$")
_ALLOWED_ROLE_NAMES = set(ALLOWED_ROLES)

_PKG_REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
_EMBEDDED_FONT_TAGS = {
    f"{{{W_NS}}}embedRegular",
    f"{{{W_NS}}}embedBold",
    f"{{{W_NS}}}embedItalic",
    f"{{{W_NS}}}embedBoldItalic",
}
_MIME_TYPE_RE = re.compile(
    r"[A-Za-z0-9!#$&^_.+-]+/[A-Za-z0-9!#$&^_.+-]+\Z"
)


def _sha256_file(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as stream:
        for chunk in iter(lambda: stream.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def _canonical_json_sha256(value: Any) -> str:
    try:
        encoded = json.dumps(
            value,
            ensure_ascii=False,
            sort_keys=True,
            separators=(",", ":"),
            allow_nan=False,
        ).encode("utf-8")
    except (TypeError, ValueError) as exc:
        raise ValueError(f"classification audit contains non-canonical JSON data: {exc}") from exc
    return hashlib.sha256(encoded).hexdigest()


def _validate_classification_audit_linkage(
    audit: Dict[str, Any],
    source_sha256: str,
    style_roles: Dict[str, Dict[str, Any]],
) -> None:
    """Validate the audit artifact and bind it to the registries being consumed."""
    _require_exact_keys(
        audit,
        {
            "audit_version",
            "source_sha256",
            "instructions_sha256",
            "paragraphs_sha256",
            "expected_paragraph_indices",
            "instructions",
            "paragraphs",
        },
        "classification_audit.json",
    )
    if type(audit["audit_version"]) is not int or audit["audit_version"] != 1:
        raise ValueError("classification_audit.json has an unsupported audit_version")
    if audit["source_sha256"] != source_sha256:
        raise ValueError(
            "classification_audit.json source_sha256 does not match manifest source SHA-256"
        )
    for field in ("instructions_sha256", "paragraphs_sha256"):
        if not isinstance(audit[field], str) or not _SHA256_RE.fullmatch(audit[field]):
            raise ValueError(f"classification_audit.json {field} must be a SHA-256 digest")

    expected = audit["expected_paragraph_indices"]
    if (
        not isinstance(expected, list)
        or len(expected) > MAX_CLASSIFICATION_AUDIT_RECORDS
        or any(type(index) is not int or index < 0 for index in expected)
        or expected != sorted(expected)
        or len(expected) != len(set(expected))
    ):
        raise ValueError(
            "classification_audit.json expected_paragraph_indices must be unique sorted non-negative integers"
        )
    expected_set = set(expected)

    instructions = audit["instructions"]
    if not isinstance(instructions, dict):
        raise ValueError("classification_audit.json instructions must be an object")
    if set(instructions) - {
        "create_styles", "apply_pStyle", "ignored_paragraphs", "roles", "notes"
    }:
        raise ValueError("classification_audit.json instructions contains unknown keys")
    if _canonical_json_sha256(instructions) != audit["instructions_sha256"]:
        raise ValueError("classification_audit.json instructions_sha256 mismatch")

    audit_roles = instructions.get("roles")
    if not isinstance(audit_roles, dict) or set(audit_roles) != set(style_roles):
        raise ValueError(
            "classification_audit.json and arch_style_registry.json role sets differ"
        )
    for role, registry_spec in style_roles.items():
        audit_spec = audit_roles.get(role)
        if not isinstance(audit_spec, dict):
            raise ValueError(f"classification_audit.json roles[{role!r}] must be an object")
        if set(audit_spec) - {"styleId", "exemplar_paragraph_index", "style_name"}:
            raise ValueError(f"classification_audit.json roles[{role!r}] has unknown fields")
        audit_style_id = audit_spec.get("styleId")
        audit_exemplar = audit_spec.get("exemplar_paragraph_index")
        if not isinstance(audit_style_id, str):
            raise ValueError(
                f"classification_audit.json roles[{role!r}].styleId must be a string"
            )
        if type(audit_exemplar) is not int or audit_exemplar < 0:
            raise ValueError(
                "classification_audit.json "
                f"roles[{role!r}].exemplar_paragraph_index must be a non-negative integer"
            )
        if (
            audit_style_id != registry_spec.get("style_id")
            or audit_exemplar != registry_spec.get("exemplar_paragraph_index")
        ):
            raise ValueError(
                f"classification_audit.json and arch_style_registry.json disagree for role {role}"
            )

    applied: Dict[int, str] = {}
    apply_items = instructions.get("apply_pStyle", [])
    if not isinstance(apply_items, list) or len(apply_items) > MAX_CLASSIFICATION_AUDIT_RECORDS:
        raise ValueError("classification_audit.json apply_pStyle must be an array")
    declared_style_ids = {
        spec.get("style_id") for spec in style_roles.values() if isinstance(spec, dict)
    }
    for position, item in enumerate(apply_items):
        if not isinstance(item, dict) or set(item) != {"paragraph_index", "styleId"}:
            raise ValueError(
                f"classification_audit.json apply_pStyle[{position}] has an invalid shape"
            )
        index = item["paragraph_index"]
        style_id = item["styleId"]
        if type(index) is not int or index < 0 or index in applied:
            raise ValueError("classification_audit.json apply_pStyle indices must be unique non-negative integers")
        if not isinstance(style_id, str) or style_id not in declared_style_ids:
            raise ValueError(
                f"classification_audit.json apply_pStyle[{position}] uses an undeclared role style"
            )
        applied[index] = style_id

    ignored: Dict[int, str] = {}
    ignored_items = instructions.get("ignored_paragraphs", [])
    if not isinstance(ignored_items, list) or len(ignored_items) > MAX_CLASSIFICATION_AUDIT_RECORDS:
        raise ValueError("classification_audit.json ignored_paragraphs must be an array")
    for position, item in enumerate(ignored_items):
        if not isinstance(item, dict) or set(item) != {"paragraph_index", "reason"}:
            raise ValueError(
                f"classification_audit.json ignored_paragraphs[{position}] has an invalid shape"
            )
        index = item["paragraph_index"]
        reason = item["reason"]
        if type(index) is not int or index < 0 or index in ignored or index in applied:
            raise ValueError(
                "classification_audit.json styled and ignored indices must be unique and disjoint"
            )
        if not isinstance(reason, str) or not reason.strip():
            raise ValueError(
                f"classification_audit.json ignored_paragraphs[{position}].reason is invalid"
            )
        ignored[index] = reason
    if set(applied) | set(ignored) != expected_set:
        raise ValueError(
            "classification_audit.json instructions do not exactly partition expected paragraphs"
        )

    paragraphs = audit["paragraphs"]
    if not isinstance(paragraphs, list) or len(paragraphs) > MAX_CLASSIFICATION_AUDIT_RECORDS:
        raise ValueError("classification_audit.json paragraphs must be an array")
    seen: Set[int] = set()
    previous = -1
    for position, paragraph in enumerate(paragraphs):
        context = f"classification_audit.json paragraphs[{position}]"
        if not isinstance(paragraph, dict):
            raise ValueError(f"{context} must be an object")
        index = paragraph.get("paragraph_index")
        if type(index) is not int or index < 0 or index in seen or index <= previous:
            raise ValueError("classification_audit.json paragraphs must have unique sorted indices")
        seen.add(index)
        previous = index
        text = paragraph.get("text")
        if not isinstance(text, str):
            raise ValueError(f"{context}.text must be a string")
        classification = paragraph.get("classification")
        expected_keys = {
            "paragraph_index", "text", "slim_text_sha256", "text_was_truncated",
            "skip_reason", "classification",
        }
        if classification == "styled":
            expected_keys.add("style_id")
        elif classification == "ignored":
            expected_keys.add("ignore_reason")
        if set(paragraph) != expected_keys:
            raise ValueError(f"{context} has an invalid shape")
        if paragraph.get("slim_text_sha256") != hashlib.sha256(text.encode("utf-8")).hexdigest():
            raise ValueError(f"{context}.slim_text_sha256 mismatch")
        if not isinstance(paragraph.get("text_was_truncated"), bool):
            raise ValueError(f"{context}.text_was_truncated must be boolean")
        if paragraph.get("skip_reason") is not None and not isinstance(
            paragraph.get("skip_reason"), str
        ):
            raise ValueError(f"{context}.skip_reason must be a string or null")
        if classification == "styled":
            if paragraph.get("style_id") != applied.get(index):
                raise ValueError(f"{context}.style_id does not match instructions")
        elif classification == "ignored":
            if paragraph.get("ignore_reason") != ignored.get(index):
                raise ValueError(f"{context}.ignore_reason does not match instructions")
        elif classification == "out_of_scope":
            if index in applied or index in ignored:
                raise ValueError(f"{context}.classification does not match instructions")
        else:
            raise ValueError(f"{context}.classification is invalid")
    if expected_set - seen:
        raise ValueError("classification_audit.json is missing expected paragraph records")
    if _canonical_json_sha256(paragraphs) != audit["paragraphs_sha256"]:
        raise ValueError("classification_audit.json paragraphs_sha256 mismatch")


def _require_exact_keys(raw: Dict[str, Any], expected: Set[str], context: str) -> None:
    actual = set(raw)
    if actual != expected:
        raise ValueError(
            f"{context} keys do not match the v1 contract; "
            f"missing={sorted(expected - actual)}, unexpected={sorted(actual - expected)}"
        )


def _validate_manifest_producer(raw: Any) -> None:
    if not isinstance(raw, dict):
        raise ValueError("manifest.producer must be an object")
    required = {"name", "version", "run_id"}
    allowed = required | {"classifier", "prompts"}
    missing = required - set(raw)
    unexpected = set(raw) - allowed
    if missing or unexpected:
        raise ValueError(
            "manifest.producer has invalid keys; "
            f"missing={sorted(missing)}, unexpected={sorted(unexpected)}"
        )
    for key in required:
        if not isinstance(raw[key], str) or not raw[key].strip():
            raise ValueError(f"manifest.producer.{key} must be a non-empty string")
    if "classifier" in raw:
        classifier = raw["classifier"]
        if not isinstance(classifier, dict):
            raise ValueError("manifest.producer.classifier must be an object")
        _require_exact_keys(classifier, {"provider", "model"}, "manifest.producer.classifier")
        for key in ("provider", "model"):
            if not isinstance(classifier[key], str) or not classifier[key].strip():
                raise ValueError(
                    f"manifest.producer.classifier.{key} must be a non-empty string"
                )
    if "prompts" in raw:
        prompts = raw["prompts"]
        allowed_prompts = {"master_prompt_sha256", "run_instruction_sha256"}
        if not isinstance(prompts, dict) or not prompts or set(prompts) - allowed_prompts:
            raise ValueError("manifest.producer.prompts contains no known prompt hashes")
        for key, value in prompts.items():
            if not isinstance(value, str) or not _SHA256_RE.fullmatch(value):
                raise ValueError(f"manifest.producer.prompts.{key} must be a SHA-256 digest")


def validate_phase1_bundle_directory(
    bundle_dir: Path,
) -> Tuple[Dict[str, Any], Dict[str, Path]]:
    """Validate a complete Phase 1 bundle before any document mutation."""
    bundle_dir = Path(bundle_dir)
    if not bundle_dir.is_dir():
        raise ValueError(f"Phase 1 bundle directory does not exist: {bundle_dir}")
    manifest_path = bundle_dir / PHASE1_MANIFEST_FILENAME
    if manifest_path.is_symlink() or not manifest_path.is_file():
        raise FileNotFoundError(
            f"Required {PHASE1_MANIFEST_FILENAME} not found in Phase 1 bundle: {bundle_dir}"
        )
    if manifest_path.stat().st_size > MAX_PHASE1_MANIFEST_BYTES:
        raise ValueError(
            f"Phase 1 bundle manifest exceeds {MAX_PHASE1_MANIFEST_BYTES} bytes"
        )
    try:
        manifest = json.loads(manifest_path.read_text(encoding="utf-8"))
    except (UnicodeDecodeError, json.JSONDecodeError) as exc:
        raise ValueError(f"Phase 1 bundle manifest is not valid UTF-8 JSON: {exc}") from exc
    if not isinstance(manifest, dict):
        raise ValueError("Phase 1 bundle manifest must be a JSON object")
    _require_exact_keys(
        manifest,
        {
            "bundle_format",
            "manifest_version",
            "bundle_id",
            "created_utc",
            "producer",
            "source",
            "required_artifacts",
            "artifacts",
        },
        "manifest",
    )
    if manifest["bundle_format"] != PHASE1_BUNDLE_FORMAT:
        raise ValueError(f"Unsupported bundle_format: {manifest['bundle_format']!r}")
    if type(manifest["manifest_version"]) is not int or manifest["manifest_version"] != PHASE1_MANIFEST_VERSION:
        raise ValueError(f"Unsupported manifest_version: {manifest['manifest_version']!r}")
    if not isinstance(manifest["bundle_id"], str) or not manifest["bundle_id"].strip():
        raise ValueError("manifest.bundle_id must be a non-empty string")
    created_utc = manifest["created_utc"]
    if not isinstance(created_utc, str):
        raise ValueError("manifest.created_utc must be an ISO-8601 date-time")
    try:
        parsed_created = datetime.fromisoformat(created_utc.replace("Z", "+00:00"))
    except ValueError as exc:
        raise ValueError("manifest.created_utc must be an ISO-8601 date-time") from exc
    if parsed_created.tzinfo is None:
        raise ValueError("manifest.created_utc must include a timezone")
    _validate_manifest_producer(manifest["producer"])

    source = manifest["source"]
    if not isinstance(source, dict):
        raise ValueError("manifest.source must be an object")
    _require_exact_keys(source, {"filename", "sha256", "size_bytes"}, "manifest.source")
    filename = source["filename"]
    if (
        not isinstance(filename, str)
        or not filename
        or filename in {".", ".."}
        or "/" in filename
        or "\\" in filename
    ):
        raise ValueError("manifest.source.filename must be a non-empty basename")
    if not isinstance(source["sha256"], str) or not _SHA256_RE.fullmatch(source["sha256"]):
        raise ValueError("manifest.source.sha256 must be a lowercase SHA-256 digest")
    if type(source["size_bytes"]) is not int or source["size_bytes"] < 0:
        raise ValueError("manifest.source.size_bytes must be a non-negative integer")

    required = manifest["required_artifacts"]
    if (
        not isinstance(required, list)
        or len(required) != len(set(required))
        or set(required) != set(PHASE1_REQUIRED_ARTIFACT_IDS)
    ):
        raise ValueError(
            "manifest.required_artifacts must contain exactly "
            f"{list(PHASE1_REQUIRED_ARTIFACT_IDS)}"
        )

    artifacts_raw = manifest["artifacts"]
    if not isinstance(artifacts_raw, list):
        raise ValueError("manifest.artifacts must be an array")
    artifacts: Dict[str, Path] = {}
    declared_names = {PHASE1_MANIFEST_FILENAME}
    seen_paths: Set[str] = set()
    artifact_keys = {
        "artifact_id", "path", "media_type", "sha256", "size_bytes", "required", "source_kind"
    }
    for index, record in enumerate(artifacts_raw):
        context = f"manifest.artifacts[{index}]"
        if not isinstance(record, dict):
            raise ValueError(f"{context} must be an object")
        _require_exact_keys(record, artifact_keys, context)
        artifact_id = record["artifact_id"]
        if artifact_id not in _PHASE1_ARTIFACT_SPECS:
            raise ValueError(f"{context}.artifact_id is unknown: {artifact_id!r}")
        if artifact_id in artifacts:
            raise ValueError(f"manifest.artifacts contains duplicate artifact_id: {artifact_id}")
        expected_path, expected_media, expected_required, expected_kind = _PHASE1_ARTIFACT_SPECS[artifact_id]
        expected_values = {
            "path": expected_path,
            "media_type": expected_media,
            "required": expected_required,
            "source_kind": expected_kind,
        }
        for key, expected in expected_values.items():
            if record[key] != expected:
                raise ValueError(
                    f"{context}.{key} must be {expected!r} for {artifact_id}, got {record[key]!r}"
                )
        if record["path"] in seen_paths:
            raise ValueError(f"manifest.artifacts contains duplicate path: {record['path']}")
        seen_paths.add(record["path"])
        if not isinstance(record["sha256"], str) or not _SHA256_RE.fullmatch(record["sha256"]):
            raise ValueError(f"{context}.sha256 must be a lowercase SHA-256 digest")
        if type(record["size_bytes"]) is not int or record["size_bytes"] < 0:
            raise ValueError(f"{context}.size_bytes must be a non-negative integer")
        if record["size_bytes"] > MAX_PHASE1_ARTIFACT_BYTES:
            raise ValueError(
                f"{context}.size_bytes exceeds the safe artifact limit "
                f"({MAX_PHASE1_ARTIFACT_BYTES} bytes)"
            )

        artifact_path = bundle_dir / record["path"]
        if artifact_path.is_symlink() or not artifact_path.is_file():
            raise ValueError(f"Bundle artifact is missing or not a regular file: {record['path']}")
        actual_size = artifact_path.stat().st_size
        if actual_size != record["size_bytes"]:
            raise ValueError(
                f"Bundle artifact size mismatch for {record['path']}: "
                f"manifest={record['size_bytes']}, actual={actual_size}"
            )
        actual_sha = _sha256_file(artifact_path)
        if actual_sha != record["sha256"]:
            raise ValueError(
                f"Bundle artifact SHA-256 mismatch for {record['path']}: "
                f"manifest={record['sha256']}, actual={actual_sha}"
            )
        artifacts[artifact_id] = artifact_path
        declared_names.add(record["path"])

    expected_ids = set(PHASE1_REQUIRED_ARTIFACT_IDS) | (
        {"source_settings"} if "source_settings" in artifacts else set()
    )
    if set(artifacts) != expected_ids:
        raise ValueError(f"manifest.artifacts has an invalid artifact set: {sorted(artifacts)}")
    actual_names = {entry.name for entry in bundle_dir.iterdir()}
    extras = sorted(actual_names - declared_names)
    if extras:
        raise ValueError(f"Bundle contains unlisted/stale artifacts: {extras}")

    try:
        style_registry = json.loads(artifacts["style_registry"].read_text(encoding="utf-8"))
        template_registry = json.loads(artifacts["template_registry"].read_text(encoding="utf-8"))
        classification_audit = json.loads(
            artifacts["classification_audit"].read_text(encoding="utf-8")
        )
    except (UnicodeDecodeError, json.JSONDecodeError) as exc:
        raise ValueError(f"Bundle registry artifact is not valid UTF-8 JSON: {exc}") from exc
    if not isinstance(style_registry, dict) or not isinstance(template_registry, dict):
        raise ValueError("Bundle registry artifacts must be JSON objects")
    if not isinstance(classification_audit, dict):
        raise ValueError("classification_audit.json must be a JSON object")
    required_template_sections = {
        "meta", "package_inventory", "doc_defaults", "styles", "theme",
        "settings", "page_layout", "headers_footers", "numbering", "fonts",
        "custom_xml", "capture_policy",
    }
    missing_template_sections = required_template_sections - set(template_registry)
    if missing_template_sections:
        raise ValueError(
            "arch_template_registry.json is missing required sections: "
            f"{sorted(missing_template_sections)}"
        )
    _require_exact_keys(
        style_registry,
        {"version", "source_docx", "source_sha256", "source_tokens", "roles"},
        "arch_style_registry.json",
    )
    if style_registry.get("version") != 2:
        raise ValueError("Strict Phase 1 bundles require arch_style_registry.json version 2")
    roles = style_registry.get("roles")
    if not isinstance(roles, dict) or not roles:
        raise ValueError("arch_style_registry.json roles must be a non-empty object")
    unknown_roles = set(roles) - _ALLOWED_ROLE_NAMES
    if unknown_roles:
        raise ValueError(f"arch_style_registry.json contains unknown roles: {sorted(unknown_roles)}")
    allowed_role_keys = {
        "style_id", "exemplar_paragraph_index", "style_name", "resolved_formatting",
        "warning", "numbering_provenance", "numbering_pattern",
    }
    allowed_pattern_keys = {
        "numId", "ilvl", "abstractNumId", "startOverride", "numFmt", "lvlText",
    }
    for role, spec in roles.items():
        if not isinstance(spec, dict):
            raise ValueError(f"roles[{role!r}] must be an object")
        unexpected = set(spec) - allowed_role_keys
        if unexpected:
            raise ValueError(f"roles[{role!r}] contains unknown fields: {sorted(unexpected)}")
        if not isinstance(spec.get("style_id"), str) or not spec["style_id"].strip():
            raise ValueError(f"roles[{role!r}].style_id must be a non-empty string")
        exemplar = spec.get("exemplar_paragraph_index")
        if type(exemplar) is not int or exemplar < 0:
            raise ValueError(
                f"roles[{role!r}].exemplar_paragraph_index must be a non-negative integer"
            )
        provenance = spec.get("numbering_provenance")
        if provenance not in {"style_numpr", "direct_numpr", "text_literal", "none"}:
            raise ValueError(f"roles[{role!r}] has invalid or missing numbering_provenance")
        pattern = spec.get("numbering_pattern")
        if pattern is not None:
            if not isinstance(pattern, dict) or set(pattern) - allowed_pattern_keys:
                raise ValueError(f"roles[{role!r}].numbering_pattern is invalid")
            if any(not isinstance(value, str) for value in pattern.values()):
                raise ValueError(f"roles[{role!r}].numbering_pattern values must be strings")
        if provenance in {"style_numpr", "direct_numpr"}:
            if not isinstance(pattern, dict) or not pattern.get("numId"):
                raise ValueError(
                    f"roles[{role!r}] provenance {provenance} requires numbering_pattern.numId"
                )
        elif pattern is not None:
            raise ValueError(
                f"roles[{role!r}] provenance {provenance} must not define numbering_pattern"
            )
    if style_registry.get("source_docx") != filename:
        raise ValueError("Style registry source_docx does not match manifest source filename")
    if style_registry.get("source_sha256") != source["sha256"]:
        raise ValueError("Style registry source_sha256 does not match manifest source SHA-256")
    _validate_classification_audit_linkage(
        classification_audit,
        source["sha256"],
        roles,
    )
    template_source = template_registry.get("meta", {}).get("source_docx", {})
    if template_registry.get("meta", {}).get("schema_version") != "1.0.0":
        raise ValueError("Unsupported arch_template_registry.json meta.schema_version")
    if not isinstance(template_source, dict) or template_source.get("filename") != filename:
        raise ValueError("Template registry source filename does not match manifest source filename")
    if template_source.get("sha256") != source["sha256"]:
        raise ValueError("Template registry source SHA-256 does not match manifest source SHA-256")
    package_inventory = template_registry.get("package_inventory")
    if not isinstance(package_inventory, dict) or not isinstance(
        package_inventory.get("has_settings"), bool
    ):
        raise ValueError(
            "arch_template_registry.json package_inventory.has_settings must be boolean"
        )
    if package_inventory["has_settings"] != ("source_settings" in artifacts):
        raise ValueError(
            "arch_template_registry.json settings inventory disagrees with source_settings artifact presence"
        )
    headers_footers = template_registry.get("headers_footers")
    if not isinstance(headers_footers, dict):
        raise ValueError("arch_template_registry.json headers_footers must be an object")
    for collection in ("headers", "footers"):
        if not isinstance(headers_footers.get(collection), list):
            raise ValueError(
                f"arch_template_registry.json headers_footers.{collection} must be an array"
            )

    xml_roots: Dict[str, _ET.Element] = {}
    for artifact_id in ("source_styles", "portable_styles", "source_settings"):
        path = artifacts.get(artifact_id)
        if path is None:
            continue
        try:
            xml_roots[artifact_id] = _ET.fromstring(path.read_bytes())
        except _ET.ParseError as exc:
            raise ValueError(f"Bundle artifact {artifact_id} is not well-formed XML: {exc}") from exc
    source_ids = {
        node.attrib.get(f"{{{W_NS}}}styleId")
        for node in xml_roots["source_styles"].findall(f"{{{W_NS}}}style")
        if node.attrib.get(f"{{{W_NS}}}styleId")
    }
    portable_ids = {
        node.attrib.get(f"{{{W_NS}}}styleId")
        for node in xml_roots["portable_styles"].findall(f"{{{W_NS}}}style")
        if node.attrib.get(f"{{{W_NS}}}styleId")
    }
    registry_ids = {
        item.get("style_id")
        for item in template_registry.get("styles", {}).get("style_defs", [])
        if isinstance(item, dict) and item.get("style_id")
    }
    if source_ids != registry_ids:
        raise ValueError(
            "source_styles and template registry style inventories differ; "
            f"missing={sorted(registry_ids - source_ids)[:20]}, "
            f"unexpected={sorted(source_ids - registry_ids)[:20]}"
        )
    role_ids = {
        spec.get("style_id")
        for spec in roles.values()
        if isinstance(spec, dict) and spec.get("style_id")
    }
    expected_portable_ids = source_ids | role_ids
    if portable_ids != expected_portable_ids:
        raise ValueError(
            "portable_styles must contain exactly the source styles plus declared role styles; "
            f"missing={sorted(expected_portable_ids - portable_ids)[:20]}, "
            f"unexpected={sorted(portable_ids - expected_portable_ids)[:20]}"
        )
    return manifest, artifacts


def load_role_specs_from_registry(registry_path: Path) -> Dict[str, Dict[str, Any]]:
    """Load the validated, non-flattened role contract."""
    path = Path(registry_path)
    if path.is_dir():
        path = path / "arch_style_registry.json"
    raw = json.loads(path.read_text(encoding="utf-8"))
    roles = raw.get("roles", {}) if isinstance(raw, dict) else {}
    return {
        role: dict(spec)
        for role, spec in roles.items()
        if isinstance(role, str) and isinstance(spec, dict)
    }


def _xml_escape_attr(value: str) -> str:
    """Escape a string for safe use inside an XML attribute value (double-quoted)."""
    return _sax_escape(str(value), {'"': "&quot;"})


def build_arch_styles_xml_from_registry(registry: Dict[str, Any]) -> str:
    """
    Reconstruct styles.xml for the explicit trusted-legacy compatibility path.

    Current strict Phase 2 loads the manifest-verified portable_styles.xml
    artifact. This JSON reconstruction remains only for callers that explicitly
    opt into a trusted pre-bundle registry directory.

    The output is a well-formed XML string containing <w:docDefaults> and all
    <w:style> blocks. The existing regex-based functions (extract_style_block_raw,
    _extract_basedOn, _find_style_numpr_in_chain, etc.) work on this string
    identically to how they work on a real styles.xml file.
    """
    style_defs = registry.get("styles", {}).get("style_defs", [])
    doc_defaults = registry.get("doc_defaults", {})

    default_rpr = doc_defaults.get("default_run_props", {}).get("rPr") or ""
    default_ppr = doc_defaults.get("default_paragraph_props", {}).get("pPr") or ""

    parts = [
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
        '<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
        ' xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"'
        ' xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"'
        ' xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml">',
    ]

    # docDefaults
    parts.append("<w:docDefaults>")
    if default_rpr:
        parts.append(f"<w:rPrDefault>{default_rpr}</w:rPrDefault>")
    else:
        parts.append("<w:rPrDefault><w:rPr/></w:rPrDefault>")
    if default_ppr:
        parts.append(f"<w:pPrDefault>{default_ppr}</w:pPrDefault>")
    else:
        parts.append("<w:pPrDefault><w:pPr/></w:pPrDefault>")
    parts.append("</w:docDefaults>")

    # All style definitions
    for sd in style_defs:
        sid = sd.get("style_id", "")
        if not sid:
            continue

        stype = sd.get("type", "paragraph")
        name = sd.get("name") or sid
        based_on = sd.get("based_on")
        next_style = sd.get("next")
        link = sd.get("link")

        # XML-escape all attribute values (NOT raw XML property fragments)
        e_sid = _xml_escape_attr(sid)
        e_stype = _xml_escape_attr(stype)
        e_name = _xml_escape_attr(name)

        parts.append(f'<w:style w:type="{e_stype}" w:styleId="{e_sid}">')
        parts.append(f'<w:name w:val="{e_name}"/>')
        if based_on:
            parts.append(f'<w:basedOn w:val="{_xml_escape_attr(based_on)}"/>')
        if next_style:
            parts.append(f'<w:next w:val="{_xml_escape_attr(next_style)}"/>')
        if link:
            parts.append(f'<w:link w:val="{_xml_escape_attr(link)}"/>')
        if sd.get("ui_priority") is not None:
            parts.append(f'<w:uiPriority w:val="{_xml_escape_attr(sd["ui_priority"])}"/>')
        if sd.get("semi_hidden"):
            parts.append("<w:semiHidden/>")
        if sd.get("unhide_when_used"):
            parts.append("<w:unhideWhenUsed/>")
        if sd.get("qformat"):
            parts.append("<w:qFormat/>")

        # Raw XML fragments — inserted verbatim, never escaped
        if sd.get("pPr"):
            parts.append(sd["pPr"])
        if sd.get("rPr"):
            parts.append(sd["rPr"])
        if sd.get("tblPr"):
            parts.append(sd["tblPr"])
        if sd.get("trPr"):
            parts.append(sd["trPr"])
        if sd.get("tcPr"):
            parts.append(sd["tcPr"])

        parts.append("</w:style>")

    parts.append("</w:styles>")
    result = "\n".join(parts)

    # Validate that the generated XML is well-formed
    try:
        _ET.fromstring(result.encode("utf-8"))
    except _ET.ParseError as exc:
        raise ValueError(
            f"Synthetic styles.xml failed XML well-formedness check: {exc}"
        ) from exc

    return result


def resolve_arch_extract_root(p: Path) -> Path:
    """
    Resolve a trusted legacy architect-registry directory.

    Accepts a path that contains arch_style_registry.json and
    arch_template_registry.json. The current Phase 1/Phase 2 interface is the
    manifest-verified bundle with portable_styles.xml; this helper is used only
    after an explicit legacy opt-in.

    Returns the directory path.
    """
    p = Path(p)

    # If they passed a file, use its parent directory
    if p.is_file():
        p = p.parent

    # Check for required contract files
    style_reg = p / "arch_style_registry.json"
    template_reg = p / "arch_template_registry.json"

    if not style_reg.exists():
        raise FileNotFoundError(
            f"arch_style_registry.json not found at: {style_reg}\n"
            "Point Phase 2 to the folder containing both JSON files from Phase 1."
        )
    if not template_reg.exists():
        raise FileNotFoundError(
            f"arch_template_registry.json not found at: {template_reg}\n"
            "Point Phase 2 to the folder containing both JSON files from Phase 1."
        )

    return p


def load_available_roles_from_registry(registry_path: Path) -> Optional[List[str]]:
    """
    Load the list of available role names from arch_style_registry.json.

    Args:
        registry_path: Path to arch_style_registry.json or the extracted folder containing it

    Returns:
        List of role names (e.g., ["SectionTitle", "PART", "ARTICLE", ...])
        Returns None if registry not found.
    """
    registry_path = Path(registry_path)

    # Handle both direct JSON path and folder path
    if registry_path.is_dir():
        registry_path = registry_path / "arch_style_registry.json"

    if not registry_path.exists():
        return None

    registry = json.loads(registry_path.read_text(encoding="utf-8"))
    roles = registry.get("roles", {})

    return sorted(roles.keys())


def load_arch_style_registry(arch_extract_dir: Path) -> Dict[str, str]:
    """
    Phase 2 contract (STRICT):
    - arch_style_registry.json must exist (emitted by Phase 1).
    - NO inference / NO heuristics.
    Returns: { role: styleId }
    """
    arch_extract_dir = Path(arch_extract_dir)

    # Allow passing the registry JSON directly
    if arch_extract_dir.is_file() and arch_extract_dir.suffix.lower() == ".json":
        reg_path = arch_extract_dir
        root_dir = arch_extract_dir.parent
    else:
        root_dir = resolve_arch_extract_root(arch_extract_dir)
        reg_path = root_dir / "arch_style_registry.json"

    if not reg_path.exists():
        raise FileNotFoundError(
            f"arch_style_registry.json not found at {reg_path}. "
            f"Run Phase 1 on the architect template and copy the extracted folder here."
        )

    reg = json.loads(reg_path.read_text(encoding="utf-8"))
    if not isinstance(reg, dict):
        raise ValueError("arch_style_registry.json must be a JSON object")

    # Expected shape:
    # { "version": 1, "source_docx": "...", "roles": { "PART": { "style_id": "X", ... }, ... } }
    roles = reg.get("roles")
    if not isinstance(roles, dict):
        raise ValueError("arch_style_registry.json missing 'roles' object")

    out: Dict[str, str] = {}
    for role, info in roles.items():
        if not isinstance(role, str) or not isinstance(info, dict):
            continue
        sid = info.get("style_id") or info.get("styleId")
        if isinstance(sid, str) and sid.strip():
            out[role.strip()] = sid.strip()

    if not out:
        raise ValueError("arch_style_registry.json contained no usable role->style mappings")

    return out


def write_phase2_preflight(
    extract_dir: Path,
    arch_root: Path,
    arch_registry: Dict[str, str],
    classifications: Dict[str, Any],
    out_path: Path
) -> Dict[str, Any]:
    # Count classifications per role
    role_counts: Dict[str, int] = {}
    for item in classifications.get("classifications", []):
        r = item.get("csi_role")
        if isinstance(r, str):
            role_counts[r] = role_counts.get(r, 0) + 1

    # Identify which roles are unmapped
    needed_roles = sorted(role_counts.keys())
    unmapped_roles = [r for r in needed_roles if r not in arch_registry]

    report = {
        "arch_extract_root": str(arch_root),
        "target_extract_root": str(extract_dir),
        "roles_in_classifications": role_counts,
        "arch_style_registry": arch_registry,
        "unmapped_roles": unmapped_roles,
    }

    out_path.write_text(json.dumps(report, indent=2), encoding="utf-8")
    return report


# ---------------------------------------------------------------------------
# Phase 2 preflight contract validation
# ---------------------------------------------------------------------------

_EXPECTED_TEMPLATE_SECTIONS = {
    "theme": dict,
    "settings": dict,
    "fonts": dict,
    "doc_defaults": dict,
    "styles": dict,
    "numbering": dict,
    "page_layout": dict,
}

# Maps style_def property keys to the Word XML tag they should contain.
_STYLE_PR_TAG_MAP = {
    "pPr": "w:pPr",
    "rPr": "w:rPr",
    "tblPr": "w:tblPr",
    "trPr": "w:trPr",
    "tcPr": "w:tcPr",
}


def _check_xml_fragment(fragment: str, expected_tag: str) -> Optional[str]:
    """Return an error message if *fragment* lacks matching open/close tags for *expected_tag*."""
    if not isinstance(fragment, str) or not fragment.strip():
        return None  # empty/absent is fine — caller decides whether the key is required
    escaped = re.escape(expected_tag)
    # Self-closing is valid: <w:pPr/>
    if re.search(r"<" + escaped + r"(?:\s[^>]*)?\s*/\s*>", fragment):
        return None
    has_open = bool(re.search(r"<" + escaped + r"[\s>/]", fragment))
    has_close = bool(re.search(r"</" + escaped + r"\s*>", fragment))
    if not has_open or not has_close:
        return f"XML fragment for <{expected_tag}> is malformed: missing open or close tag"
    return None


def _validate_template_sections(
    template_registry: Dict[str, Any], errors: List[str]
) -> None:
    """Check 1: top-level section types."""
    for key, expected_type in _EXPECTED_TEMPLATE_SECTIONS.items():
        if key in template_registry:
            val = template_registry[key]
            if not isinstance(val, expected_type):
                errors.append(
                    f"Template registry section '{key}' must be {expected_type.__name__}, "
                    f"got {type(val).__name__}"
                )


def _validate_style_defs(
    template_registry: Dict[str, Any], errors: List[str]
) -> Set[str]:
    """Checks 2-4: style_defs is list, style_ids unique/usable, XML fragments parseable.

    Returns the set of known style IDs for cross-reference validation.
    """
    known_ids: Set[str] = set()
    styles_section = template_registry.get("styles")
    if styles_section is None:
        return known_ids
    if not isinstance(styles_section, dict):
        return known_ids  # already caught by _validate_template_sections

    style_defs = styles_section.get("style_defs")
    if style_defs is None:
        return known_ids
    if not isinstance(style_defs, list):
        errors.append(
            f"styles.style_defs must be a list, got {type(style_defs).__name__}"
        )
        return known_ids

    seen_ids: Dict[str, int] = {}  # style_id -> first index
    for idx, sd in enumerate(style_defs):
        if not isinstance(sd, dict):
            errors.append(f"styles.style_defs[{idx}] must be a dict, got {type(sd).__name__}")
            continue

        sid = sd.get("style_id")
        if not isinstance(sid, str) or not sid.strip():
            errors.append(f"styles.style_defs[{idx}] missing or empty 'style_id'")
            continue

        sid = sid.strip()
        if sid in seen_ids:
            errors.append(
                f"Duplicate style_id '{sid}' in styles.style_defs "
                f"(indices {seen_ids[sid]} and {idx})"
            )
        else:
            seen_ids[sid] = idx
        known_ids.add(sid)

        # Validate XML property fragments
        for pr_key, tag in _STYLE_PR_TAG_MAP.items():
            val = sd.get(pr_key)
            if val:
                err = _check_xml_fragment(val, tag)
                if err:
                    errors.append(f"styles.style_defs[{idx}] ('{sid}'): {err}")

    return known_ids


def _validate_compat_xml(
    template_registry: Dict[str, Any], errors: List[str]
) -> None:
    """Check 5: settings.compat.compat_xml is a full valid block."""
    settings = template_registry.get("settings")
    if not isinstance(settings, dict):
        return
    compat = settings.get("compat")
    if not isinstance(compat, dict):
        return
    compat_xml = compat.get("compat_xml")
    if not compat_xml:
        return
    if not isinstance(compat_xml, str):
        errors.append(
            f"settings.compat.compat_xml must be a string, got {type(compat_xml).__name__}"
        )
        return
    err = _check_xml_fragment(compat_xml, "w:compat")
    if err:
        errors.append(f"settings.compat.compat_xml: {err}")


def _validate_top_level_xml_fragments(
    template_registry: Dict[str, Any], errors: List[str]
) -> None:
    """Check 4 (continued): validate top-level XML fragments (theme, fonts)."""
    theme = template_registry.get("theme")
    if isinstance(theme, dict):
        xml = theme.get("theme1_xml")
        if xml:
            err = _check_xml_fragment(xml, "a:theme")
            if err:
                errors.append(f"theme.theme1_xml: {err}")

    fonts = template_registry.get("fonts")
    if isinstance(fonts, dict):
        xml = fonts.get("font_table_xml")
        if xml:
            err = _check_xml_fragment(xml, "w:fonts")
            if err:
                errors.append(f"fonts.font_table_xml: {err}")


def _parse_registry_xml(
    xml_text: str,
    context: str,
    errors: List[str],
) -> Optional[_ET.Element]:
    """Parse registry-carried XML after making its declaration truthful."""
    if "<!DOCTYPE" in xml_text.upper():
        errors.append(f"{context} must not contain a DOCTYPE declaration")
        return None
    try:
        normalized = prepare_xml_text_for_utf8(xml_text).encode("utf-8")
        return _ET.fromstring(normalized)
    except (UnicodeEncodeError, _ET.ParseError) as exc:
        errors.append(f"{context} is malformed XML: {exc}")
        return None


def _resolve_internal_relationship_target(source_part: str, target: str) -> str:
    """Resolve a safe internal OPC relationship target to a package part."""
    try:
        return _resolve_safe_internal_relationship_target(source_part, target)
    except ValueError as exc:
        raise ValueError(f"unsafe relationship target {target!r}: {exc}") from exc


def _media_target_candidates(target: str) -> Set[str]:
    """Return package-part interpretations accepted by the v1 registry."""
    if not target or "\x00" in target:
        return set()

    decoded_target = unquote(target)
    parsed = urlsplit(decoded_target)
    windows_target = PureWindowsPath(decoded_target)
    target_path = PurePosixPath(decoded_target)
    if (
        not decoded_target
        or "\x00" in decoded_target
        or parsed.scheme
        or parsed.netloc
        or parsed.query
        or parsed.fragment
        or "\\" in decoded_target
        or decoded_target.startswith(("/", "//"))
        or windows_target.drive
        or windows_target.is_absolute()
        or target_path.is_absolute()
        or ".." in target_path.parts
        or any(":" in part for part in target_path.parts)
    ):
        return set()
    normalized = PurePosixPath(
        *(part for part in target_path.parts if part not in {"", "."})
    ).as_posix()
    if normalized in {"", "."}:
        return set()
    candidates = {normalized}
    if not normalized.startswith("word/"):
        candidates.add(f"word/{normalized}")
    return candidates


def _relationship_ids_used_by_part(
    root: _ET.Element,
    context: str,
    errors: List[str],
) -> Set[str]:
    ids: Set[str] = set()
    for element in root.iter():
        for attr_name, value in element.attrib.items():
            if not attr_name.startswith("{") or "}" not in attr_name:
                continue
            namespace, local_name = attr_name[1:].split("}", 1)
            if namespace.endswith("/relationships") and local_name in {
                "id", "embed", "link"
            }:
                if not value.strip():
                    errors.append(
                        f"{context} contains an empty r:{local_name} relationship reference"
                    )
                else:
                    ids.add(value)
    return ids


def _validate_embedded_fonts(
    template_registry: Dict[str, Any], errors: List[str]
) -> None:
    fonts = template_registry.get("fonts")
    if not isinstance(fonts, dict):
        return
    font_xml = fonts.get("font_table_xml")
    if not isinstance(font_xml, str) or not font_xml.strip():
        return
    root = _parse_registry_xml(font_xml, "fonts.font_table_xml", errors)
    if root is not None and any(node.tag in _EMBEDDED_FONT_TAGS for node in root.iter()):
        errors.append(
            "fonts.font_table_xml contains embedded-font references; Phase 2 "
            "does not support embedded font payloads"
        )


def _validate_header_footer_contract(
    template_registry: Dict[str, Any], errors: List[str]
) -> None:
    """Validate the v1 header/footer relationship and media payload contract."""
    hf = template_registry.get("headers_footers")
    if hf is None:
        return
    if not isinstance(hf, dict):
        errors.append("headers_footers must be a dict")
        return

    document_rel_ids: Dict[str, str] = {}
    owner_parts: Dict[str, Tuple[str, str, Dict[str, Any]]] = {}
    owner_payload_keys = (
        "xml",
        "part_xml",
        "rels_part_name",
        "rels_xml",
        "relationships_xml",
        "media",
        "media_files",
    )
    total_media_bytes = 0
    for collection, kind in (("headers", "header"), ("footers", "footer")):
        entries = hf.get(collection)
        if not isinstance(entries, list):
            errors.append(f"headers_footers.{collection} must be a list")
            continue

        for index, entry in enumerate(entries):
            context = f"headers_footers.{collection}[{index}]"
            if not isinstance(entry, dict):
                errors.append(f"{context} must be a dict")
                continue

            part_name = entry.get("part_name")
            if (
                not isinstance(part_name, str)
                or not is_safe_header_footer_part_name(
                    part_name,
                    expected_kind=kind,
                )
            ):
                errors.append(f"{context}.part_name is invalid for a {kind}")
                continue

            owner_key = part_name.casefold()
            previous_owner = owner_parts.get(owner_key)
            if previous_owner is None:
                owner_parts[owner_key] = (part_name, kind, entry)
            else:
                previous_name, previous_kind, previous_entry = previous_owner
                if previous_name != part_name or previous_kind != kind:
                    errors.append(
                        f"{context}.part_name collides case-insensitively or by kind "
                        f"with {previous_name!r} ({previous_kind})"
                    )
                elif any(
                    previous_entry.get(payload_key) != entry.get(payload_key)
                    for payload_key in owner_payload_keys
                ):
                    errors.append(
                        f"{context} repeats owner part {part_name!r} with conflicting content"
                    )

            rels_part_name = entry.get("rels_part_name")
            expected_rels_part_name = relationship_part_name_for_owner(part_name)
            if rels_part_name is not None and (
                not isinstance(rels_part_name, str)
                or rels_part_name != expected_rels_part_name
            ):
                errors.append(
                    f"{context}.rels_part_name must be {expected_rels_part_name!r}"
                )

            document_rel_id = entry.get("rel_id")
            if document_rel_id is not None:
                if not isinstance(document_rel_id, str) or not document_rel_id:
                    errors.append(f"{context}.rel_id must be a non-empty string or null")
                elif document_rel_id in document_rel_ids:
                    errors.append(
                        f"{context}.rel_id duplicates document relationship "
                        f"{document_rel_id!r}"
                    )
                else:
                    document_rel_ids[document_rel_id] = kind

            part_xml = entry.get("xml")
            part_root = None
            if not isinstance(part_xml, str) or not part_xml.strip():
                errors.append(f"{context}.xml must be a non-empty string")
            else:
                part_root = _parse_registry_xml(part_xml, f"{context}.xml", errors)

            rels_xml = entry.get("rels_xml")
            relationships: Dict[str, Dict[str, str]] = {}
            if rels_xml is not None:
                if not isinstance(rels_xml, str) or not rels_xml.strip():
                    errors.append(f"{context}.rels_xml must be a non-empty string or null")
                else:
                    rels_root = _parse_registry_xml(
                        rels_xml, f"{context}.rels_xml", errors
                    )
                    if rels_root is not None:
                        expected_root = f"{{{_PKG_REL_NS}}}Relationships"
                        if rels_root.tag != expected_root:
                            errors.append(
                                f"{context}.rels_xml has an invalid Relationships root"
                            )
                        else:
                            for rel_index, rel in enumerate(rels_root):
                                rel_context = (
                                    f"{context}.rels_xml.Relationship[{rel_index}]"
                                )
                                if rel.tag != f"{{{_PKG_REL_NS}}}Relationship":
                                    errors.append(
                                        f"{rel_context} has an unsupported element"
                                    )
                                    continue
                                rel_id = rel.get("Id")
                                rel_type = rel.get("Type")
                                target = rel.get("Target")
                                target_mode = rel.get("TargetMode")
                                if not all(
                                    isinstance(value, str) and value
                                    for value in (rel_id, rel_type, target)
                                ):
                                    errors.append(
                                        f"{rel_context} requires non-empty Id, Type, and Target"
                                    )
                                    continue
                                if rel_id in relationships:
                                    errors.append(
                                        f"{context}.rels_xml contains duplicate relationship "
                                        f"Id {rel_id!r}"
                                    )
                                    continue
                                if target_mode not in {None, "Internal", "External"}:
                                    errors.append(
                                        f"{rel_context} has unsupported TargetMode "
                                        f"{target_mode!r}"
                                    )
                                    continue
                                relationship = {
                                    "type": rel_type,
                                    "target": target,
                                    "mode": target_mode or "Internal",
                                }
                                relationships[rel_id] = relationship
                                if target_mode == "External":
                                    continue
                                if not rel_type.endswith("/image"):
                                    errors.append(
                                        f"{rel_context} is an unsupported internal "
                                        f"non-image relationship ({rel_type!r})"
                                    )
                                    continue
                                try:
                                    relationship["package_part"] = (
                                        _resolve_internal_relationship_target(
                                            part_name, target
                                        )
                                    )
                                except ValueError as exc:
                                    errors.append(f"{rel_context}: {exc}")

            media = entry.get("media", [])
            media_by_rel_id: Dict[str, Dict[str, Any]] = {}
            if not isinstance(media, list):
                errors.append(f"{context}.media must be a list")
                media = []
            for media_index, item in enumerate(media):
                media_context = f"{context}.media[{media_index}]"
                if not isinstance(item, dict):
                    errors.append(f"{media_context} must be a dict")
                    continue
                rel_id = item.get("rel_id")
                target = item.get("target")
                content_type = item.get("content_type")
                payload = item.get("data_base64")
                if not isinstance(rel_id, str) or not rel_id:
                    errors.append(f"{media_context}.rel_id must be a non-empty string")
                    continue
                if rel_id in media_by_rel_id:
                    errors.append(f"{media_context}.rel_id duplicates {rel_id!r}")
                    continue
                media_by_rel_id[rel_id] = item
                if not isinstance(target, str) or not _media_target_candidates(target):
                    errors.append(f"{media_context}.target is unsafe or empty")
                if (
                    not isinstance(content_type, str)
                    or len(content_type) > 255
                    or any(ord(character) < 32 for character in content_type)
                    or not _MIME_TYPE_RE.fullmatch(content_type)
                    or not content_type.casefold().startswith("image/")
                ):
                    errors.append(f"{media_context}.content_type must be an image media type")
                if not isinstance(payload, str):
                    errors.append(f"{media_context}.data_base64 must be a string")
                else:
                    max_encoded_size = 4 * ((MAX_HEADER_FOOTER_MEDIA_BYTES + 2) // 3)
                    if len(payload) > max_encoded_size:
                        errors.append(
                            f"{media_context}.data_base64 exceeds the "
                            f"{MAX_HEADER_FOOTER_MEDIA_BYTES}-byte decoded media limit"
                        )
                    else:
                        try:
                            decoded_payload = base64.b64decode(payload, validate=True)
                        except (binascii.Error, ValueError) as exc:
                            errors.append(
                                f"{media_context}.data_base64 is invalid base64: {exc}"
                            )
                        else:
                            if len(decoded_payload) > MAX_HEADER_FOOTER_MEDIA_BYTES:
                                errors.append(
                                    f"{media_context} decodes to {len(decoded_payload)} bytes; "
                                    f"limit is {MAX_HEADER_FOOTER_MEDIA_BYTES} bytes"
                                )
                            else:
                                total_media_bytes += len(decoded_payload)
                                if total_media_bytes > MAX_HEADER_FOOTER_MEDIA_TOTAL_BYTES:
                                    errors.append(
                                        "headers_footers captured media exceeds the total "
                                        f"{MAX_HEADER_FOOTER_MEDIA_TOTAL_BYTES}-byte limit"
                                    )

                relationship = relationships.get(rel_id)
                if relationship is None:
                    errors.append(
                        f"{media_context} references missing relationship Id {rel_id!r}"
                    )
                    continue
                if relationship.get("mode") == "External" or not relationship.get(
                    "type", ""
                ).endswith("/image"):
                    errors.append(
                        f"{media_context} does not reference a supported internal image"
                    )
                    continue
                package_part = relationship.get("package_part")
                if (
                    isinstance(target, str)
                    and package_part
                    and package_part not in _media_target_candidates(target)
                ):
                    errors.append(
                        f"{media_context}.target does not match relationship {rel_id!r} "
                        f"target {relationship.get('target')!r}"
                    )

            for rel_id, relationship in relationships.items():
                if (
                    relationship.get("mode") != "External"
                    and relationship.get("type", "").endswith("/image")
                    and rel_id not in media_by_rel_id
                ):
                    errors.append(
                        f"{context}.rels_xml internal image {rel_id!r} has no "
                        "captured media payload"
                    )

            if part_root is not None:
                for rel_id in sorted(
                    _relationship_ids_used_by_part(
                        part_root, f"{context}.xml", errors
                    )
                ):
                    if rel_id not in relationships:
                        errors.append(
                            f"{context}.xml references missing relationship Id {rel_id!r}"
                        )

    page_layout = template_registry.get("page_layout")
    if not isinstance(page_layout, dict):
        return
    sections: List[Tuple[str, Any]] = [
        ("page_layout.default_section", page_layout.get("default_section"))
    ]
    chain = page_layout.get("section_chain")
    if isinstance(chain, list):
        sections.extend(
            (f"page_layout.section_chain[{index}]", section)
            for index, section in enumerate(chain)
        )
    for section_context, section in sections:
        if not isinstance(section, dict):
            continue
        for refs_key, expected_kind in (
            ("header_refs", "header"),
            ("footer_refs", "footer"),
        ):
            refs = section.get(refs_key)
            if refs is None:
                continue
            if not isinstance(refs, dict):
                errors.append(f"{section_context}.{refs_key} must be a dict")
                continue
            for reference_type, rel_id in refs.items():
                if reference_type not in {"default", "even", "first"}:
                    errors.append(
                        f"{section_context}.{refs_key} has unsupported reference type "
                        f"{reference_type!r}"
                    )
                    continue
                if rel_id is None:
                    continue
                if not isinstance(rel_id, str) or not rel_id:
                    errors.append(
                        f"{section_context}.{refs_key}[{reference_type!r}] must be "
                        "a relationship Id or null"
                    )
                    continue
                actual_kind = document_rel_ids.get(rel_id)
                if actual_kind != expected_kind:
                    errors.append(
                        f"{section_context}.{refs_key}[{reference_type!r}] references "
                        f"missing {expected_kind} relationship Id {rel_id!r}"
                    )


def _validate_page_layout(
    template_registry: Dict[str, Any], errors: List[str]
) -> None:
    """Check 8: page_layout contract required for Phase 2 layout sync."""
    page_layout = template_registry.get("page_layout")
    if not isinstance(page_layout, dict):
        errors.append(
            "arch_template_registry.json is missing page_layout, but Phase 2 page layout sync requires it."
        )
        return

    default_section = page_layout.get("default_section")
    if not isinstance(default_section, dict):
        errors.append(
            "arch_template_registry.json is missing page_layout.default_section, but Phase 2 page layout sync requires it."
        )
        return

    sectpr = default_section.get("sectPr")
    if not isinstance(sectpr, str) or not sectpr.strip():
        errors.append(
            "arch_template_registry.json is missing page_layout.default_section.sectPr, but Phase 2 page layout sync requires it."
        )
    else:
        err = _check_xml_fragment(sectpr, "w:sectPr")
        if err:
            errors.append(f"page_layout.default_section.sectPr: {err}")

    section_chain = page_layout.get("section_chain")
    if section_chain is None:
        return
    if not isinstance(section_chain, list):
        errors.append(f"page_layout.section_chain must be a list, got {type(section_chain).__name__}")
        return

    for idx, section in enumerate(section_chain):
        if not isinstance(section, dict):
            errors.append(f"page_layout.section_chain[{idx}] must be a dict")
            continue
        chain_sectpr = section.get("sectPr")
        if not chain_sectpr:
            continue
        if not isinstance(chain_sectpr, str):
            errors.append(f"page_layout.section_chain[{idx}].sectPr must be a string")
            continue
        err = _check_xml_fragment(chain_sectpr, "w:sectPr")
        if err:
            errors.append(f"page_layout.section_chain[{idx}].sectPr: {err}")


def _validate_style_cross_ref(
    style_registry: Dict[str, str],
    known_style_ids: Set[str],
    errors: List[str],
) -> None:
    """Check 6: every role style exists in the validated source/portable inventory."""
    for role, sid in sorted(style_registry.items()):
        if sid not in known_style_ids:
            errors.append(
                f"Style ID '{sid}' (mapped from role '{role}') in "
                "arch_style_registry.json not found in the validated style inventory"
            )


def _validate_numbering_consistency(
    template_registry: Dict[str, Any], errors: List[str]
) -> None:
    """Check 7: numbering num -> abstractNum references are internally consistent."""
    numbering = template_registry.get("numbering")
    if not isinstance(numbering, dict):
        return

    # Collect known abstractNumIds
    abstract_nums = numbering.get("abstract_nums", [])
    if not isinstance(abstract_nums, list):
        errors.append(
            f"numbering.abstract_nums must be a list, got {type(abstract_nums).__name__}"
        )
        return

    abstract_ids: Set[int] = set()
    for idx, an in enumerate(abstract_nums):
        if not isinstance(an, dict):
            errors.append(f"numbering.abstract_nums[{idx}] must be a dict")
            continue
        aid = an.get("abstractNumId")
        if not isinstance(aid, int):
            errors.append(
                f"numbering.abstract_nums[{idx}] missing or non-integer 'abstractNumId'"
            )
            continue
        abstract_ids.add(aid)

    # Validate num references
    nums = numbering.get("nums", [])
    if not isinstance(nums, list):
        errors.append(f"numbering.nums must be a list, got {type(nums).__name__}")
        return

    for idx, num in enumerate(nums):
        if not isinstance(num, dict):
            errors.append(f"numbering.nums[{idx}] must be a dict")
            continue
        nid = num.get("numId")
        if not isinstance(nid, int):
            errors.append(f"numbering.nums[{idx}] missing or non-integer 'numId'")
            continue
        ref_aid = num.get("abstractNumId")
        if not isinstance(ref_aid, int):
            errors.append(
                f"numbering.nums[{idx}] (numId={nid}) missing or non-integer 'abstractNumId'"
            )
            continue
        if ref_aid not in abstract_ids:
            errors.append(
                f"numbering.nums[{idx}] (numId={nid}) references "
                f"abstractNumId={ref_aid} which is not defined in abstract_nums"
            )


def preflight_validate_registries(
    style_registry: Dict[str, str],
    template_registry: Dict[str, Any],
    *,
    additional_known_style_ids: Optional[Set[str]] = None,
) -> List[str]:
    """
    Validate both Phase 2 contract files before any mutation.

    Runs all checks and collects every error so the caller sees the full
    picture in a single pass.  An empty return list means validation passed.

    Args:
        style_registry:    role -> styleId mapping (from load_arch_style_registry).
        template_registry: full dict from arch_template_registry.json.
        additional_known_style_ids: style IDs validated from a strict bundle's
            portable stylesheet. Legacy callers should leave this unset.

    Returns:
        List of error strings.  Empty list means the contract is valid.
    """
    errors: List[str] = []

    _validate_template_sections(template_registry, errors)
    known_style_ids = _validate_style_defs(template_registry, errors)
    if additional_known_style_ids:
        known_style_ids.update(additional_known_style_ids)
    _validate_compat_xml(template_registry, errors)
    _validate_top_level_xml_fragments(template_registry, errors)
    _validate_embedded_fonts(template_registry, errors)
    _validate_header_footer_contract(template_registry, errors)
    _validate_style_cross_ref(style_registry, known_style_ids, errors)
    _validate_numbering_consistency(template_registry, errors)
    _validate_page_layout(template_registry, errors)

    return errors
