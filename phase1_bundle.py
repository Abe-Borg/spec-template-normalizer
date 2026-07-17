"""Create, publish, and verify immutable Phase 1 artifact bundles.

The public workflow is intentionally two-step::

    staged = stage_phase1_bundle(...)
    published_dir = publish_staged_bundle(staged, overwrite=False)

Staging copies every artifact into a private sibling directory, records its
size and SHA-256 digest, verifies source identity across the two registries,
and writes a versioned manifest.  Publication is a same-filesystem directory
rename, so consumers never observe a partially copied new bundle.
"""

from __future__ import annotations

import hashlib
import json
import os
import re
import shutil
import tempfile
import uuid
import xml.etree.ElementTree as ET
import zipfile
from dataclasses import dataclass, field
from datetime import datetime, timezone
from pathlib import Path
from typing import Any, Dict, Mapping, Optional, Sequence, Set, Tuple, Union

from phase1_validator import (
    validate_instruction_contract,
    validate_style_registry,
    validate_template_registry,
)


BUNDLE_FORMAT = "spec-template-normalizer.phase1"
MANIFEST_VERSION = 1
MANIFEST_FILENAME = "phase1_bundle_manifest.json"

STYLE_REGISTRY_FILENAME = "arch_style_registry.json"
TEMPLATE_REGISTRY_FILENAME = "arch_template_registry.json"
SOURCE_STYLES_FILENAME = "source_styles.xml"
PORTABLE_STYLES_FILENAME = "portable_styles.xml"
SOURCE_SETTINGS_FILENAME = "source_settings.xml"
MAX_MANIFEST_BYTES = 2 * 1024 * 1024
MAX_ARTIFACT_BYTES = 256 * 1024 * 1024
CLASSIFICATION_AUDIT_FILENAME = "classification_audit.json"
CLASSIFICATION_AUDIT_VERSION = 1

REQUIRED_ARTIFACT_IDS: Tuple[str, ...] = (
    "style_registry",
    "template_registry",
    "classification_audit",
    "source_styles",
    "portable_styles",
)

_SHA256_RE = re.compile(r"^[0-9a-f]{64}$")
_STAGING_PREFIX = ".phase1-bundle-staging-"
_W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"


@dataclass(frozen=True)
class ProducerIdentity:
    """Identity and reproducibility metadata for the creating process."""

    name: str
    version: str
    run_id: str = field(default_factory=lambda: uuid.uuid4().hex)
    classifier_provider: Optional[str] = None
    classifier_model: Optional[str] = None
    master_prompt_sha256: Optional[str] = None
    run_instruction_sha256: Optional[str] = None

    def __post_init__(self) -> None:
        for field_name in ("name", "version", "run_id"):
            value = getattr(self, field_name)
            if not isinstance(value, str) or not value.strip():
                raise ValueError(f"producer {field_name} must be a non-empty string")
        classifier_values = (self.classifier_provider, self.classifier_model)
        if any(value is not None for value in classifier_values) and not all(
            isinstance(value, str) and value.strip() for value in classifier_values
        ):
            raise ValueError(
                "classifier_provider and classifier_model must either both be "
                "non-empty strings or both be omitted"
            )
        for field_name in ("master_prompt_sha256", "run_instruction_sha256"):
            value = getattr(self, field_name)
            if value is not None and not _is_sha256(value):
                raise ValueError(f"producer {field_name} must be a lowercase SHA-256 digest")

    def to_dict(self) -> Dict[str, Any]:
        result: Dict[str, Any] = {
            "name": self.name,
            "version": self.version,
            "run_id": self.run_id,
        }
        if self.classifier_provider is not None:
            result["classifier"] = {
                "provider": self.classifier_provider,
                "model": self.classifier_model,
            }
        prompt_hashes: Dict[str, str] = {}
        if self.master_prompt_sha256 is not None:
            prompt_hashes["master_prompt_sha256"] = self.master_prompt_sha256
        if self.run_instruction_sha256 is not None:
            prompt_hashes["run_instruction_sha256"] = self.run_instruction_sha256
        if prompt_hashes:
            result["prompts"] = prompt_hashes
        return result


@dataclass(frozen=True)
class SourceIdentity:
    filename: str
    sha256: str
    size_bytes: int

    def to_dict(self) -> Dict[str, Any]:
        return {
            "filename": self.filename,
            "sha256": self.sha256,
            "size_bytes": self.size_bytes,
        }


@dataclass(frozen=True)
class BundleArtifacts:
    """Files that form one Phase 1 bundle.

    ``source_styles`` and ``source_settings`` must be captured before Phase 1
    mutates the extraction tree. ``portable_styles`` is the generated,
    post-normalization stylesheet used by Phase 2.
    """

    style_registry: Path
    template_registry: Path
    classification_audit: Path
    source_styles: Path
    portable_styles: Path
    source_settings: Optional[Path] = None

    def as_sources(self) -> Dict[str, Path]:
        result = {
            "style_registry": Path(self.style_registry),
            "template_registry": Path(self.template_registry),
            "classification_audit": Path(self.classification_audit),
            "source_styles": Path(self.source_styles),
            "portable_styles": Path(self.portable_styles),
        }
        if self.source_settings is not None:
            result["source_settings"] = Path(self.source_settings)
        return result


@dataclass(frozen=True)
class ArtifactRecord:
    artifact_id: str
    path: str
    media_type: str
    sha256: str
    size_bytes: int
    required: bool
    source_kind: str

    def to_dict(self) -> Dict[str, Any]:
        return {
            "artifact_id": self.artifact_id,
            "path": self.path,
            "media_type": self.media_type,
            "sha256": self.sha256,
            "size_bytes": self.size_bytes,
            "required": self.required,
            "source_kind": self.source_kind,
        }


@dataclass(frozen=True)
class BundleManifest:
    bundle_format: str
    manifest_version: int
    bundle_id: str
    created_utc: str
    producer: Mapping[str, Any]
    source: SourceIdentity
    required_artifacts: Tuple[str, ...]
    artifacts: Tuple[ArtifactRecord, ...]

    def to_dict(self) -> Dict[str, Any]:
        return {
            "bundle_format": self.bundle_format,
            "manifest_version": self.manifest_version,
            "bundle_id": self.bundle_id,
            "created_utc": self.created_utc,
            "producer": dict(self.producer),
            "source": self.source.to_dict(),
            "required_artifacts": list(self.required_artifacts),
            "artifacts": [artifact.to_dict() for artifact in self.artifacts],
        }


@dataclass(frozen=True)
class StagedBundle:
    path: Path
    output_root: Path
    directory_name: str
    manifest: BundleManifest


_ARTIFACT_SPECS: Mapping[str, Tuple[str, str, bool, str]] = {
    "style_registry": (STYLE_REGISTRY_FILENAME, "application/json", True, "generated"),
    "template_registry": (TEMPLATE_REGISTRY_FILENAME, "application/json", True, "generated"),
    "classification_audit": (CLASSIFICATION_AUDIT_FILENAME, "application/json", True, "generated"),
    "source_styles": (SOURCE_STYLES_FILENAME, "application/xml", True, "source_exact"),
    "portable_styles": (PORTABLE_STYLES_FILENAME, "application/xml", True, "generated"),
    "source_settings": (SOURCE_SETTINGS_FILENAME, "application/xml", False, "source_exact"),
}


def sha256_file(path: Path) -> str:
    """Hash a file without loading an arbitrary artifact fully into memory."""
    digest = hashlib.sha256()
    with Path(path).open("rb") as stream:
        for chunk in iter(lambda: stream.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def source_identity(source_docx: Path) -> SourceIdentity:
    source_docx = Path(source_docx)
    if not source_docx.is_file():
        raise FileNotFoundError(f"Source DOCX does not exist: {source_docx}")
    return SourceIdentity(
        filename=source_docx.name,
        sha256=sha256_file(source_docx),
        size_bytes=source_docx.stat().st_size,
    )


def bundle_directory_name(source: SourceIdentity, run_id: str) -> str:
    """Return a filesystem-safe, run-unique directory name."""
    stem = Path(source.filename).stem
    safe_stem = re.sub(r"[^A-Za-z0-9._-]+", "-", stem).strip(".-") or "template"
    safe_stem = safe_stem[:80]
    run_token = hashlib.sha256(run_id.encode("utf-8")).hexdigest()[:12]
    return f"{safe_stem}--{source.sha256[:12]}--{run_token}.phase1"


def build_classification_audit(
    instructions: Mapping[str, Any],
    slim_bundle: Mapping[str, Any],
    source_sha256: str,
    expected_paragraph_indices: Sequence[int],
) -> Dict[str, Any]:
    """Build a deterministic audit of every slim paragraph and its disposition."""
    if not _is_sha256(source_sha256):
        raise ValueError("source_sha256 must be a lowercase SHA-256 digest")
    if not isinstance(slim_bundle, Mapping):
        raise ValueError("slim_bundle must be an object")

    expected = tuple(expected_paragraph_indices)
    validate_instruction_contract(dict(instructions), expected)
    # JSON round-tripping proves the instructions are publishable JSON and
    # freezes a content-equivalent copy that callers cannot later mutate.
    try:
        instructions_copy = json.loads(
            json.dumps(instructions, ensure_ascii=False, allow_nan=False)
        )
    except (TypeError, ValueError) as exc:
        raise ValueError(f"instructions are not JSON serializable: {exc}") from exc

    applied = {
        item["paragraph_index"]: item["styleId"]
        for item in instructions_copy.get("apply_pStyle", [])
    }
    ignored = {
        item["paragraph_index"]: item["reason"]
        for item in instructions_copy.get("ignored_paragraphs", [])
    }
    raw_paragraphs = slim_bundle.get("paragraphs")
    if not isinstance(raw_paragraphs, list):
        raise ValueError("slim_bundle.paragraphs must be an array")

    paragraph_records = []
    seen_indices = set()
    for position, paragraph in enumerate(raw_paragraphs):
        if not isinstance(paragraph, Mapping):
            raise ValueError(f"slim_bundle.paragraphs[{position}] must be an object")
        index = paragraph.get("paragraph_index")
        if type(index) is not int or index < 0:
            raise ValueError(
                f"slim_bundle.paragraphs[{position}].paragraph_index must be a non-negative int"
            )
        if index in seen_indices:
            raise ValueError(f"Duplicate slim paragraph_index: {index}")
        seen_indices.add(index)
        text_value = paragraph.get("text", "")
        if not isinstance(text_value, str):
            raise ValueError(f"slim_bundle.paragraphs[{position}].text must be a string")
        skip_reason = paragraph.get("skip_reason")
        if skip_reason is not None and not isinstance(skip_reason, str):
            raise ValueError(
                f"slim_bundle.paragraphs[{position}].skip_reason must be a string or null"
            )
        truncated = paragraph.get("text_was_truncated", False)
        if not isinstance(truncated, bool):
            raise ValueError(
                f"slim_bundle.paragraphs[{position}].text_was_truncated must be boolean"
            )
        record: Dict[str, Any] = {
            "paragraph_index": index,
            "text": text_value,
            "slim_text_sha256": hashlib.sha256(text_value.encode("utf-8")).hexdigest(),
            "text_was_truncated": truncated,
            "skip_reason": skip_reason,
        }
        if index in applied:
            record["classification"] = "styled"
            record["style_id"] = applied[index]
        elif index in ignored:
            record["classification"] = "ignored"
            record["ignore_reason"] = ignored[index]
        else:
            record["classification"] = "out_of_scope"
        paragraph_records.append(record)

    missing_paragraphs = sorted(set(expected) - seen_indices)
    if missing_paragraphs:
        raise ValueError(
            "expected_paragraph_indices contains indices absent from slim_bundle: "
            f"{missing_paragraphs[:20]}"
        )
    paragraph_records.sort(key=lambda item: item["paragraph_index"])
    expected_sorted = sorted(expected)
    return {
        "audit_version": CLASSIFICATION_AUDIT_VERSION,
        "source_sha256": source_sha256,
        "instructions_sha256": _canonical_json_sha256(instructions_copy),
        "paragraphs_sha256": _canonical_json_sha256(paragraph_records),
        "expected_paragraph_indices": expected_sorted,
        "instructions": instructions_copy,
        "paragraphs": paragraph_records,
    }


def write_classification_audit(
    path: Path,
    instructions: Mapping[str, Any],
    slim_bundle: Mapping[str, Any],
    source_sha256: str,
    expected_paragraph_indices: Sequence[int],
) -> Path:
    """Build, validate, and write ``classification_audit.json``."""
    audit = build_classification_audit(
        instructions,
        slim_bundle,
        source_sha256,
        expected_paragraph_indices,
    )
    destination = Path(path)
    destination.write_text(
        json.dumps(audit, indent=2, sort_keys=True, ensure_ascii=False) + "\n",
        encoding="utf-8",
    )
    return destination


def validate_classification_audit(
    audit_or_path: Union[Mapping[str, Any], Path],
    *,
    expected_source_sha256: Optional[str] = None,
) -> Dict[str, Any]:
    """Strictly validate a classification audit and return a detached copy."""
    if isinstance(audit_or_path, Mapping):
        raw: Any = dict(audit_or_path)
    else:
        raw = _read_json_object(Path(audit_or_path), "classification audit")
    if not isinstance(raw, dict):
        raise ValueError("Classification audit must be a JSON object")
    top_keys = {
        "audit_version",
        "source_sha256",
        "instructions_sha256",
        "paragraphs_sha256",
        "expected_paragraph_indices",
        "instructions",
        "paragraphs",
    }
    _require_exact_keys(raw, top_keys, "classification audit")
    if type(raw["audit_version"]) is not int or raw["audit_version"] != CLASSIFICATION_AUDIT_VERSION:
        raise ValueError(f"Unsupported classification audit version: {raw['audit_version']!r}")
    if not _is_sha256(raw["source_sha256"]):
        raise ValueError("classification audit source_sha256 must be a lowercase SHA-256 digest")
    if expected_source_sha256 is not None and raw["source_sha256"] != expected_source_sha256:
        raise ValueError(
            "Classification audit source_sha256 does not match bundle source SHA-256"
        )
    for hash_field in ("instructions_sha256", "paragraphs_sha256"):
        if not _is_sha256(raw[hash_field]):
            raise ValueError(f"classification audit {hash_field} must be a SHA-256 digest")

    expected = raw["expected_paragraph_indices"]
    if not isinstance(expected, list):
        raise ValueError("classification audit expected_paragraph_indices must be an array")
    for index in expected:
        if type(index) is not int or index < 0:
            raise ValueError(
                "classification audit expected_paragraph_indices must contain non-negative ints"
            )
    if expected != sorted(expected) or len(expected) != len(set(expected)):
        raise ValueError(
            "classification audit expected_paragraph_indices must be unique and sorted"
        )

    instructions = raw["instructions"]
    validate_instruction_contract(instructions, expected)
    if _canonical_json_sha256(instructions) != raw["instructions_sha256"]:
        raise ValueError("classification audit instructions_sha256 mismatch")
    applied = {
        item["paragraph_index"]: item["styleId"]
        for item in instructions.get("apply_pStyle", [])
    }
    ignored = {
        item["paragraph_index"]: item["reason"]
        for item in instructions.get("ignored_paragraphs", [])
    }

    paragraphs = raw["paragraphs"]
    if not isinstance(paragraphs, list):
        raise ValueError("classification audit paragraphs must be an array")
    seen = set()
    previous_index = -1
    for position, paragraph in enumerate(paragraphs):
        context = f"classification audit paragraphs[{position}]"
        if not isinstance(paragraph, dict):
            raise ValueError(f"{context} must be an object")
        base_keys = {
            "paragraph_index", "text", "slim_text_sha256",
            "text_was_truncated", "skip_reason", "classification",
        }
        classification = paragraph.get("classification")
        allowed_keys = set(base_keys)
        if classification == "styled":
            allowed_keys.add("style_id")
        elif classification == "ignored":
            allowed_keys.add("ignore_reason")
        elif classification != "out_of_scope":
            raise ValueError(f"{context}.classification is invalid: {classification!r}")
        _require_exact_keys(paragraph, allowed_keys, context)
        index = paragraph["paragraph_index"]
        if type(index) is not int or index < 0:
            raise ValueError(f"{context}.paragraph_index must be a non-negative int")
        if index in seen or index <= previous_index:
            raise ValueError("classification audit paragraphs must have unique sorted indices")
        seen.add(index)
        previous_index = index
        text_value = paragraph["text"]
        if not isinstance(text_value, str):
            raise ValueError(f"{context}.text must be a string")
        expected_text_hash = hashlib.sha256(text_value.encode("utf-8")).hexdigest()
        if paragraph["slim_text_sha256"] != expected_text_hash:
            raise ValueError(f"{context}.slim_text_sha256 mismatch")
        if not isinstance(paragraph["text_was_truncated"], bool):
            raise ValueError(f"{context}.text_was_truncated must be boolean")
        if paragraph["skip_reason"] is not None and not isinstance(paragraph["skip_reason"], str):
            raise ValueError(f"{context}.skip_reason must be a string or null")
        if classification == "styled":
            if paragraph["style_id"] != applied.get(index):
                raise ValueError(f"{context}.style_id does not match instructions")
        elif classification == "ignored":
            if paragraph["ignore_reason"] != ignored.get(index):
                raise ValueError(f"{context}.ignore_reason does not match instructions")
        elif index in applied or index in ignored:
            raise ValueError(f"{context}.classification does not match instructions")

    missing_expected = sorted(set(expected) - seen)
    if missing_expected:
        raise ValueError(
            "classification audit is missing expected paragraph records: "
            f"{missing_expected[:20]}"
        )
    if _canonical_json_sha256(paragraphs) != raw["paragraphs_sha256"]:
        raise ValueError("classification audit paragraphs_sha256 mismatch")
    # Return a JSON-detached value so later caller mutation cannot affect a
    # previously validated in-memory object.
    return json.loads(json.dumps(raw, ensure_ascii=False, allow_nan=False))


def stage_phase1_bundle(
    output_root: Path,
    source_docx: Path,
    artifacts: BundleArtifacts,
    producer: ProducerIdentity,
    *,
    directory_name: Optional[str] = None,
) -> StagedBundle:
    """Copy and validate a complete bundle in a private staging directory.

    No final bundle directory is created by this function.  Call
    :func:`publish_staged_bundle` after staging succeeds.
    """
    output_root = Path(output_root)
    output_root.mkdir(parents=True, exist_ok=True)
    if not output_root.is_dir():
        raise NotADirectoryError(f"Bundle output root is not a directory: {output_root}")

    source_docx = Path(source_docx)
    identity = source_identity(source_docx)
    final_name = directory_name or bundle_directory_name(identity, producer.run_id)
    _validate_directory_name(final_name)

    sources = artifacts.as_sources()
    _validate_artifact_sources(sources)
    _verify_exact_source_parts(source_docx, sources)
    classification_audit = validate_classification_audit(
        sources["classification_audit"],
        expected_source_sha256=identity.sha256,
    )
    _validate_registry_files(
        sources["style_registry"],
        sources["template_registry"],
        sources["source_styles"],
        sources["portable_styles"],
        identity,
        classification_audit,
        source_settings_present="source_settings" in sources,
    )

    staging_path = Path(tempfile.mkdtemp(prefix=_STAGING_PREFIX, dir=str(output_root)))
    try:
        records = []
        for artifact_id, spec in _ARTIFACT_SPECS.items():
            source_path = sources.get(artifact_id)
            if source_path is None:
                continue
            filename, media_type, required, source_kind = spec
            destination = staging_path / filename
            shutil.copyfile(source_path, destination)
            records.append(
                ArtifactRecord(
                    artifact_id=artifact_id,
                    path=filename,
                    media_type=media_type,
                    sha256=sha256_file(destination),
                    size_bytes=destination.stat().st_size,
                    required=required,
                    source_kind=source_kind,
                )
            )

        # Re-verify the bytes that will actually be published.  This closes
        # the gap where an input path could be replaced after the initial
        # preflight but before copy.
        staged_sources = {
            record.artifact_id: staging_path / record.path for record in records
        }
        _verify_exact_source_parts(source_docx, staged_sources)
        current_identity = source_identity(source_docx)
        if current_identity != identity:
            raise ValueError("Source DOCX changed while the Phase 1 bundle was being staged")
        _validate_registry_files(
            staged_sources["style_registry"],
            staged_sources["template_registry"],
            staged_sources["source_styles"],
            staged_sources["portable_styles"],
            identity,
            validate_classification_audit(
                staged_sources["classification_audit"],
                expected_source_sha256=identity.sha256,
            ),
            source_settings_present="source_settings" in staged_sources,
        )

        now = datetime.now(timezone.utc).isoformat(timespec="seconds").replace("+00:00", "Z")
        bundle_id = f"{identity.sha256[:12]}-{hashlib.sha256(producer.run_id.encode('utf-8')).hexdigest()[:12]}"
        manifest = BundleManifest(
            bundle_format=BUNDLE_FORMAT,
            manifest_version=MANIFEST_VERSION,
            bundle_id=bundle_id,
            created_utc=now,
            producer=producer.to_dict(),
            source=identity,
            required_artifacts=REQUIRED_ARTIFACT_IDS,
            artifacts=tuple(records),
        )
        (staging_path / MANIFEST_FILENAME).write_text(
            json.dumps(manifest.to_dict(), indent=2, sort_keys=True) + "\n",
            encoding="utf-8",
        )

        verified = validate_bundle_directory(
            staging_path,
            expected_source_sha256=identity.sha256,
            reject_unlisted=True,
        )
        return StagedBundle(
            path=staging_path,
            output_root=output_root.resolve(),
            directory_name=final_name,
            manifest=verified,
        )
    except Exception:
        shutil.rmtree(staging_path, ignore_errors=True)
        raise


def publish_staged_bundle(staged: StagedBundle, *, overwrite: bool = False) -> Path:
    """Publish a staged bundle by same-filesystem rename.

    Existing bundles are never replaced unless ``overwrite=True`` is explicit.
    On replacement, the old directory is first moved to a private backup and
    restored if publication fails.
    """
    staging_path = Path(staged.path).resolve()
    output_root = Path(staged.output_root).resolve()
    if staging_path.parent != output_root or not staging_path.name.startswith(_STAGING_PREFIX):
        raise ValueError("Staged bundle path is not a managed child of its output root")
    validate_bundle_directory(staging_path, expected_source_sha256=staged.manifest.source.sha256)

    _validate_directory_name(staged.directory_name)
    destination = output_root / staged.directory_name
    if destination.is_symlink() or (destination.exists() and not destination.is_dir()):
        raise ValueError(
            f"Bundle destination exists but is not a replaceable directory: {destination}"
        )
    if destination.exists() and not overwrite:
        raise FileExistsError(
            f"Bundle already exists: {destination}. Pass overwrite=True to replace it explicitly."
        )

    if not destination.exists():
        os.rename(staging_path, destination)
        return destination

    backup = output_root / f".{staged.directory_name}.replaced-{uuid.uuid4().hex}"
    os.rename(destination, backup)
    try:
        os.rename(staging_path, destination)
    except Exception:
        os.rename(backup, destination)
        raise
    try:
        shutil.rmtree(backup)
    except OSError as exc:
        # The new bundle is already live and valid; do not report the committed
        # publication as failed merely because the recoverable backup remains.
        import warnings
        warnings.warn(
            f"Bundle published to {destination}, but old backup cleanup failed: {backup}: {exc}",
            RuntimeWarning,
            stacklevel=2,
        )
    return destination


def publish_phase1_bundle(
    output_root: Path,
    source_docx: Path,
    artifacts: BundleArtifacts,
    producer: ProducerIdentity,
    *,
    directory_name: Optional[str] = None,
    overwrite: bool = False,
) -> Path:
    """Stage and publish a bundle, cleaning staging residue on failure."""
    staged = stage_phase1_bundle(
        output_root,
        source_docx,
        artifacts,
        producer,
        directory_name=directory_name,
    )
    try:
        return publish_staged_bundle(staged, overwrite=overwrite)
    except Exception:
        if staged.path.exists():
            discard_staged_bundle(staged)
        raise


def discard_staged_bundle(staged: StagedBundle) -> None:
    """Delete an unpublished staging directory created by this module."""
    staging_path = Path(staged.path).resolve()
    output_root = Path(staged.output_root).resolve()
    if staging_path.parent != output_root or not staging_path.name.startswith(_STAGING_PREFIX):
        raise ValueError("Refusing to delete a path that is not a managed staging directory")
    if staging_path.exists():
        shutil.rmtree(staging_path)


def load_bundle_manifest(bundle_dir: Path) -> BundleManifest:
    """Load and strictly validate a v1 manifest model."""
    manifest_path = Path(bundle_dir) / MANIFEST_FILENAME
    if manifest_path.is_symlink():
        raise ValueError(f"Bundle manifest must not be a symlink: {manifest_path}")
    if manifest_path.is_file() and manifest_path.stat().st_size > MAX_MANIFEST_BYTES:
        raise ValueError(f"Bundle manifest exceeds {MAX_MANIFEST_BYTES} bytes")
    try:
        raw = json.loads(manifest_path.read_text(encoding="utf-8"))
    except FileNotFoundError:
        raise ValueError(f"Bundle manifest is missing: {manifest_path}") from None
    except (UnicodeDecodeError, json.JSONDecodeError) as exc:
        raise ValueError(f"Bundle manifest is not valid UTF-8 JSON: {manifest_path}: {exc}") from exc
    return _manifest_from_dict(raw)


def validate_bundle_directory(
    bundle_dir: Path,
    *,
    expected_source_sha256: Optional[str] = None,
    reject_unlisted: bool = True,
) -> BundleManifest:
    """Verify manifest shape, artifact set, hashes, and cross-file identity."""
    bundle_dir = Path(bundle_dir)
    if bundle_dir.is_symlink() or not bundle_dir.is_dir():
        raise ValueError(f"Bundle directory does not exist: {bundle_dir}")
    manifest = load_bundle_manifest(bundle_dir)

    if expected_source_sha256 is not None:
        if not _is_sha256(expected_source_sha256):
            raise ValueError("expected_source_sha256 must be a lowercase SHA-256 digest")
        if manifest.source.sha256 != expected_source_sha256:
            raise ValueError(
                "Bundle source SHA-256 does not match the expected source; "
                f"manifest={manifest.source.sha256}, expected={expected_source_sha256}"
            )

    declared_names = {MANIFEST_FILENAME}
    records_by_id = {record.artifact_id: record for record in manifest.artifacts}
    for record in manifest.artifacts:
        artifact_path = bundle_dir / record.path
        declared_names.add(record.path)
        if artifact_path.is_symlink() or not artifact_path.is_file():
            raise ValueError(f"Bundle artifact is missing or not a regular file: {record.path}")
        actual_size = artifact_path.stat().st_size
        if record.size_bytes > MAX_ARTIFACT_BYTES:
            raise ValueError(
                f"Bundle artifact exceeds safe size limit for {record.path}: "
                f"{record.size_bytes} > {MAX_ARTIFACT_BYTES}"
            )
        if actual_size != record.size_bytes:
            raise ValueError(
                f"Bundle artifact size mismatch for {record.path}: "
                f"manifest={record.size_bytes}, actual={actual_size}"
            )
        actual_sha = sha256_file(artifact_path)
        if actual_sha != record.sha256:
            raise ValueError(
                f"Bundle artifact SHA-256 mismatch for {record.path}: "
                f"manifest={record.sha256}, actual={actual_sha}"
            )

    missing_required = set(REQUIRED_ARTIFACT_IDS) - set(records_by_id)
    if missing_required:
        raise ValueError(f"Bundle is missing required artifacts: {sorted(missing_required)}")

    if reject_unlisted:
        actual_names = {entry.name for entry in bundle_dir.iterdir()}
        extras = sorted(actual_names - declared_names)
        if extras:
            raise ValueError(f"Bundle contains unlisted/stale artifacts: {extras}")

    _validate_registry_files(
        bundle_dir / records_by_id["style_registry"].path,
        bundle_dir / records_by_id["template_registry"].path,
        bundle_dir / records_by_id["source_styles"].path,
        bundle_dir / records_by_id["portable_styles"].path,
        manifest.source,
        validate_classification_audit(
            bundle_dir / records_by_id["classification_audit"].path,
            expected_source_sha256=manifest.source.sha256,
        ),
        source_settings_present="source_settings" in records_by_id,
    )
    _parse_xml_file(bundle_dir / records_by_id["source_styles"].path, "source_styles")
    if "source_settings" in records_by_id:
        _parse_xml_file(bundle_dir / records_by_id["source_settings"].path, "source_settings")
    return manifest


def _manifest_from_dict(raw: Any) -> BundleManifest:
    if not isinstance(raw, dict):
        raise ValueError("Bundle manifest must be a JSON object")
    expected_keys = {
        "bundle_format", "manifest_version", "bundle_id", "created_utc",
        "producer", "source", "required_artifacts", "artifacts",
    }
    _require_exact_keys(raw, expected_keys, "manifest")
    if raw["bundle_format"] != BUNDLE_FORMAT:
        raise ValueError(f"Unsupported bundle_format: {raw['bundle_format']!r}")
    if type(raw["manifest_version"]) is not int or raw["manifest_version"] != MANIFEST_VERSION:
        raise ValueError(f"Unsupported manifest_version: {raw['manifest_version']!r}")
    if not isinstance(raw["bundle_id"], str) or not raw["bundle_id"]:
        raise ValueError("manifest.bundle_id must be a non-empty string")
    _validate_utc_datetime(raw["created_utc"])
    producer = _validate_producer_dict(raw["producer"])
    source = _source_from_dict(raw["source"])

    required = raw["required_artifacts"]
    if not isinstance(required, list) or len(required) != len(set(required)):
        raise ValueError("manifest.required_artifacts must be a unique array")
    if set(required) != set(REQUIRED_ARTIFACT_IDS) or len(required) != len(REQUIRED_ARTIFACT_IDS):
        raise ValueError(
            "manifest.required_artifacts must contain exactly "
            f"{list(REQUIRED_ARTIFACT_IDS)}"
        )

    artifacts_raw = raw["artifacts"]
    if not isinstance(artifacts_raw, list):
        raise ValueError("manifest.artifacts must be an array")
    artifacts = tuple(_artifact_from_dict(item, i) for i, item in enumerate(artifacts_raw))
    ids = [item.artifact_id for item in artifacts]
    paths = [item.path for item in artifacts]
    if len(ids) != len(set(ids)):
        raise ValueError("manifest.artifacts contains duplicate artifact_id values")
    if len(paths) != len(set(paths)):
        raise ValueError("manifest.artifacts contains duplicate paths")
    expected_ids = set(REQUIRED_ARTIFACT_IDS) | ({"source_settings"} if "source_settings" in ids else set())
    if set(ids) != expected_ids:
        raise ValueError(f"manifest.artifacts has an invalid artifact set: {sorted(ids)}")

    return BundleManifest(
        bundle_format=BUNDLE_FORMAT,
        manifest_version=MANIFEST_VERSION,
        bundle_id=raw["bundle_id"],
        created_utc=raw["created_utc"],
        producer=producer,
        source=source,
        required_artifacts=tuple(required),
        artifacts=artifacts,
    )


def _validate_producer_dict(raw: Any) -> Dict[str, Any]:
    if not isinstance(raw, dict):
        raise ValueError("manifest.producer must be an object")
    required = {"name", "version", "run_id"}
    allowed = required | {"classifier", "prompts"}
    _require_allowed_and_required_keys(raw, allowed, required, "manifest.producer")
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
                raise ValueError(f"manifest.producer.classifier.{key} must be a non-empty string")
    if "prompts" in raw:
        prompts = raw["prompts"]
        if not isinstance(prompts, dict):
            raise ValueError("manifest.producer.prompts must be an object")
        allowed_prompts = {"master_prompt_sha256", "run_instruction_sha256"}
        if not prompts or set(prompts) - allowed_prompts:
            raise ValueError("manifest.producer.prompts contains no known prompt hashes")
        for key, value in prompts.items():
            if not _is_sha256(value):
                raise ValueError(f"manifest.producer.prompts.{key} must be a SHA-256 digest")
    return dict(raw)


def _source_from_dict(raw: Any) -> SourceIdentity:
    if not isinstance(raw, dict):
        raise ValueError("manifest.source must be an object")
    _require_exact_keys(raw, {"filename", "sha256", "size_bytes"}, "manifest.source")
    filename = raw["filename"]
    if (
        not isinstance(filename, str)
        or not filename
        or filename in {".", ".."}
        or "/" in filename
        or "\\" in filename
    ):
        raise ValueError("manifest.source.filename must be a non-empty basename")
    if not _is_sha256(raw["sha256"]):
        raise ValueError("manifest.source.sha256 must be a lowercase SHA-256 digest")
    if type(raw["size_bytes"]) is not int or raw["size_bytes"] < 0:
        raise ValueError("manifest.source.size_bytes must be a non-negative integer")
    return SourceIdentity(filename, raw["sha256"], raw["size_bytes"])


def _artifact_from_dict(raw: Any, index: int) -> ArtifactRecord:
    context = f"manifest.artifacts[{index}]"
    if not isinstance(raw, dict):
        raise ValueError(f"{context} must be an object")
    keys = {"artifact_id", "path", "media_type", "sha256", "size_bytes", "required", "source_kind"}
    _require_exact_keys(raw, keys, context)
    artifact_id = raw["artifact_id"]
    if artifact_id not in _ARTIFACT_SPECS:
        raise ValueError(f"{context}.artifact_id is unknown: {artifact_id!r}")
    expected_path, expected_media, expected_required, expected_kind = _ARTIFACT_SPECS[artifact_id]
    expected_values = {
        "path": expected_path,
        "media_type": expected_media,
        "required": expected_required,
        "source_kind": expected_kind,
    }
    for field_name, expected in expected_values.items():
        if raw[field_name] != expected:
            raise ValueError(
                f"{context}.{field_name} must be {expected!r} for {artifact_id}, "
                f"got {raw[field_name]!r}"
            )
    if not _is_sha256(raw["sha256"]):
        raise ValueError(f"{context}.sha256 must be a lowercase SHA-256 digest")
    if type(raw["size_bytes"]) is not int or raw["size_bytes"] < 0:
        raise ValueError(f"{context}.size_bytes must be a non-negative integer")
    return ArtifactRecord(
        artifact_id=artifact_id,
        path=raw["path"],
        media_type=raw["media_type"],
        sha256=raw["sha256"],
        size_bytes=raw["size_bytes"],
        required=raw["required"],
        source_kind=raw["source_kind"],
    )


def _validate_artifact_sources(sources: Mapping[str, Path]) -> None:
    missing = set(REQUIRED_ARTIFACT_IDS) - set(sources)
    if missing:
        raise ValueError(f"Missing required bundle artifact sources: {sorted(missing)}")
    unknown = set(sources) - set(_ARTIFACT_SPECS)
    if unknown:
        raise ValueError(f"Unknown bundle artifact sources: {sorted(unknown)}")
    for artifact_id, path in sources.items():
        if not Path(path).is_file():
            raise FileNotFoundError(f"Artifact source does not exist ({artifact_id}): {path}")


def _verify_exact_source_parts(source_docx: Path, sources: Mapping[str, Path]) -> None:
    try:
        with zipfile.ZipFile(source_docx, "r") as archive:
            try:
                expected_styles = archive.read("word/styles.xml")
            except KeyError:
                raise ValueError("Source DOCX has no word/styles.xml part") from None
            actual_styles = sources["source_styles"].read_bytes()
            if actual_styles != expected_styles:
                raise ValueError(
                    "source_styles is not the exact pre-mutation word/styles.xml from the source DOCX"
                )

            has_settings = "word/settings.xml" in archive.namelist()
            supplied_settings = sources.get("source_settings")
            if has_settings and supplied_settings is None:
                raise ValueError(
                    "Source DOCX contains word/settings.xml; source_settings must be supplied"
                )
            if not has_settings and supplied_settings is not None:
                raise ValueError(
                    "source_settings was supplied, but the source DOCX has no word/settings.xml"
                )
            if has_settings and supplied_settings is not None:
                if supplied_settings.read_bytes() != archive.read("word/settings.xml"):
                    raise ValueError(
                        "source_settings is not the exact pre-mutation word/settings.xml from the source DOCX"
                    )
    except zipfile.BadZipFile as exc:
        raise ValueError(f"Source file is not a valid DOCX ZIP package: {source_docx}") from exc


def _validate_registry_files(
    style_registry_path: Path,
    template_registry_path: Path,
    source_styles_path: Path,
    portable_styles_path: Path,
    source: SourceIdentity,
    classification_audit: Mapping[str, Any],
    *,
    source_settings_present: bool,
) -> None:
    style_registry = _read_json_object(style_registry_path, "style registry")
    template_registry = _read_json_object(template_registry_path, "template registry")
    validate_style_registry(style_registry)
    validate_template_registry(template_registry)

    if style_registry.get("source_docx") != source.filename:
        raise ValueError(
            "Style registry source_docx does not match bundle source filename: "
            f"{style_registry.get('source_docx')!r} != {source.filename!r}"
        )
    if style_registry.get("source_sha256") != source.sha256:
        raise ValueError(
            "Style registry source_sha256 does not match bundle source SHA-256: "
            f"{style_registry.get('source_sha256')!r} != {source.sha256!r}"
        )
    template_source = template_registry.get("meta", {}).get("source_docx", {})
    if template_source.get("filename") != source.filename:
        raise ValueError(
            "Template registry source filename does not match bundle source filename: "
            f"{template_source.get('filename')!r} != {source.filename!r}"
        )
    if template_source.get("sha256") != source.sha256:
        raise ValueError(
            "Template registry source SHA-256 does not match bundle source SHA-256: "
            f"{template_source.get('sha256')!r} != {source.sha256!r}"
        )
    inventory_has_settings = template_registry.get("package_inventory", {}).get("has_settings")
    if not isinstance(inventory_has_settings, bool):
        raise ValueError(
            "Template registry package_inventory.has_settings must be boolean for a bundle"
        )
    if inventory_has_settings != source_settings_present:
        raise ValueError(
            "Template registry settings inventory disagrees with source_settings artifact presence"
        )

    source_ids = _style_ids_from_xml(source_styles_path)
    portable_ids = _style_ids_from_xml(portable_styles_path)
    registry_ids = {
        item.get("style_id")
        for item in template_registry.get("styles", {}).get("style_defs", [])
        if isinstance(item, dict) and item.get("style_id")
    }
    if source_ids != registry_ids:
        missing = sorted(source_ids - registry_ids)
        unexpected = sorted(registry_ids - source_ids)
        raise ValueError(
            "source_styles and template registry style inventories differ; "
            f"missing_from_registry={missing[:20]}, unexpected_in_registry={unexpected[:20]}"
        )
    missing_source_styles = sorted(source_ids - portable_ids)
    if missing_source_styles:
        raise ValueError(
            "portable_styles does not preserve every source style; missing="
            + ", ".join(missing_source_styles[:20])
        )
    role_ids = {
        spec.get("style_id")
        for spec in style_registry.get("roles", {}).values()
        if isinstance(spec, dict) and spec.get("style_id")
    }
    missing_roles = sorted(role_ids - portable_ids)
    if missing_roles:
        raise ValueError(
            "portable_styles is missing role style IDs: " + ", ".join(missing_roles)
        )
    unexpected_portable_styles = sorted(portable_ids - (source_ids | role_ids))
    if unexpected_portable_styles:
        raise ValueError(
            "portable_styles contains styles that are neither source styles nor declared role styles: "
            + ", ".join(unexpected_portable_styles[:20])
        )

    audit_instructions = classification_audit["instructions"]
    audit_roles = audit_instructions.get("roles", {})
    if set(audit_roles) != set(style_registry.get("roles", {})):
        raise ValueError(
            "Classification audit and style registry role sets differ; "
            f"audit={sorted(audit_roles)}, registry={sorted(style_registry.get('roles', {}))}"
        )
    for role, audit_spec in audit_roles.items():
        registry_spec = style_registry["roles"][role]
        if (
            audit_spec.get("styleId") != registry_spec.get("style_id")
            or audit_spec.get("exemplar_paragraph_index")
            != registry_spec.get("exemplar_paragraph_index")
        ):
            raise ValueError(
                f"Classification audit and style registry disagree for role {role}"
            )
    referenced_style_ids = {
        item.get("styleId")
        for item in audit_instructions.get("create_styles", [])
        if isinstance(item, dict) and item.get("styleId")
    } | {
        item.get("styleId")
        for item in audit_instructions.get("apply_pStyle", [])
        if isinstance(item, dict) and item.get("styleId")
    }
    missing_referenced = sorted(referenced_style_ids - portable_ids)
    if missing_referenced:
        raise ValueError(
            "portable_styles is missing styles referenced by the classification audit: "
            + ", ".join(missing_referenced[:20])
        )


def _read_json_object(path: Path, label: str) -> Dict[str, Any]:
    try:
        value = json.loads(Path(path).read_text(encoding="utf-8"))
    except (OSError, UnicodeDecodeError, json.JSONDecodeError) as exc:
        raise ValueError(f"Invalid {label} JSON at {path}: {exc}") from exc
    if not isinstance(value, dict):
        raise ValueError(f"{label.capitalize()} must contain a JSON object")
    return value


def _style_ids_from_xml(path: Path) -> Set[str]:
    root = _parse_xml_file(path, "portable_styles")
    return {
        value
        for style in root.iter(f"{{{_W_NS}}}style")
        if (value := style.attrib.get(f"{{{_W_NS}}}styleId"))
    }


def _parse_xml_file(path: Path, label: str) -> ET.Element:
    try:
        return ET.fromstring(Path(path).read_bytes())
    except (OSError, ET.ParseError) as exc:
        raise ValueError(f"Bundle {label} is not well-formed XML: {path}: {exc}") from exc


def _validate_directory_name(name: str) -> None:
    if (
        not isinstance(name, str)
        or not name
        or name in {".", ".."}
        or "/" in name
        or "\\" in name
        or Path(name).name != name
    ):
        raise ValueError("Bundle directory_name must be a non-empty basename")


def _validate_utc_datetime(value: Any) -> None:
    if not isinstance(value, str) or not value.endswith("Z"):
        raise ValueError("manifest.created_utc must be an RFC 3339 UTC timestamp ending in Z")
    try:
        parsed = datetime.fromisoformat(value[:-1] + "+00:00")
    except ValueError as exc:
        raise ValueError("manifest.created_utc must be a valid RFC 3339 timestamp") from exc
    if parsed.tzinfo is None or parsed.utcoffset() != timezone.utc.utcoffset(parsed):
        raise ValueError("manifest.created_utc must be UTC")


def _require_exact_keys(raw: Mapping[str, Any], expected: Set[str], context: str) -> None:
    _require_allowed_and_required_keys(raw, expected, expected, context)


def _require_allowed_and_required_keys(
    raw: Mapping[str, Any],
    allowed: Set[str],
    required: Set[str],
    context: str,
) -> None:
    missing = required - set(raw)
    extra = set(raw) - allowed
    if missing:
        raise ValueError(f"{context} is missing required keys: {sorted(missing)}")
    if extra:
        raise ValueError(f"{context} contains unknown keys: {sorted(extra)}")


def _is_sha256(value: Any) -> bool:
    return isinstance(value, str) and _SHA256_RE.fullmatch(value) is not None


def _canonical_json_sha256(value: Any) -> str:
    encoded = json.dumps(
        value,
        ensure_ascii=False,
        sort_keys=True,
        separators=(",", ":"),
        allow_nan=False,
    ).encode("utf-8")
    return hashlib.sha256(encoded).hexdigest()
