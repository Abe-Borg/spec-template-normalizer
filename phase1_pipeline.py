"""Headless, transactional Phase 1 orchestration.

The GUI is deliberately a thin caller of this module.  One run snapshots the
selected DOCX, analyzes only that immutable snapshot, produces a validated
artifact set, and publishes it as a versioned bundle by atomic directory rename.
The architect document and extracted package are never normalized in place.
"""

from __future__ import annotations

import hashlib
import json
import shutil
import tempfile
import warnings
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Callable, Dict, Optional

from arch_env_extractor import extract_arch_template_registry
from docx_decomposer import (
    build_portable_styles_xml,
    build_slim_bundle,
    build_style_registry_dict,
    extract_docx,
    validate_instructions,
)
from llm_classifier import classify_document, compute_coverage
from paragraph_rules import is_classifiable_paragraph
from phase1_bundle import (
    BundleArtifacts,
    ProducerIdentity,
    discard_staged_bundle,
    publish_staged_bundle,
    source_identity,
    stage_phase1_bundle,
    write_classification_audit,
)
from phase1_validator import validate_style_registry, validate_template_registry


PIPELINE_VERSION = "2.0.0"
DEFAULT_MODEL = "claude-opus-4-6"
ProgressCallback = Callable[[str], None]
Classifier = Callable[..., Dict[str, Any]]


@dataclass(frozen=True)
class Phase1Result:
    bundle_dir: Path
    manifest_path: Path
    source_sha256: str
    handled_paragraphs: int
    classifiable_paragraphs: int

    @property
    def coverage(self) -> float:
        if self.classifiable_paragraphs == 0:
            return 1.0
        return self.handled_paragraphs / self.classifiable_paragraphs


def _emit(progress: Optional[ProgressCallback], message: str) -> None:
    if progress is not None:
        try:
            progress(message)
        except Exception as exc:
            warnings.warn(
                f"Phase 1 progress callback failed and was ignored: {exc}",
                RuntimeWarning,
                stacklevel=2,
            )


def _sha256_text(value: str) -> str:
    return hashlib.sha256(value.encode("utf-8")).hexdigest()


def _snapshot_source(source: Path, destination: Path) -> None:
    """Copy one stable source version and reject a concurrent source change."""
    before = source.stat()
    destination.parent.mkdir(parents=True, exist_ok=True)
    with source.open("rb") as reader, destination.open("xb") as writer:
        shutil.copyfileobj(reader, writer, length=1024 * 1024)
    after = source.stat()
    copied_digest = hashlib.sha256()
    with destination.open("rb") as reader:
        while chunk := reader.read(1024 * 1024):
            copied_digest.update(chunk)
    source_digest = hashlib.sha256()
    with source.open("rb") as reader:
        while chunk := reader.read(1024 * 1024):
            source_digest.update(chunk)
    final = source.stat()
    if (
        (before.st_size, before.st_mtime_ns) != (after.st_size, after.st_mtime_ns)
        or (after.st_size, after.st_mtime_ns) != (final.st_size, final.st_mtime_ns)
        or copied_digest.digest() != source_digest.digest()
    ):
        raise RuntimeError(
            "The selected DOCX changed while it was being snapshotted. Close/save it and run again."
        )


def run_phase1(
    source_docx: Path,
    output_root: Path,
    api_key: str,
    *,
    model: str = DEFAULT_MODEL,
    prompt_dir: Optional[Path] = None,
    progress: Optional[ProgressCallback] = None,
    classifier: Optional[Classifier] = None,
) -> Phase1Result:
    """Analyze *source_docx* and atomically publish one validated Phase 1 bundle."""
    source_docx = Path(source_docx)
    output_root = Path(output_root)
    if not source_docx.is_file():
        raise FileNotFoundError(f"Source DOCX does not exist: {source_docx}")
    if source_docx.suffix.lower() != ".docx":
        raise ValueError(f"Source must be a .docx file: {source_docx}")
    if classifier is None and (not isinstance(api_key, str) or not api_key.strip()):
        raise ValueError("An Anthropic API key is required")
    if not isinstance(model, str) or not model.strip():
        raise ValueError("Classifier model must be a non-empty string")

    output_root.mkdir(parents=True, exist_ok=True)
    if not output_root.is_dir():
        raise NotADirectoryError(f"Output root is not a directory: {output_root}")

    prompt_dir = Path(prompt_dir) if prompt_dir is not None else Path(__file__).resolve().parent
    master_prompt = (prompt_dir / "master_prompt.txt").read_text(encoding="utf-8")
    run_instruction = (prompt_dir / "run_instruction_prompt.txt").read_text(encoding="utf-8")
    classifier_is_injected = classifier is not None
    classify = classifier or classify_document

    work_dir = Path(tempfile.mkdtemp(prefix=".phase1-work-", dir=str(output_root)))
    try:
        snapshot_path = work_dir / "source" / source_docx.name
        _emit(progress, f"Snapshotting {source_docx.name}...")
        _snapshot_source(source_docx, snapshot_path)
        identity = source_identity(snapshot_path)

        extract_dir = work_dir / "extracted"
        _emit(progress, "Safely unpacking the DOCX package...")
        extract_docx(snapshot_path, extract_dir)

        artifact_dir = work_dir / "artifacts"
        artifact_dir.mkdir()
        source_styles_path = artifact_dir / "source_styles.xml"
        source_styles_path.write_bytes((extract_dir / "word" / "styles.xml").read_bytes())
        source_settings_part = extract_dir / "word" / "settings.xml"
        source_settings_path: Optional[Path] = None
        if source_settings_part.is_file():
            source_settings_path = artifact_dir / "source_settings.xml"
            source_settings_path.write_bytes(source_settings_part.read_bytes())

        _emit(progress, "Reading paragraph, style, and numbering structure...")
        slim_bundle = build_slim_bundle(extract_dir)
        _emit(progress, f"Classifying {len(slim_bundle.get('paragraphs', []))} paragraphs...")
        instructions = classify(
            slim_bundle=slim_bundle,
            master_prompt=master_prompt,
            run_instruction=run_instruction,
            api_key=api_key,
            model=model,
        )
        validate_instructions(instructions, slim_bundle=slim_bundle)
        coverage, handled, classifiable = compute_coverage(slim_bundle, instructions)
        if coverage != 1.0:
            raise ValueError(
                f"Classification coverage must be 100%; handled {handled}/{classifiable} paragraphs"
            )

        _emit(progress, "Deriving portable styles without changing the source package...")
        portable_styles_path = artifact_dir / "portable_styles.xml"
        portable_styles_path.write_text(
            build_portable_styles_xml(extract_dir, instructions),
            encoding="utf-8",
        )

        _emit(progress, "Capturing the source formatting environment...")
        style_registry = build_style_registry_dict(
            extract_dir,
            identity.filename,
            instructions,
            pre_apply_bundle=slim_bundle,
            styles_xml_path=portable_styles_path,
            source_sha256=identity.sha256,
        )
        template_registry = extract_arch_template_registry(extract_dir, snapshot_path)
        validate_style_registry(style_registry)
        validate_template_registry(template_registry)

        style_registry_path = artifact_dir / "arch_style_registry.json"
        style_registry_path.write_text(
            json.dumps(style_registry, indent=2, sort_keys=True, ensure_ascii=False) + "\n",
            encoding="utf-8",
        )
        template_registry_path = artifact_dir / "arch_template_registry.json"
        template_registry_path.write_text(
            json.dumps(template_registry, indent=2, sort_keys=True, ensure_ascii=False) + "\n",
            encoding="utf-8",
        )

        expected_indices = [
            int(paragraph["paragraph_index"])
            for paragraph in slim_bundle.get("paragraphs", [])
            if is_classifiable_paragraph(paragraph)
        ]
        classification_audit_path = write_classification_audit(
            artifact_dir / "classification_audit.json",
            instructions,
            slim_bundle,
            identity.sha256,
            expected_indices,
        )

        producer = ProducerIdentity(
            name="spec-template-normalizer",
            version=PIPELINE_VERSION,
            classifier_provider="injected" if classifier_is_injected else "anthropic",
            classifier_model=(
                getattr(classify, "__name__", "custom-classifier")
                if classifier_is_injected
                else model
            ),
            master_prompt_sha256=_sha256_text(master_prompt),
            run_instruction_sha256=_sha256_text(run_instruction),
        )
        artifacts = BundleArtifacts(
            style_registry=style_registry_path,
            template_registry=template_registry_path,
            classification_audit=classification_audit_path,
            source_styles=source_styles_path,
            portable_styles=portable_styles_path,
            source_settings=source_settings_path,
        )

        _emit(progress, "Validating checksums and publishing the bundle...")
        staged = stage_phase1_bundle(
            output_root=output_root,
            source_docx=snapshot_path,
            artifacts=artifacts,
            producer=producer,
        )
        try:
            bundle_dir = publish_staged_bundle(staged)
        except Exception:
            if staged.path.exists():
                try:
                    discard_staged_bundle(staged)
                except OSError as cleanup_exc:
                    warnings.warn(
                        f"Could not remove failed staging directory {staged.path}: {cleanup_exc}",
                        RuntimeWarning,
                        stacklevel=2,
                    )
            raise
        _emit(progress, f"Published validated bundle: {bundle_dir.name}")
        return Phase1Result(
            bundle_dir=bundle_dir,
            manifest_path=bundle_dir / "phase1_bundle_manifest.json",
            source_sha256=identity.sha256,
            handled_paragraphs=handled,
            classifiable_paragraphs=classifiable,
        )
    finally:
        try:
            shutil.rmtree(work_dir)
        except OSError as exc:
            _emit(progress, f"WARNING: private work directory remains at {work_dir}: {exc}")
            warnings.warn(
                f"Phase 1 could not remove private work directory {work_dir}: {exc}",
                RuntimeWarning,
                stacklevel=2,
            )
