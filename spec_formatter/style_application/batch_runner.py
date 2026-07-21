"""Shared Phase 2 file pipeline and concurrent batch runner."""

from __future__ import annotations

import hashlib
import json
import re
import tempfile
import time
import zipfile
import os
from concurrent.futures import ThreadPoolExecutor, as_completed
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Callable, Dict, List, Optional

from .arch_env_applier import apply_environment_to_target
from .core.classification import apply_phase2_classifications, build_phase2_slim_bundle
from .core.csi_to_canadian import (
    CSI_TO_CANADIAN,
    FORMAT_ONLY,
    CanadianConversionReport,
    apply_csi_to_canadian,
    classifications_for_canadian_application,
    validate_conversion_mode,
)
from .core.token_utils import extract_target_tokens
from .core.batch_classifier import (
    BatchClassificationError,
    build_batch_requests,
    reassemble_file_classifications,
    submit_and_poll,
)
from .core.llm_classifier import classify_target_document
from .core.ooxml_text import read_xml_text
from .core.registry import (
    PHASE1_MANIFEST_FILENAME,
    build_arch_styles_xml_from_registry,
    load_arch_style_registry,
    load_available_roles_from_registry,
    load_role_specs_from_registry,
    preflight_validate_registries,
    resolve_arch_extract_root,
    validate_phase1_bundle_directory,
)
from .core.stability import snapshot_stability, verify_stability
from .core.style_import import import_arch_styles_into_target
from .docx_decomposer import DocxDecomposer
from .docx_patch import patch_docx
from .header_footer_importer import (
    patch_header_footer_tokens,
    remap_header_footer_numids,
)
from .phase2_invariants import validate_docx_package, verify_phase2_invariants
from .core.style_import import collect_style_dependency_closure

try:
    from .numbering_importer import build_numbering_import_plan, import_numbering

    HAS_NUMBERING_IMPORTER = True
except ImportError:
    HAS_NUMBERING_IMPORTER = False


@dataclass
class BatchResult:
    filename: str
    success: bool
    output_path: Optional[Path]
    log: List[str]
    error: Optional[str]
    duration_seconds: float
    conversion_report: Optional[CanadianConversionReport] = None


@dataclass(frozen=True)
class SharedConfig:
    arch_registry: Dict[str, str]
    env_registry: Dict[str, Any]
    arch_styles_xml: str
    available_roles: List[str]
    source_tokens: Dict[str, str]
    arch_root: Path
    role_specs: Optional[Dict[str, Dict[str, Any]]] = None
    bundle_manifest: Optional[Dict[str, Any]] = None
    legacy_mode: bool = False


@dataclass(frozen=True)
class PreparedFile:
    file_key: str
    docx_path: Path
    extract_dir: Path
    bundle: Dict[str, Any]
    prep_log: List[str]


def _coverage_counts(bundle: Dict[str, Any], classifications: Dict[str, Any]) -> tuple[int, int, int]:
    total = len(bundle.get("paragraphs", [])) + len(bundle.get("deterministic_classifications", []))
    classified = len(classifications.get("classifications", []))
    return classified, total, len(bundle.get("paragraphs", []))


def _check_numbering_module_needed(arch_styles_xml: str, needed_style_ids: List[str]) -> None:
    """Raise if styles need numbering but numbering_importer is unavailable."""
    for sid in collect_style_dependency_closure(arch_styles_xml, needed_style_ids):
        pat = r'<w:style[^>]*w:styleId="' + re.escape(sid) + r'"[^>]*>[\s\S]*?</w:style>'
        m = re.search(pat, arch_styles_xml)
        if m and '<w:numId' in m.group(0):
            raise ImportError(
                "numbering_importer module is not available but imported styles "
                f"require numbering definitions (e.g. style '{sid}'). "
                "Ensure numbering_importer.py is on the Python path."
            )


def load_and_validate_shared_config(
    arch_path: Path,
    *,
    allow_legacy_bundle: bool = False,
) -> SharedConfig:
    requested_path = Path(arch_path)
    candidate_root = requested_path.parent if requested_path.is_file() else requested_path
    manifest_path = candidate_root / PHASE1_MANIFEST_FILENAME

    bundle_manifest: Optional[Dict[str, Any]] = None
    legacy_mode = False
    if manifest_path.exists():
        bundle_manifest, artifact_paths = validate_phase1_bundle_directory(candidate_root)
        arch_root = candidate_root
        style_registry_path = artifact_paths["style_registry"]
        template_registry_path = artifact_paths["template_registry"]
        portable_styles_path = artifact_paths["portable_styles"]
    else:
        if not allow_legacy_bundle:
            raise FileNotFoundError(
                f"Strict Phase 1 bundle required: {manifest_path} was not found. "
                "Regenerate the template with Phase 1, or explicitly call "
                "load_and_validate_shared_config(..., allow_legacy_bundle=True) "
                "for a trusted legacy bundle."
            )
        legacy_mode = True
        arch_root = resolve_arch_extract_root(requested_path)
        style_registry_path = arch_root / "arch_style_registry.json"
        template_registry_path = arch_root / "arch_template_registry.json"
        portable_styles_path = arch_root / "arch_styles_raw.xml"

    arch_registry = load_arch_style_registry(style_registry_path)
    # Legacy registries predate the numbering provenance contract. Passing
    # their partial role records into the strict numbering path turns an
    # explicitly opted-in compatibility mode into a runtime failure.
    role_specs = None if legacy_mode else load_role_specs_from_registry(style_registry_path)
    available_roles = load_available_roles_from_registry(style_registry_path)
    if not available_roles:
        raise ValueError("Could not load architect registry")

    env_registry = json.loads(template_registry_path.read_text(encoding="utf-8"))

    preflight_errors = preflight_validate_registries(
        arch_registry,
        env_registry,
        additional_known_style_ids=(set(arch_registry.values()) if not legacy_mode else None),
    )
    if preflight_errors:
        error_report = "\n".join(f"  - {e}" for e in preflight_errors)
        raise ValueError(
            f"Preflight validation failed ({len(preflight_errors)} error(s)):\n{error_report}"
        )

    if portable_styles_path.exists():
        arch_styles_xml = portable_styles_path.read_text(encoding="utf-8")
    else:
        arch_styles_xml = build_arch_styles_xml_from_registry(env_registry)
    if not legacy_mode:
        if not HAS_NUMBERING_IMPORTER:
            raise ImportError("numbering_importer is required for strict Phase 1 bundles")
        all_style_ids = {
            item.get("style_id")
            for item in env_registry.get("styles", {}).get("style_defs", [])
            if isinstance(item, dict) and isinstance(item.get("style_id"), str)
        } | set(arch_registry.values())
        build_numbering_import_plan(
            env_registry,
            arch_styles_xml,
            '<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"></w:numbering>',
            sorted(all_style_ids),
            role_specs=role_specs,
            roles_to_apply=sorted(role_specs or {}),
        )
    raw_style_registry = json.loads(style_registry_path.read_text(encoding="utf-8"))
    source_tokens = raw_style_registry.get("source_tokens", {})
    return SharedConfig(
        arch_registry=arch_registry,
        env_registry=env_registry,
        arch_styles_xml=arch_styles_xml,
        available_roles=available_roles,
        source_tokens=source_tokens if isinstance(source_tokens, dict) else {},
        arch_root=arch_root,
        role_specs=role_specs,
        bundle_manifest=bundle_manifest,
        legacy_mode=legacy_mode,
    )


OPTIONAL_REPLACEMENT_PARTS = [
    ("word/theme/theme1.xml", lambda d: d / "word" / "theme" / "theme1.xml"),
    ("word/settings.xml", lambda d: d / "word" / "settings.xml"),
    ("word/fontTable.xml", lambda d: d / "word" / "fontTable.xml"),
    ("word/numbering.xml", lambda d: d / "word" / "numbering.xml"),
    ("[Content_Types].xml", lambda d: d / "[Content_Types].xml"),
    ("word/_rels/document.xml.rels", lambda d: d / "word" / "_rels" / "document.xml.rels"),
]


def _patch_header_footer_tokens_if_imported(
    extract_dir: Path,
    env_result: Dict[str, Any],
    source_tokens: Optional[Dict[str, str]],
    target_tokens: Optional[Dict[str, str]],
    log: List[str],
) -> bool:
    """Patch project tokens only in architect parts imported during this run."""
    if not source_tokens or not target_tokens:
        return False
    imported_parts = env_result.get("header_footer_import", {}).get("part_names", set())
    if not imported_parts:
        log.append(
            "No architect header/footer parts imported; preserved target tokens unchanged"
        )
        return False
    patch_header_footer_tokens(
        extract_dir,
        source_tokens,
        target_tokens,
        log,
        part_names=list(imported_parts),
    )
    return True


def _build_and_patch_output(
    docx_path: Path,
    extract_dir: Path,
    env_result: Dict[str, Any],
    output_dir: Path,
    arch_template_registry: Optional[Dict[str, Any]] = None,
    conversion_mode: str = FORMAT_ONLY,
) -> Path:
    conversion_mode = validate_conversion_mode(conversion_mode)
    output_dir.mkdir(parents=True, exist_ok=True)
    suffix = (
        "_CANADIAN_FORMATTED.docx"
        if conversion_mode == CSI_TO_CANADIAN
        else "_PHASE2_FORMATTED.docx"
    )
    output_path = output_dir / (docx_path.stem + suffix)
    replacements = {
        "word/document.xml": (extract_dir / "word" / "document.xml").read_bytes(),
        "word/styles.xml": (extract_dir / "word" / "styles.xml").read_bytes(),
    }

    for rel_path, path_builder in OPTIONAL_REPLACEMENT_PARTS:
        local_path = path_builder(extract_dir)
        if local_path.exists():
            replacements[rel_path] = local_path.read_bytes()

    hf_manifest = env_result.get("header_footer_import", {}) if isinstance(env_result, dict) else {}
    # Only explicitly imported architect parts are eligible for replacement.
    # When the bundle supplies no mapped header/footer, the source package
    # entries remain byte-identical even if an extracted working copy drifts.
    for key in ("part_names", "rels_names", "media_names"):
        for part_name in sorted(hf_manifest.get(key, [])):
            local_path = extract_dir / part_name
            if local_path.exists():
                replacements[part_name] = local_path.read_bytes()

    exclude_parts = set()
    if any(hf_manifest.get(key) for key in ("part_names", "rels_names")):
        old_hf_parts = set(hf_manifest.get("removed_part_names", set()))
        old_hf_rels = set(hf_manifest.get("removed_rels_names", set()))
        exclude_parts = (old_hf_parts | old_hf_rels) - set(replacements.keys())
    dynamic_parts = set().union(
        *(set(hf_manifest.get(key, set())) for key in (
            "part_names",
            "rels_names",
            "media_names",
            "removed_part_names",
            "removed_rels_names",
        ))
    )
    with tempfile.NamedTemporaryFile(
        prefix=f".{output_path.stem}.",
        suffix=".tmp.docx",
        dir=output_dir,
        delete=False,
    ) as tmp_file:
        temp_output_path = Path(tmp_file.name)

    try:
        patch_docx(
            src_docx=docx_path,
            out_docx=temp_output_path,
            replacements=replacements,
            exclude_parts=exclude_parts,
            allowed_dynamic_parts=dynamic_parts,
        )
        validate_docx_package(temp_output_path)
        verify_phase2_invariants(
            src_docx=docx_path,
            new_document_xml=replacements["word/document.xml"],
            new_docx=temp_output_path,
            arch_template_registry=arch_template_registry,
        )
        os.replace(temp_output_path, output_path)
    except Exception:
        temp_output_path.unlink(missing_ok=True)
        raise
    return output_path


def process_single_file(
    docx_path: Path,
    arch_registry: Dict[str, str],
    env_registry: Dict[str, Any],
    arch_styles_xml: str,
    available_roles: List[str],
    api_key: str,
    output_dir: Path,
    source_tokens: Optional[Dict[str, str]] = None,
    arch_root: Optional[Path] = None,
    model: str = "claude-sonnet-5",
    role_specs: Optional[Dict[str, Dict[str, Any]]] = None,
    conversion_mode: str = FORMAT_ONLY,
) -> BatchResult:
    start = time.monotonic()
    per_file_log: List[str] = []
    filename = docx_path.name
    output_path: Optional[Path] = None
    conversion_report: Optional[CanadianConversionReport] = None

    try:
        conversion_mode = validate_conversion_mode(conversion_mode)
        with tempfile.TemporaryDirectory(prefix="phase2_") as tmp_root:
            digest = hashlib.sha256(str(docx_path.resolve()).encode("utf-8")).hexdigest()[:8]
            extract_dir_name = f"{docx_path.stem}_{digest}_extracted"

            per_file_log.append("Extracting DOCX...")
            decomposer = DocxDecomposer(str(docx_path))
            extract_dir = decomposer.extract(output_dir=Path(tmp_root) / extract_dir_name)

            per_file_log.append("Building slim bundle...")
            bundle = build_phase2_slim_bundle(
                extract_dir,
                available_roles=available_roles,
                role_specs=role_specs,
            )
            unresolved = len(bundle.get("paragraphs", []))
            deterministic = len(bundle.get("deterministic_classifications", []))
            per_file_log.append(
                f"Built slim bundle: {unresolved} unresolved + {deterministic} deterministic"
            )

            if unresolved > 0 and not api_key:
                raise ValueError("Anthropic API key is required when unresolved paragraphs exist.")

            if unresolved:
                per_file_log.append("Classifying unresolved paragraphs with Anthropic...")
            else:
                per_file_log.append(
                    "All paragraphs classified deterministically; Anthropic request skipped"
                )
            classifications = classify_target_document(
                slim_bundle=bundle,
                available_roles=available_roles,
                api_key=api_key,
                model=model,
            )

            classifications_path = extract_dir / "phase2_classifications.json"
            classifications_path.write_text(json.dumps(classifications, indent=2), encoding="utf-8")
            per_file_log.append(f"Classifications saved: {classifications_path}")

            target_tokens = extract_target_tokens(extract_dir, classifications)
            application_classifications = classifications

            if conversion_mode == CSI_TO_CANADIAN:
                per_file_log.append("Converting CSI hierarchy to Canadian CSC PageFormat...")
                conversion_report = apply_csi_to_canadian(
                    extract_dir,
                    classifications,
                    role_specs,
                    per_file_log,
                    architect_numbering_xml=(
                        env_registry.get("numbering", {}).get("numbering_xml") or ""
                    ),
                )
                application_classifications = classifications_for_canadian_application(
                    classifications,
                    conversion_report,
                )

            env_result = apply_environment_to_target(
                target_extract_dir=extract_dir,
                registry=env_registry,
                log=per_file_log,
                registry_dir=arch_root,
            )
            per_file_log.append("Applied environment")

            _patch_header_footer_tokens_if_imported(
                extract_dir,
                env_result,
                source_tokens,
                target_tokens,
                per_file_log,
            )

            source_styles_xml = read_xml_text(extract_dir / "word" / "styles.xml")

            used_roles = {
                item.get("csi_role")
                for item in application_classifications.get("classifications", [])
                if isinstance(item, dict) and isinstance(item.get("csi_role"), str)
            }
            hf_style_ids = env_result.get("header_footer_import", {}).get("style_ids", set())
            hf_direct_num_ids = env_result.get("header_footer_import", {}).get("direct_num_ids", set())
            needed_style_ids = sorted(
                {arch_registry[r] for r in used_roles if r in arch_registry}
                | set(hf_style_ids)
            )

            style_numid_remap = {}
            role_numpr_remap = {}
            num_id_remap = {}
            if HAS_NUMBERING_IMPORTER:
                numbering_contract = import_numbering(
                    target_extract_dir=extract_dir,
                    arch_template_registry=env_registry,
                    arch_styles_xml=arch_styles_xml,
                    style_ids_to_import=needed_style_ids,
                    log=per_file_log,
                    role_specs=role_specs,
                    roles_to_apply=sorted(used_roles),
                    additional_num_ids=sorted(hf_direct_num_ids),
                    return_contract=True,
                )
                style_numid_remap = numbering_contract["style_numid_remap"]
                role_numpr_remap = numbering_contract["role_numpr_remap"]
                num_id_remap = numbering_contract["num_id_remap"]
            else:
                _check_numbering_module_needed(arch_styles_xml, needed_style_ids)
                if hf_direct_num_ids:
                    raise ImportError("numbering_importer is required by architect headers/footers")

            remap_header_footer_numids(
                extract_dir,
                list(env_result.get("header_footer_import", {}).get("part_names", set())),
                num_id_remap,
                per_file_log,
            )

            import_arch_styles_into_target(
                target_extract_dir=extract_dir,
                arch_styles_xml=arch_styles_xml,
                needed_style_ids=needed_style_ids,
                log=per_file_log,
                style_numid_remap=style_numid_remap,
            )
            per_file_log.append(f"Imported {len(needed_style_ids)} styles")

            snap = snapshot_stability(extract_dir)
            apply_report = apply_phase2_classifications(
                extract_dir=extract_dir,
                classifications=application_classifications,
                arch_style_registry=arch_registry,
                log=per_file_log,
                role_specs=role_specs,
                role_numpr_remap=role_numpr_remap,
                source_styles_xml=source_styles_xml,
            )
            verify_stability(extract_dir, snap)
            per_file_log.append("Applied classifications, stability verified")

            output_path = _build_and_patch_output(
                docx_path,
                extract_dir,
                env_result,
                output_dir,
                arch_template_registry=env_registry,
                conversion_mode=conversion_mode,
            )

            classified, total, unresolved = _coverage_counts(bundle, classifications)
            class_coverage = (classified / total * 100) if total > 0 else 100.0
            expected_targetable = apply_report.requested - len(apply_report.skipped_sectpr)
            app_coverage = (apply_report.modified / expected_targetable * 100) if expected_targetable > 0 else 100.0
            per_file_log.append(f"Output: {output_path}")
            per_file_log.append(f"Classification coverage: {classified}/{total} ({class_coverage:.1f}%)")
            per_file_log.append(
                f"Application coverage: {apply_report.modified}/{expected_targetable} ({app_coverage:.1f}%)"
            )

        return BatchResult(
            filename=filename,
            success=True,
            output_path=output_path,
            log=per_file_log,
            error=None,
            duration_seconds=time.monotonic() - start,
            conversion_report=conversion_report,
        )
    except Exception as exc:
        per_file_log.append(f"FAILED: {exc}")
        return BatchResult(
            filename=filename,
            success=False,
            output_path=output_path,
            log=per_file_log,
            error=str(exc),
            duration_seconds=time.monotonic() - start,
            conversion_report=conversion_report,
        )


def _prepare_file_for_batch(
    docx_path: Path,
    available_roles: List[str],
    extract_base_dir: Path,
    role_specs: Optional[Dict[str, Dict[str, Any]]] = None,
) -> PreparedFile:
    per_file_log: List[str] = []
    digest = hashlib.sha256(str(docx_path.resolve()).encode("utf-8")).hexdigest()[:8]
    extract_dir_name = f"{docx_path.stem}_{digest}_extracted"

    per_file_log.append("Extracting DOCX...")
    decomposer = DocxDecomposer(str(docx_path))
    extract_dir = decomposer.extract(output_dir=extract_base_dir / extract_dir_name)

    per_file_log.append("Building slim bundle...")
    bundle = build_phase2_slim_bundle(
        extract_dir,
        available_roles=available_roles,
        role_specs=role_specs,
    )
    unresolved = len(bundle.get("paragraphs", []))
    deterministic = len(bundle.get("deterministic_classifications", []))
    per_file_log.append(
        f"Built slim bundle: {unresolved} unresolved + {deterministic} deterministic"
    )
    return PreparedFile(file_key=_build_file_key(docx_path), docx_path=docx_path, extract_dir=extract_dir, bundle=bundle, prep_log=per_file_log)


def _apply_batch_result(
    prepared: PreparedFile,
    classifications: Dict[str, Any],
    arch_registry: Dict[str, str],
    env_registry: Dict[str, Any],
    arch_styles_xml: str,
    output_dir: Path,
    source_tokens: Optional[Dict[str, str]] = None,
    arch_root: Optional[Path] = None,
    role_specs: Optional[Dict[str, Dict[str, Any]]] = None,
    conversion_mode: str = FORMAT_ONLY,
) -> BatchResult:
    start = time.monotonic()
    per_file_log = list(prepared.prep_log)
    output_path: Optional[Path] = None
    conversion_report: Optional[CanadianConversionReport] = None
    filename = prepared.docx_path.name

    try:
        conversion_mode = validate_conversion_mode(conversion_mode)
        classifications_path = prepared.extract_dir / "phase2_classifications.json"
        classifications_path.write_text(json.dumps(classifications, indent=2), encoding="utf-8")
        per_file_log.append(f"Classifications saved: {classifications_path}")

        target_tokens = extract_target_tokens(prepared.extract_dir, classifications)
        application_classifications = classifications

        if conversion_mode == CSI_TO_CANADIAN:
            per_file_log.append("Converting CSI hierarchy to Canadian CSC PageFormat...")
            conversion_report = apply_csi_to_canadian(
                prepared.extract_dir,
                classifications,
                role_specs,
                per_file_log,
                architect_numbering_xml=(
                    env_registry.get("numbering", {}).get("numbering_xml") or ""
                ),
            )
            application_classifications = classifications_for_canadian_application(
                classifications,
                conversion_report,
            )

        env_result = apply_environment_to_target(
            target_extract_dir=prepared.extract_dir,
            registry=env_registry,
            log=per_file_log,
            registry_dir=arch_root,
        )
        per_file_log.append("Applied environment")
        _patch_header_footer_tokens_if_imported(
            prepared.extract_dir,
            env_result,
            source_tokens,
            target_tokens,
            per_file_log,
        )

        source_styles_xml = read_xml_text(
            prepared.extract_dir / "word" / "styles.xml"
        )

        used_roles = {
            item.get("csi_role")
            for item in application_classifications.get("classifications", [])
            if isinstance(item, dict) and isinstance(item.get("csi_role"), str)
        }
        hf_style_ids = env_result.get("header_footer_import", {}).get("style_ids", set())
        hf_direct_num_ids = env_result.get("header_footer_import", {}).get("direct_num_ids", set())
        needed_style_ids = sorted(
            {arch_registry[r] for r in used_roles if r in arch_registry}
            | set(hf_style_ids)
        )

        style_numid_remap = {}
        role_numpr_remap = {}
        num_id_remap = {}
        if HAS_NUMBERING_IMPORTER:
            numbering_contract = import_numbering(
                target_extract_dir=prepared.extract_dir,
                arch_template_registry=env_registry,
                arch_styles_xml=arch_styles_xml,
                style_ids_to_import=needed_style_ids,
                log=per_file_log,
                role_specs=role_specs,
                roles_to_apply=sorted(used_roles),
                additional_num_ids=sorted(hf_direct_num_ids),
                return_contract=True,
            )
            style_numid_remap = numbering_contract["style_numid_remap"]
            role_numpr_remap = numbering_contract["role_numpr_remap"]
            num_id_remap = numbering_contract["num_id_remap"]
        else:
            _check_numbering_module_needed(arch_styles_xml, needed_style_ids)
            if hf_direct_num_ids:
                raise ImportError("numbering_importer is required by architect headers/footers")

        remap_header_footer_numids(
            prepared.extract_dir,
            list(env_result.get("header_footer_import", {}).get("part_names", set())),
            num_id_remap,
            per_file_log,
        )

        import_arch_styles_into_target(
            target_extract_dir=prepared.extract_dir,
            arch_styles_xml=arch_styles_xml,
            needed_style_ids=needed_style_ids,
            log=per_file_log,
            style_numid_remap=style_numid_remap,
        )
        per_file_log.append(f"Imported {len(needed_style_ids)} styles")

        snap = snapshot_stability(prepared.extract_dir)
        apply_report = apply_phase2_classifications(
            extract_dir=prepared.extract_dir,
            classifications=application_classifications,
            arch_style_registry=arch_registry,
            log=per_file_log,
            role_specs=role_specs,
            role_numpr_remap=role_numpr_remap,
            source_styles_xml=source_styles_xml,
        )
        verify_stability(prepared.extract_dir, snap)
        per_file_log.append("Applied classifications, stability verified")

        output_path = _build_and_patch_output(
            prepared.docx_path,
            prepared.extract_dir,
            env_result,
            output_dir,
            arch_template_registry=env_registry,
            conversion_mode=conversion_mode,
        )

        classified, total, unresolved = _coverage_counts(prepared.bundle, classifications)
        class_coverage = (classified / total * 100) if total > 0 else 100.0
        expected_targetable = apply_report.requested - len(apply_report.skipped_sectpr)
        app_coverage = (apply_report.modified / expected_targetable * 100) if expected_targetable > 0 else 100.0
        per_file_log.append(f"Output: {output_path}")
        per_file_log.append(f"Classification coverage: {classified}/{total} ({class_coverage:.1f}%)")
        per_file_log.append(
            f"Application coverage: {apply_report.modified}/{expected_targetable} ({app_coverage:.1f}%)"
        )

        return BatchResult(
            filename=filename,
            success=True,
            output_path=output_path,
            log=per_file_log,
            error=None,
            duration_seconds=time.monotonic() - start,
            conversion_report=conversion_report,
        )
    except Exception as exc:
        per_file_log.append(f"FAILED: {exc}")
        return BatchResult(
            filename=filename,
            success=False,
            output_path=output_path,
            log=per_file_log,
            error=str(exc),
            duration_seconds=time.monotonic() - start,
            conversion_report=conversion_report,
        )


def run_batch_concurrent(
    docx_paths: List[Path],
    arch_registry: Dict[str, str],
    env_registry: Dict[str, Any],
    arch_styles_xml: str,
    available_roles: List[str],
    api_key: str,
    output_dir: Path,
    source_tokens: Optional[Dict[str, str]] = None,
    arch_root: Optional[Path] = None,
    max_workers: int = 3,
    on_file_complete: Optional[Callable[[BatchResult], None]] = None,
    role_specs: Optional[Dict[str, Dict[str, Any]]] = None,
    conversion_mode: str = FORMAT_ONLY,
) -> List[BatchResult]:
    conversion_mode = validate_conversion_mode(conversion_mode)
    if not docx_paths:
        return []

    workers = max(1, min(max_workers, len(docx_paths)))
    results: List[BatchResult] = []

    with ThreadPoolExecutor(max_workers=workers) as executor:
        if (
            source_tokens is None
            and arch_root is None
            and role_specs is None
            and conversion_mode == FORMAT_ONLY
        ):
            futures = {
                executor.submit(
                    process_single_file,
                    docx_path,
                    arch_registry,
                    env_registry,
                    arch_styles_xml,
                    available_roles,
                    api_key,
                    output_dir,
                ): docx_path
                for docx_path in docx_paths
            }
        else:
            futures = {
                executor.submit(
                    process_single_file,
                    docx_path,
                    arch_registry,
                    env_registry,
                    arch_styles_xml,
                    available_roles,
                    api_key,
                    output_dir,
                    source_tokens=source_tokens,
                    arch_root=arch_root,
                    role_specs=role_specs,
                    conversion_mode=conversion_mode,
                ): docx_path
                for docx_path in docx_paths
            }

        for future in as_completed(futures):
            result = future.result()
            results.append(result)
            if on_file_complete:
                on_file_complete(result)

    return sorted(results, key=lambda item: item.filename)


def run_batch_api(
    docx_paths: List[Path],
    arch_registry: Dict[str, str],
    env_registry: Dict[str, Any],
    arch_styles_xml: str,
    available_roles: List[str],
    api_key: str,
    output_dir: Path,
    source_tokens: Optional[Dict[str, str]] = None,
    arch_root: Optional[Path] = None,
    max_workers: int = 3,
    poll_interval: int = 30,
    on_file_complete: Optional[Callable[[BatchResult], None]] = None,
    on_batch_poll: Optional[Callable[[str, str, Any], None]] = None,
    model: str = "claude-sonnet-5",
    role_specs: Optional[Dict[str, Dict[str, Any]]] = None,
    conversion_mode: str = FORMAT_ONLY,
) -> List[BatchResult]:
    conversion_mode = validate_conversion_mode(conversion_mode)
    if not docx_paths:
        return []

    workers = max(1, min(max_workers, len(docx_paths)))
    prepared_files: Dict[str, PreparedFile] = {}

    with tempfile.TemporaryDirectory(prefix="phase2_batch_") as tmp_root:
        tmp_base = Path(tmp_root)
        with ThreadPoolExecutor(max_workers=workers) as executor:
            futures = {
                executor.submit(
                    _prepare_file_for_batch,
                    docx_path,
                    available_roles,
                    tmp_base,
                    role_specs,
                ): docx_path
                for docx_path in docx_paths
            }
            for future in as_completed(futures):
                prepared = future.result()
                prepared_files[prepared.file_key] = prepared

        file_bundles = {key: prepared.bundle for key, prepared in prepared_files.items()}
        requests = build_batch_requests(file_bundles, available_roles, model)

        raw_results = submit_and_poll(
            requests=requests,
            api_key=api_key,
            poll_interval=poll_interval,
            on_poll=on_batch_poll,
        )

        try:
            per_file_classifications = reassemble_file_classifications(raw_results, file_bundles, available_roles)
        except BatchClassificationError:
            raise

        results: List[BatchResult] = []
        with ThreadPoolExecutor(max_workers=workers) as executor:
            if (
                source_tokens is None
                and arch_root is None
                and role_specs is None
                and conversion_mode == FORMAT_ONLY
            ):
                futures = {
                    executor.submit(
                        _apply_batch_result,
                        prepared,
                        per_file_classifications[file_key],
                        arch_registry,
                        env_registry,
                        arch_styles_xml,
                        output_dir,
                    ): file_key
                    for file_key, prepared in prepared_files.items()
                }
            else:
                futures = {
                    executor.submit(
                        _apply_batch_result,
                        prepared,
                        per_file_classifications[file_key],
                        arch_registry,
                        env_registry,
                        arch_styles_xml,
                        output_dir,
                        source_tokens,
                        arch_root,
                        role_specs,
                        conversion_mode,
                    ): file_key
                    for file_key, prepared in prepared_files.items()
                }
            for future in as_completed(futures):
                result = future.result()
                results.append(result)
                if on_file_complete:
                    on_file_complete(result)

        return sorted(results, key=lambda item: item.filename)


def _build_file_key(docx_path: Path) -> str:
    # Batch API custom_ids must match [a-zA-Z0-9_-]{1,64}. The key becomes
    # "<stem>__<digest>__chunk<N>", so strip '.' from the stem and bound its
    # length; the path digest keeps truncated keys unique.
    safe_stem = re.sub(r"[^A-Za-z0-9_-]", "_", docx_path.stem)[:38]
    digest = hashlib.sha1(str(docx_path.resolve()).encode("utf-8")).hexdigest()[:12]
    return f"{safe_stem}__{digest}"
