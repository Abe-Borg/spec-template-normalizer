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
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any, Callable, Dict, List, Optional

from .. import diagnostics as diag
from .arch_env_applier import apply_environment_to_target
from .core.classification import apply_phase2_classifications, build_phase2_slim_bundle
from .core.application_policy import ApplicationPolicy, application_policy_for_mode
from .core.csi_to_canadian import (
    CSI_TO_CANADIAN,
    FORMAT_ONLY,
    CanadianConversionReport,
    ConversionIssue,
    MarkerEdit,
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
from .core.ooxml_text import read_xml_text, write_xml_text
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
    audit_summary: Dict[str, int] = field(default_factory=dict)
    audit: Dict[str, Any] = field(default_factory=dict)
    numbering_checks: Dict[str, Any] = field(default_factory=dict)
<<<<<<< HEAD
    stage: Optional[str] = None


@dataclass(frozen=True)
class ApplicationFailureDiagnostics:
    """Text-free checkpoints retained when classified-target application fails."""

    stage: str
    conversion_report: Optional[CanadianConversionReport] = None
    audit_summary: Dict[str, int] = field(default_factory=dict)
    audit: Dict[str, Any] = field(default_factory=dict)
    numbering_checks: Dict[str, Any] = field(default_factory=dict)

    def as_dict(self) -> Dict[str, Any]:
        """Return the structural, JSON-safe representation for outer pipelines."""

        return {
            "stage": self.stage,
            "conversion_report": (
                self.conversion_report.as_dict()
                if self.conversion_report is not None
                else None
            ),
            "audit_summary": dict(self.audit_summary),
            "audit": dict(self.audit),
            "numbering_checks": dict(self.numbering_checks),
        }


class ApplicationStageError(RuntimeError):
    """Application failure augmented with the last safe diagnostic checkpoint."""

    def __init__(
        self,
        message: str,
        *,
        diagnostics: ApplicationFailureDiagnostics,
    ) -> None:
        super().__init__(message)
        self.diagnostics = diagnostics

    @property
    def stage(self) -> str:
        return self.diagnostics.stage

    @property
    def conversion_report(self) -> Optional[CanadianConversionReport]:
        return self.diagnostics.conversion_report

    @property
    def audit_summary(self) -> Dict[str, int]:
        return self.diagnostics.audit_summary

    @property
    def audit(self) -> Dict[str, Any]:
        return self.diagnostics.audit

    @property
    def numbering_checks(self) -> Dict[str, Any]:
        return self.diagnostics.numbering_checks


@dataclass
class _ApplicationCheckpoint:
    stage: str
    audit_summary: Dict[str, int]
    audit: Dict[str, Any]
    conversion_report: Optional[CanadianConversionReport] = None
    numbering_checks: Dict[str, Any] = field(default_factory=dict)

    def failure_diagnostics(self) -> ApplicationFailureDiagnostics:
        return ApplicationFailureDiagnostics(
            stage=self.stage,
            conversion_report=_safe_conversion_report(self.conversion_report),
            audit_summary=dict(self.audit_summary),
            audit=_safe_application_audit(self.audit),
            numbering_checks=_safe_numbering_checks(self.numbering_checks),
        )
=======
    # Structured, redaction-safe phase-timing/count events for this target.
    # Carries no free text; the pipeline folds it into the run diagnostics.
    diagnostics: List[Dict[str, Any]] = field(default_factory=list)
>>>>>>> 769bf0ca3a9a2744c852a1a23a7b3e5f88efb5b3


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
    total = (
        len(bundle.get("paragraphs", []))
        + len(bundle.get("deterministic_classifications", []))
        + len(bundle.get("deterministic_ignored_paragraphs", []))
    )
    resolved = len(classifications.get("classifications", [])) + len(
        classifications.get("ignored_paragraphs", [])
    )
    return resolved, total, len(bundle.get("paragraphs", []))


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
    if not target_tokens:
        return False
    imported_parts = env_result.get("header_footer_import", {}).get("part_names", set())
    if not imported_parts:
        log.append(
            "No architect header/footer parts imported; preserved target tokens unchanged"
        )
        return False
    patch_header_footer_tokens(
        extract_dir,
        source_tokens or {},
        target_tokens,
        log,
        part_names=sorted(imported_parts),
    )
    return True


def _remap_imported_header_footer_style_ids(
    extract_dir: Path,
    part_names: List[str],
    style_id_map: Dict[str, str],
    log: List[str],
) -> None:
    """Point imported header/footer content at collision-safe style clones."""

    replacements = {
        source: destination
        for source, destination in style_id_map.items()
        if source != destination
    }
    if not replacements:
        return
    changed_parts = 0
    for part_name in sorted(set(part_names)):
        path = extract_dir / part_name
        if not path.is_file() or path.suffix.lower() != ".xml":
            continue
        original = read_xml_text(path)
        updated = re.sub(
            r'(<w:(?:pStyle|rStyle|tblStyle)\b[^>]*w:val=")([^"]+)(")',
            lambda match: (
                match.group(1)
                + replacements.get(match.group(2), match.group(2))
                + match.group(3)
            ),
            original,
        )
        if updated != original:
            write_xml_text(path, updated)
            changed_parts += 1
    if changed_parts:
        log.append(
            f"Remapped collision-safe style IDs in {changed_parts} imported header/footer parts"
        )


_DIAGNOSTIC_IDENTIFIER = re.compile(r"[A-Za-z0-9_.:/#@+\-]{1,160}")
_SAFE_AUDIT_REASONS = frozenset(
    {
        "boilerplate",
        "drawing_or_textbox_subtree",
        "editorial_comment_style",
        "end_of_section_no_role",
        "non_csi_content",
        "section_header_no_role",
        "section_title_no_role",
        "table",
    }
)


def _safe_identifier(value: Any, fallback: str = "unspecified") -> str:
    if isinstance(value, str) and _DIAGNOSTIC_IDENTIFIER.fullmatch(value):
        return value
    return fallback


def _safe_conversion_report(
    report: Optional[CanadianConversionReport],
) -> Optional[CanadianConversionReport]:
    """Strip paragraph previews, messages, and literal markers from a report."""

    if report is None:
        return None
    edits = tuple(
        MarkerEdit(
            paragraph_index=item.paragraph_index,
            role=_safe_identifier(item.role),
            source_kind=_safe_identifier(item.source_kind),
            target_kind=_safe_identifier(item.target_kind),
            source_marker=None,
            target_marker=None,
        )
        for item in report.edits
    )
    warnings = tuple(
        ConversionIssue(
            paragraph_index=item.paragraph_index,
            code=_safe_identifier(item.code),
            message=f"Conversion warning: {_safe_identifier(item.code)}",
            text_preview="",
        )
        for item in report.warnings
    )
    return CanadianConversionReport(
        paragraphs_examined=report.paragraphs_examined,
        paragraphs_converted=report.paragraphs_converted,
        literal_markers_removed=report.literal_markers_removed,
        automatic_numbering_retargeted=report.automatic_numbering_retargeted,
        unnumbered_paragraphs_numbered=report.unnumbered_paragraphs_numbered,
        edits=edits,
        warnings=warnings,
    )


def _safe_numbering_checks(value: Any) -> Dict[str, Any]:
    """Retain only structural scalar/list values from numbering diagnostics."""

    if not isinstance(value, dict):
        return {}
    safe: Dict[str, Any] = {}
    for key, item in value.items():
        if not isinstance(key, str):
            continue
        if item is None or isinstance(item, (bool, int, float)):
            safe[key] = item
        elif (
            key in {"conversion_mode", "policy", "status"}
            and isinstance(item, str)
            and _DIAGNOSTIC_IDENTIFIER.fullmatch(item)
        ):
            safe[key] = item
        elif isinstance(item, (list, tuple)) and all(
            element is None or isinstance(element, (bool, int, float))
            for element in item
        ):
            safe[key] = list(item)
    return safe


def _safe_application_audit(value: Any) -> Dict[str, Any]:
    """Project an application audit to disposition metadata only."""

    if not isinstance(value, dict):
        return {}

    def dispositions(items: Any, value_key: str) -> List[Dict[str, Any]]:
        if not isinstance(items, list):
            return []
        safe_items: List[Dict[str, Any]] = []
        for item in items:
            if not isinstance(item, dict):
                continue
            index = item.get("paragraph_index")
            if not isinstance(index, int) or isinstance(index, bool):
                continue
            raw_value = item.get(value_key)
            if value_key == "reason":
                safe_value = (
                    raw_value
                    if isinstance(raw_value, str)
                    and raw_value in _SAFE_AUDIT_REASONS
                    else "unspecified"
                )
            else:
                safe_value = _safe_identifier(raw_value)
            safe_items.append(
                {"paragraph_index": index, value_key: safe_value}
            )
        return safe_items

    raw_summary = value.get("summary", {})
    summary = (
        {
            key: count
            for key, count in raw_summary.items()
            if isinstance(key, str)
            and isinstance(count, int)
            and not isinstance(count, bool)
            and count >= 0
        }
        if isinstance(raw_summary, dict)
        else {}
    )
    return {
        "schema_version": 1,
        "summary": summary,
        "classifications": dispositions(
            value.get("classifications", []), "csi_role"
        ),
        "ignored_paragraphs": dispositions(
            value.get("ignored_paragraphs", []), "reason"
        ),
        "out_of_scope": dispositions(value.get("out_of_scope", []), "reason"),
    }


def _classification_audit(
    bundle: Dict[str, Any],
    classifications: Dict[str, Any],
) -> tuple[Dict[str, int], Dict[str, Any]]:
    styled = classifications.get("classifications", [])
    ignored = classifications.get("ignored_paragraphs", [])
    filter_report = bundle.get("filter_report", {})
    out_of_scope = filter_report.get("paragraphs_out_of_scope", [])
    total = (
        len(bundle.get("paragraphs", []))
        + len(bundle.get("deterministic_classifications", []))
        + len(bundle.get("deterministic_ignored_paragraphs", []))
    )
    resolved = len(styled) + len(ignored)
    summary = {
        "styled": len(styled),
        "ignored": len(ignored),
        "out_of_scope": len(out_of_scope),
        "unresolved": max(0, total - resolved),
    }
    audit = {
        "schema_version": 1,
        "summary": summary,
        "classifications": styled,
        "ignored_paragraphs": ignored,
        "out_of_scope": out_of_scope,
    }
    return summary, audit


def _build_and_patch_output(
    docx_path: Path,
    extract_dir: Path,
    env_result: Dict[str, Any],
    output_dir: Path,
    arch_template_registry: Optional[Dict[str, Any]] = None,
    conversion_mode: str = FORMAT_ONLY,
    allowed_rpr_properties_by_paragraph: Optional[Dict[int, set[str]]] = None,
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
        prefix=".sf-",
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
            conversion_mode=conversion_mode,
            allowed_rpr_properties_by_paragraph=(
                allowed_rpr_properties_by_paragraph
            ),
        )
        os.replace(temp_output_path, output_path)
    except Exception:
        temp_output_path.unlink(missing_ok=True)
        raise
    return output_path


def _apply_classified_target_impl(
    *,
    docx_path: Path,
    extract_dir: Path,
    bundle: Dict[str, Any],
    classifications: Dict[str, Any],
    arch_registry: Dict[str, str],
    env_registry: Dict[str, Any],
    arch_styles_xml: str,
    output_dir: Path,
    log: List[str],
    source_tokens: Optional[Dict[str, str]],
    arch_root: Optional[Path],
    role_specs: Optional[Dict[str, Dict[str, Any]]],
    conversion_mode: str,
<<<<<<< HEAD
    checkpoint: _ApplicationCheckpoint,
) -> tuple[Path, Optional[CanadianConversionReport], Dict[str, int], Dict[str, Any], Dict[str, Any]]:
    """Apply one validated classification payload through the shared engine."""

    checkpoint.stage = "application_policy"
    policy: ApplicationPolicy = application_policy_for_mode(conversion_mode)
    checkpoint.stage = "classification_checkpoint"
=======
    diagnostics: Optional[List[Dict[str, Any]]] = None,
) -> tuple[Path, Optional[CanadianConversionReport], Dict[str, int], Dict[str, Any], Dict[str, Any]]:
    """Apply one validated classification payload through the shared engine."""

    # A throwaway sink keeps callers that do not collect diagnostics working
    # without scattering ``if diagnostics is not None`` across every phase.
    diag_events: List[Dict[str, Any]] = diagnostics if diagnostics is not None else []

    policy: ApplicationPolicy = application_policy_for_mode(conversion_mode)
    diag.emit(diag_events, "INFO", "target", "policy",
              conversion_mode=policy.conversion_mode,
              import_body_numbering=bool(policy.import_body_numbering),
              preserve_target_numbering=bool(policy.preserve_target_numbering))
>>>>>>> 769bf0ca3a9a2744c852a1a23a7b3e5f88efb5b3
    classifications_path = extract_dir / "phase2_classifications.json"
    classifications_path.write_text(json.dumps(classifications, indent=2), encoding="utf-8")
    log.append("Classification checkpoint saved")

    # Capture the target's style and list catalogs before any architect
    # environment or numbering parts are imported.
    checkpoint.stage = "source_catalog_snapshot"
    source_styles_xml = read_xml_text(extract_dir / "word" / "styles.xml")
    source_numbering_path = extract_dir / "word" / "numbering.xml"
    source_numbering_xml = (
        read_xml_text(source_numbering_path) if source_numbering_path.is_file() else ""
    )

    checkpoint.stage = "target_token_extraction"
    target_tokens = extract_target_tokens(extract_dir, classifications)
    application_classifications = classifications
    conversion_report: Optional[CanadianConversionReport] = None

    if policy.convert_to_canadian:
        checkpoint.stage = "csi_conversion"
        log.append("Converting CSI hierarchy to Canadian CSC PageFormat...")
<<<<<<< HEAD
        conversion_report = apply_csi_to_canadian(
            extract_dir,
            classifications,
            role_specs,
            log,
            architect_numbering_xml=(
                env_registry.get("numbering", {}).get("numbering_xml") or ""
            ),
        )
        checkpoint.conversion_report = conversion_report
        checkpoint.stage = "canadian_classification_mapping"
=======
        with diag.timed(diag_events, "target", "csi_to_canadian") as phase:
            conversion_report = apply_csi_to_canadian(
                extract_dir,
                classifications,
                role_specs,
                log,
                architect_numbering_xml=(
                    env_registry.get("numbering", {}).get("numbering_xml") or ""
                ),
            )
            phase.set(
                paragraphs_examined=conversion_report.paragraphs_examined,
                paragraphs_converted=conversion_report.paragraphs_converted,
                literal_markers_removed=conversion_report.literal_markers_removed,
                automatic_numbering_retargeted=(
                    conversion_report.automatic_numbering_retargeted
                ),
                unnumbered_paragraphs_numbered=(
                    conversion_report.unnumbered_paragraphs_numbered
                ),
                warnings=len(conversion_report.warnings),
            )
>>>>>>> 769bf0ca3a9a2744c852a1a23a7b3e5f88efb5b3
        application_classifications = classifications_for_canadian_application(
            classifications,
            conversion_report,
        )

<<<<<<< HEAD
    checkpoint.stage = "environment_application"
    env_result = apply_environment_to_target(
        target_extract_dir=extract_dir,
        registry=env_registry,
        log=log,
        registry_dir=arch_root,
    )
=======
    with diag.timed(diag_events, "target", "apply_environment") as phase:
        env_result = apply_environment_to_target(
            target_extract_dir=extract_dir,
            registry=env_registry,
            log=log,
            registry_dir=arch_root,
        )
        _hf_import = env_result.get("header_footer_import", {}) if isinstance(env_result, dict) else {}
        phase.set(
            header_footer_parts=len(_hf_import.get("part_names", set()) or set()),
            header_footer_media=len(_hf_import.get("media_names", set()) or set()),
        )
>>>>>>> 769bf0ca3a9a2744c852a1a23a7b3e5f88efb5b3
    log.append("Applied environment")
    checkpoint.stage = "header_footer_token_patch"
    _patch_header_footer_tokens_if_imported(
        extract_dir,
        env_result,
        source_tokens,
        target_tokens,
        log,
    )

    used_roles = {
        item.get("csi_role")
        for item in application_classifications.get("classifications", [])
        if isinstance(item, dict) and isinstance(item.get("csi_role"), str)
    }
    body_style_ids = {arch_registry[r] for r in used_roles if r in arch_registry}
    hf_manifest = env_result.get("header_footer_import", {})
    hf_style_ids = set(hf_manifest.get("style_ids", set()))
    hf_direct_num_ids = set(hf_manifest.get("direct_num_ids", set()))
    needed_style_ids = sorted(body_style_ids | hf_style_ids)

    style_numid_remap: Dict[str, Dict[str, int]] = {}
    role_numpr_remap: Dict[str, Dict[str, Any]] = {}
    num_id_remap: Dict[int, int] = {}
    numbering_style_ids = needed_style_ids if policy.import_body_numbering else sorted(hf_style_ids)
    numbering_roles = sorted(used_roles) if policy.import_body_numbering else []
<<<<<<< HEAD
    checkpoint.stage = "numbering_import"
    if HAS_NUMBERING_IMPORTER:
        numbering_contract = import_numbering(
            target_extract_dir=extract_dir,
            arch_template_registry=env_registry,
            arch_styles_xml=arch_styles_xml,
            style_ids_to_import=numbering_style_ids,
            log=log,
            role_specs=role_specs,
            roles_to_apply=numbering_roles,
            additional_num_ids=sorted(hf_direct_num_ids),
            return_contract=True,
=======
    with diag.timed(diag_events, "target", "numbering_import") as phase:
        if HAS_NUMBERING_IMPORTER:
            numbering_contract = import_numbering(
                target_extract_dir=extract_dir,
                arch_template_registry=env_registry,
                arch_styles_xml=arch_styles_xml,
                style_ids_to_import=numbering_style_ids,
                log=log,
                role_specs=role_specs,
                roles_to_apply=numbering_roles,
                additional_num_ids=sorted(hf_direct_num_ids),
                return_contract=True,
            )
            style_numid_remap = numbering_contract["style_numid_remap"]
            role_numpr_remap = numbering_contract["role_numpr_remap"]
            num_id_remap = numbering_contract["num_id_remap"]
        else:
            _check_numbering_module_needed(arch_styles_xml, numbering_style_ids)
            if hf_direct_num_ids:
                raise ImportError("numbering_importer is required by architect headers/footers")
        phase.set(
            importer_available=bool(HAS_NUMBERING_IMPORTER),
            styles_considered=len(numbering_style_ids),
            roles_considered=len(numbering_roles),
            num_id_remaps=len(num_id_remap),
            style_numid_remaps=len(style_numid_remap),
            role_numpr_remaps=len(role_numpr_remap),
>>>>>>> 769bf0ca3a9a2744c852a1a23a7b3e5f88efb5b3
        )

    checkpoint.stage = "header_footer_numbering_remap"
    remap_header_footer_numids(
        extract_dir,
        list(hf_manifest.get("part_names", set())),
        num_id_remap,
        log,
    )

<<<<<<< HEAD
    checkpoint.stage = "style_import"
    style_result = import_arch_styles_into_target(
        target_extract_dir=extract_dir,
        arch_styles_xml=arch_styles_xml,
        needed_style_ids=needed_style_ids,
        log=log,
        style_numid_remap=style_numid_remap,
        format_only_body_style_ids=(body_style_ids if policy.preserve_target_numbering else None),
        shell_style_ids=hf_style_ids,
        namespace_seed=hashlib.sha256(arch_styles_xml.encode("utf-8")).hexdigest(),
    )
=======
    with diag.timed(diag_events, "target", "style_import") as phase:
        style_result = import_arch_styles_into_target(
            target_extract_dir=extract_dir,
            arch_styles_xml=arch_styles_xml,
            needed_style_ids=needed_style_ids,
            log=log,
            style_numid_remap=style_numid_remap,
            format_only_body_style_ids=(body_style_ids if policy.preserve_target_numbering else None),
            shell_style_ids=hf_style_ids,
            namespace_seed=hashlib.sha256(arch_styles_xml.encode("utf-8")).hexdigest(),
        )
        phase.set(
            requested_styles=len(needed_style_ids),
            body_style_ids=len(body_style_ids),
            header_footer_style_ids=len(hf_style_ids),
            namespaced_collisions=sum(
                1 for src, dst in style_result.style_id_map.items() if src != dst
            ),
        )
>>>>>>> 769bf0ca3a9a2744c852a1a23a7b3e5f88efb5b3
    applied_arch_registry = {
        role: style_result.body_style_id_map.get(style_id, style_id)
        for role, style_id in arch_registry.items()
    }
    checkpoint.stage = "header_footer_style_remap"
    _remap_imported_header_footer_style_ids(
        extract_dir,
        list(hf_manifest.get("part_names", set())),
        style_result.style_id_map,
        log,
    )
    log.append(f"Imported {len(needed_style_ids)} requested styles collision-safely")

    checkpoint.stage = "stability_snapshot"
    snap = snapshot_stability(extract_dir)
<<<<<<< HEAD
    checkpoint.stage = "classification_application"
    apply_report = apply_phase2_classifications(
        extract_dir=extract_dir,
        classifications=application_classifications,
        arch_style_registry=applied_arch_registry,
        log=log,
        role_specs=role_specs,
        role_numpr_remap=role_numpr_remap,
        source_styles_xml=source_styles_xml,
        source_numbering_xml=source_numbering_xml,
        policy=policy,
    )
    checkpoint.numbering_checks = dict(
        getattr(apply_report, "numbering_checks", {}) or {}
    )
    checkpoint.stage = "stability_verification"
    verify_stability(extract_dir, snap)
    log.append("Applied classifications, stability verified")

    checkpoint.stage = "application_reporting"
    classified, total, _unresolved = _coverage_counts(bundle, classifications)
=======
    with diag.timed(diag_events, "target", "apply_classifications") as phase:
        apply_report = apply_phase2_classifications(
            extract_dir=extract_dir,
            classifications=application_classifications,
            arch_style_registry=applied_arch_registry,
            log=log,
            role_specs=role_specs,
            role_numpr_remap=role_numpr_remap,
            source_styles_xml=source_styles_xml,
            source_numbering_xml=source_numbering_xml,
            policy=policy,
        )
        phase.set(
            requested=apply_report.requested,
            modified=apply_report.modified,
            skipped_sectpr=len(apply_report.skipped_sectpr),
            invalid_indices=len(apply_report.invalid_indices),
            unmapped_roles=len(apply_report.unmapped_roles),
            missing_style_ids=len(apply_report.missing_style_ids),
            stripped_direct_ppr=apply_report.stripped_direct_ppr,
            preserved_direct_ppr=apply_report.preserved_direct_ppr,
            preserved_automatic_numbering=apply_report.preserved_automatic_numbering,
            suppressed_architect_numbering=apply_report.suppressed_architect_numbering,
            stripped_run_fonts=apply_report.stripped_run_fonts,
            ignored=apply_report.ignored,
        )
    verify_stability(extract_dir, snap)
    log.append("Applied classifications, stability verified")

    with diag.timed(diag_events, "target", "build_output"):
        output_path = _build_and_patch_output(
            docx_path,
            extract_dir,
            env_result,
            output_dir,
            arch_template_registry=env_registry,
            conversion_mode=policy.conversion_mode,
            allowed_rpr_properties_by_paragraph=(
                apply_report.allowed_rpr_properties_by_paragraph
            ),
        )

    classified, total, unresolved = _coverage_counts(bundle, classifications)
>>>>>>> 769bf0ca3a9a2744c852a1a23a7b3e5f88efb5b3
    class_coverage = (classified / total * 100) if total > 0 else 100.0
    expected_targetable = apply_report.requested - len(apply_report.skipped_sectpr)
    app_coverage = (
        (apply_report.modified / expected_targetable * 100)
        if expected_targetable > 0
        else 100.0
    )
<<<<<<< HEAD
    classification_coverage_log = (
        f"Classification coverage: {classified}/{total} ({class_coverage:.1f}%)"
    )
    application_coverage_log = (
        f"Application coverage: {apply_report.modified}/{expected_targetable} ({app_coverage:.1f}%)"
    )

    # Atomic packaging is deliberately last: if any earlier stage fails there
    # is no output path to publish, and the builder itself removes its temp file.
    checkpoint.stage = "output_publication"
    output_path = _build_and_patch_output(
        docx_path,
        extract_dir,
        env_result,
        output_dir,
        arch_template_registry=env_registry,
        conversion_mode=policy.conversion_mode,
        allowed_rpr_properties_by_paragraph=(
            apply_report.allowed_rpr_properties_by_paragraph
        ),
    )
=======
>>>>>>> 769bf0ca3a9a2744c852a1a23a7b3e5f88efb5b3
    log.append(f"Output: {output_path}")
    log.append(classification_coverage_log)
    log.append(application_coverage_log)
    checkpoint.stage = "complete"
    return (
        output_path,
        conversion_report,
        checkpoint.audit_summary,
        checkpoint.audit,
        checkpoint.numbering_checks,
    )


def _apply_classified_target(
    *,
    docx_path: Path,
    extract_dir: Path,
    bundle: Dict[str, Any],
    classifications: Dict[str, Any],
    arch_registry: Dict[str, str],
    env_registry: Dict[str, Any],
    arch_styles_xml: str,
    output_dir: Path,
    log: List[str],
    source_tokens: Optional[Dict[str, str]],
    arch_root: Optional[Path],
    role_specs: Optional[Dict[str, Dict[str, Any]]],
    conversion_mode: str,
) -> tuple[Path, Optional[CanadianConversionReport], Dict[str, int], Dict[str, Any], Dict[str, Any]]:
    """Apply classifications while preserving safe late-failure diagnostics."""

    audit_summary, audit = _classification_audit(bundle, classifications)
    checkpoint = _ApplicationCheckpoint(
        stage="classification_ready",
        audit_summary=audit_summary,
        audit=audit,
    )
    try:
        return _apply_classified_target_impl(
            docx_path=docx_path,
            extract_dir=extract_dir,
            bundle=bundle,
            classifications=classifications,
            arch_registry=arch_registry,
            env_registry=env_registry,
            arch_styles_xml=arch_styles_xml,
            output_dir=output_dir,
            log=log,
            source_tokens=source_tokens,
            arch_root=arch_root,
            role_specs=role_specs,
            conversion_mode=conversion_mode,
            checkpoint=checkpoint,
        )
    except ApplicationStageError:
        raise
    except Exception as exc:
        raise ApplicationStageError(
            str(exc),
            diagnostics=checkpoint.failure_diagnostics(),
        ) from exc


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
    per_file_diag: List[Dict[str, Any]] = []
    filename = docx_path.name
    output_path: Optional[Path] = None
    conversion_report: Optional[CanadianConversionReport] = None
    audit_summary: Dict[str, int] = {}
    audit: Dict[str, Any] = {}
    numbering_checks: Dict[str, Any] = {}
    stage = "validation"

    try:
        conversion_mode = validate_conversion_mode(conversion_mode)
        with tempfile.TemporaryDirectory(prefix="phase2_") as tmp_root:
            digest = hashlib.sha256(str(docx_path.resolve()).encode("utf-8")).hexdigest()[:8]
            extract_dir_name = f"work_{digest}"

            stage = "extraction"
            per_file_log.append("Extracting DOCX...")
            with diag.timed(per_file_diag, "target", "extract"):
                decomposer = DocxDecomposer(str(docx_path))
                extract_dir = decomposer.extract(output_dir=Path(tmp_root) / extract_dir_name)

            stage = "bundle_build"
            per_file_log.append("Building slim bundle...")
            with diag.timed(per_file_diag, "target", "slim_bundle") as phase:
                bundle = build_phase2_slim_bundle(
                    extract_dir,
                    available_roles=available_roles,
                    role_specs=role_specs,
                )
                unresolved = len(bundle.get("paragraphs", []))
                deterministic = len(bundle.get("deterministic_classifications", []))
                phase.set(
                    unresolved=unresolved,
                    deterministic=deterministic,
                    deterministic_ignored=len(
                        bundle.get("deterministic_ignored_paragraphs", [])
                    ),
                )
            per_file_log.append(
                f"Built slim bundle: {unresolved} unresolved + {deterministic} deterministic"
            )

            stage = "classification_preflight"
            if unresolved > 0 and not api_key:
                raise ValueError("Anthropic API key is required when unresolved paragraphs exist.")

            stage = "classification"
            if unresolved:
                per_file_log.append("Classifying unresolved paragraphs with Anthropic...")
            else:
                per_file_log.append(
                    "All paragraphs classified deterministically; Anthropic request skipped"
                )
            with diag.timed(per_file_diag, "target", "classify") as phase:
                phase.set(unresolved_sent=unresolved, llm_used=bool(unresolved), model=model)
                classifications = classify_target_document(
                    slim_bundle=bundle,
                    available_roles=available_roles,
                    api_key=api_key,
                    model=model,
                )

            stage = "application"
            (
                output_path,
                conversion_report,
                audit_summary,
                audit,
                numbering_checks,
            ) = _apply_classified_target(
                docx_path=docx_path,
                extract_dir=extract_dir,
                bundle=bundle,
                classifications=classifications,
                arch_registry=arch_registry,
                env_registry=env_registry,
                arch_styles_xml=arch_styles_xml,
                output_dir=output_dir,
                log=per_file_log,
                source_tokens=source_tokens,
                arch_root=arch_root,
                role_specs=role_specs,
                conversion_mode=conversion_mode,
                diagnostics=per_file_diag,
            )

        return BatchResult(
            filename=filename,
            success=True,
            output_path=output_path,
            log=per_file_log,
            error=None,
            duration_seconds=time.monotonic() - start,
            conversion_report=conversion_report,
            audit_summary=audit_summary,
            audit=audit,
            numbering_checks=numbering_checks,
<<<<<<< HEAD
            stage="complete",
=======
            diagnostics=per_file_diag,
>>>>>>> 769bf0ca3a9a2744c852a1a23a7b3e5f88efb5b3
        )
    except Exception as exc:
        if isinstance(exc, ApplicationStageError):
            stage = exc.stage
            conversion_report = exc.conversion_report
            audit_summary = dict(exc.audit_summary)
            audit = dict(exc.audit)
            numbering_checks = dict(exc.numbering_checks)
        per_file_log.append(f"FAILED: {exc}")
        return BatchResult(
            filename=filename,
            success=False,
            output_path=output_path,
            log=per_file_log,
            error=str(exc),
            duration_seconds=time.monotonic() - start,
            conversion_report=conversion_report,
            audit_summary=audit_summary,
            audit=audit,
            numbering_checks=numbering_checks,
<<<<<<< HEAD
            stage=stage,
=======
            diagnostics=per_file_diag,
>>>>>>> 769bf0ca3a9a2744c852a1a23a7b3e5f88efb5b3
        )


def _prepare_file_for_batch(
    docx_path: Path,
    available_roles: List[str],
    extract_base_dir: Path,
    role_specs: Optional[Dict[str, Dict[str, Any]]] = None,
) -> PreparedFile:
    per_file_log: List[str] = []
    digest = hashlib.sha256(str(docx_path.resolve()).encode("utf-8")).hexdigest()[:8]
    extract_dir_name = f"work_{digest}"

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
    per_file_diag: List[Dict[str, Any]] = []
    output_path: Optional[Path] = None
    conversion_report: Optional[CanadianConversionReport] = None
    filename = prepared.docx_path.name
    audit_summary: Dict[str, int] = {}
    audit: Dict[str, Any] = {}
    numbering_checks: Dict[str, Any] = {}
    stage = "validation"

    try:
        conversion_mode = validate_conversion_mode(conversion_mode)
        stage = "application"
        (
            output_path,
            conversion_report,
            audit_summary,
            audit,
            numbering_checks,
        ) = _apply_classified_target(
            docx_path=prepared.docx_path,
            extract_dir=prepared.extract_dir,
            bundle=prepared.bundle,
            classifications=classifications,
            arch_registry=arch_registry,
            env_registry=env_registry,
            arch_styles_xml=arch_styles_xml,
            output_dir=output_dir,
            log=per_file_log,
            source_tokens=source_tokens,
            arch_root=arch_root,
            role_specs=role_specs,
            conversion_mode=conversion_mode,
            diagnostics=per_file_diag,
        )

        return BatchResult(
            filename=filename,
            success=True,
            output_path=output_path,
            log=per_file_log,
            error=None,
            duration_seconds=time.monotonic() - start,
            conversion_report=conversion_report,
            audit_summary=audit_summary,
            audit=audit,
            numbering_checks=numbering_checks,
<<<<<<< HEAD
            stage="complete",
=======
            diagnostics=per_file_diag,
>>>>>>> 769bf0ca3a9a2744c852a1a23a7b3e5f88efb5b3
        )
    except Exception as exc:
        if isinstance(exc, ApplicationStageError):
            stage = exc.stage
            conversion_report = exc.conversion_report
            audit_summary = dict(exc.audit_summary)
            audit = dict(exc.audit)
            numbering_checks = dict(exc.numbering_checks)
        per_file_log.append(f"FAILED: {exc}")
        return BatchResult(
            filename=filename,
            success=False,
            output_path=output_path,
            log=per_file_log,
            error=str(exc),
            duration_seconds=time.monotonic() - start,
            conversion_report=conversion_report,
            audit_summary=audit_summary,
            audit=audit,
            numbering_checks=numbering_checks,
<<<<<<< HEAD
            stage=stage,
=======
            diagnostics=per_file_diag,
>>>>>>> 769bf0ca3a9a2744c852a1a23a7b3e5f88efb5b3
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
