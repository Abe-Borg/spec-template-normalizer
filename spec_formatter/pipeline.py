"""One-call architect-template analysis and target-spec formatting.

The validated template bundle remains an internal integrity boundary.  Callers
provide the architect DOCX and target DOCX files; they never need to create,
locate, or transfer a bundle themselves.
"""

from __future__ import annotations

import hashlib
import json
import os
import queue
import re
import tempfile
import threading
import time
import uuid
import warnings
from concurrent.futures import FIRST_COMPLETED, Future, ThreadPoolExecutor, wait
from dataclasses import dataclass, field, replace
from datetime import datetime, timezone
from pathlib import Path
from typing import Any, Callable, Iterable, Mapping, Optional, Sequence

<<<<<<< HEAD
from . import __version__ as APPLICATION_VERSION
=======
from . import diagnostics as diag
>>>>>>> 769bf0ca3a9a2744c852a1a23a7b3e5f88efb5b3
from . import template_analysis
from .style_application.batch_runner import (
    BatchResult,
    SharedConfig,
    load_and_validate_shared_config,
    process_single_file,
)
from .style_application.core.application_policy import APPLICATION_POLICY_VERSION
from .style_application.core.csi_to_canadian import (
    CSI_TO_CANADIAN,
    FORMAT_ONLY,
    CanadianConversionReport,
    validate_conversion_mode,
)


ProgressCallback = Callable[[str], None]
ProgressEventCallback = Callable[[str, datetime], None]
TemplateClassifier = Callable[..., dict[str, Any]]
TemplateAnalyzer = Callable[..., template_analysis.Phase1Result]
TargetProcessor = Callable[..., BatchResult]

_FORMATTED_SUFFIXES = (
    "_FORMATTED.DOCX",
    "_CANADIAN_FORMATTED.DOCX",
    "_PHASE2_FORMATTED.DOCX",
)
_MAX_WORKERS = 6
_RUN_MANIFEST_VERSION = 2
_RUN_AUDIT_VERSION = 2
_PROFILE_CONTRACT_VERSION = "2"
_PROFILE_CACHE_NAMESPACE = f"contract-v{_PROFILE_CONTRACT_VERSION}"
_MAX_OUTPUT_COMPONENT_UTF16_UNITS = 240


def _empty_audit_summary() -> dict[str, int]:
    return {
        "styled": 0,
        "ignored": 0,
        "out_of_scope": 0,
        "unresolved": 0,
    }


@dataclass(frozen=True)
class TemplateProfile:
    """Validated internal template profile selected for a formatting run."""

    bundle_dir: Path
    source_sha256: str
    reused: bool


@dataclass(frozen=True)
class TargetFormatResult:
    """Outcome for one target specification."""

    source_path: Path
    success: bool
    output_path: Optional[Path]
    log: tuple[str, ...]
    error: Optional[str]
    duration_seconds: float
    conversion_report: Optional[CanadianConversionReport] = None
    source_sha256: Optional[str] = None
    output_sha256: Optional[str] = None
    audit_path: Optional[Path] = None
    audit_summary: dict[str, int] = field(default_factory=_empty_audit_summary)
    audit: dict[str, Any] = field(default_factory=dict)
    numbering_checks: dict[str, Any] = field(default_factory=dict)
<<<<<<< HEAD
    stage: Optional[str] = None
=======
    # Structured, redaction-safe phase-timing/count events for this target,
    # folded into the run-wide diagnostics recorder before publication.
    diagnostics: tuple[dict[str, Any], ...] = ()
>>>>>>> 769bf0ca3a9a2744c852a1a23a7b3e5f88efb5b3


@dataclass(frozen=True)
class FormatRunResult:
    """Consolidated result returned by :func:`format_specifications`."""

    template_profile: TemplateProfile
    output_dir: Path
    targets: tuple[TargetFormatResult, ...]
    run_id: str = ""
    conversion_mode: str = FORMAT_ONLY
    output_root: Optional[Path] = None
    run_dir: Optional[Path] = None
    manifest_path: Optional[Path] = None
    diagnostics_path: Optional[Path] = None

    def __post_init__(self) -> None:
        # Keep the historical ``output_dir`` attribute as a concrete alias of
        # the isolated run directory.  Defaults preserve compatibility for
        # callers that instantiate the result with the old three arguments.
        effective_run_dir = self.run_dir or self.output_dir
        object.__setattr__(self, "run_dir", effective_run_dir)
        object.__setattr__(self, "output_dir", effective_run_dir)
        if self.output_root is None:
            object.__setattr__(self, "output_root", effective_run_dir)

    @property
    def succeeded(self) -> int:
        return sum(1 for item in self.targets if item.success)

    @property
    def failed(self) -> int:
        return len(self.targets) - self.succeeded

    @property
    def success(self) -> bool:
        return bool(self.targets) and self.failed == 0

    @property
    def output_paths(self) -> tuple[Path, ...]:
        return tuple(
            item.output_path
            for item in self.targets
            if item.success and item.output_path is not None
        )


@dataclass(frozen=True)
class SafeErrorDiagnostic:
    """Persistable error identity that cannot contain document/model text."""

    code: str
    message: str


def _attach_safe_error_diagnostic(
    error: Exception,
    *,
    code: str,
    message: str,
) -> Exception:
    """Attach path-free public identity while preserving the exception type."""

    setattr(error, "safe_error_code", code)
    setattr(error, "safe_error_message", message)
    return error


def _emit(progress: Optional[ProgressCallback], message: str) -> None:
    if progress is None:
        return
    try:
        progress(message)
    except Exception as exc:
        warnings.warn(
            f"Formatting progress callback failed and was ignored: {exc}",
            RuntimeWarning,
            stacklevel=2,
        )


def _emit_progress_event(
    progress_event: Optional[ProgressEventCallback],
    message: str,
    occurred_at: datetime,
) -> None:
    """Emit a timestamped progress event without changing the legacy callback."""

    if progress_event is None:
        return
    try:
        progress_event(message, occurred_at)
    except Exception as exc:
        warnings.warn(
            f"Formatting progress event callback failed and was ignored: {exc}",
            RuntimeWarning,
            stacklevel=2,
        )


def _friendly_template_progress(message: str) -> str:
    """Translate engine-level template messages into product language."""

    lowered = message.casefold()
    if lowered.startswith("snapshotting"):
        return "Checking the architect template..."
    if "unpacking" in lowered:
        return "Reading the architect template..."
    if lowered.startswith("reading paragraph"):
        return "Identifying the template's structure..."
    if lowered.startswith("classifying"):
        return "Analyzing the template's paragraph roles..."
    if lowered.startswith("deriving portable styles"):
        return "Building the reusable formatting profile..."
    if lowered.startswith("capturing the source formatting"):
        return "Capturing fonts, numbering, headers, and page layout..."
    if lowered.startswith("validating checksums"):
        return "Validating the architect formatting profile..."
    if lowered.startswith("published validated bundle"):
        return "Architect template analysis complete."
    if lowered.startswith("warning: private work directory remains"):
        return "Warning: temporary template-analysis files could not be removed automatically."
    return message


def _is_formatted_output(path: Path) -> bool:
    return path.name.upper().endswith(_FORMATTED_SUFFIXES)


def collect_target_specs(
    inputs: Iterable[Path],
    *,
    exclude_discovered: Optional[Path] = None,
) -> tuple[Path, ...]:
    """Expand DOCX files and folders into a stable, deduplicated target list.

    Folder discovery is intentionally non-recursive and excludes Word lock
    files plus outputs from current and legacy versions of the formatter.
    ``exclude_discovered`` is ignored only during folder expansion; an
    explicitly supplied matching file remains in the result so input
    validation can reject selecting the architect as a target.
    """

    excluded_key = (
        os.path.normcase(str(Path(exclude_discovered).expanduser().resolve()))
        if exclude_discovered is not None
        else None
    )
    discovered: list[Path] = []
    for raw_path in inputs:
        path = Path(raw_path).expanduser()
        if path.is_dir():
            discovered.extend(
                candidate
                for candidate in path.glob("*.docx")
                if not candidate.name.startswith("~$")
                and not _is_formatted_output(candidate)
                and os.path.normcase(str(candidate.resolve())) != excluded_key
            )
        else:
            discovered.append(path)

    unique: dict[str, Path] = {}
    for path in discovered:
        resolved = path.resolve()
        key = os.path.normcase(str(resolved))
        unique.setdefault(key, resolved)
    return tuple(sorted(unique.values(), key=lambda item: (item.name.casefold(), str(item))))


def default_template_cache_dir() -> Path:
    """Return the per-user cache location used by the GUI and headless API."""

    local_app_data = os.environ.get("LOCALAPPDATA")
    if local_app_data:
        return Path(local_app_data) / "SpecificationFormatter" / "TemplateCache"
    xdg_cache = os.environ.get("XDG_CACHE_HOME")
    if xdg_cache:
        return Path(xdg_cache) / "specification-formatter" / "template-cache"
    return Path.home() / ".cache" / "specification-formatter" / "template-cache"


def _validate_inputs(
    architect_template: Path,
    target_specs: Sequence[Path],
    output_dir: Path,
) -> tuple[Path, tuple[Path, ...], Path]:
    architect = Path(architect_template).expanduser().resolve()
    if not architect.is_file():
        raise _attach_safe_error_diagnostic(
            FileNotFoundError(f"Architect template does not exist: {architect}"),
            code="input_architect_missing",
            message="Architect template does not exist.",
        )
    if architect.suffix.lower() != ".docx":
        raise _attach_safe_error_diagnostic(
            ValueError(f"Architect template must be a .docx file: {architect}"),
            code="input_architect_not_docx",
            message="Architect template must be a .docx file.",
        )
    if architect.name.startswith("~$"):
        message = "Select the saved architect DOCX, not Word's temporary lock file."
        raise _attach_safe_error_diagnostic(
            ValueError(message),
            code="input_architect_lock_file",
            message=message,
        )

    targets = collect_target_specs(target_specs, exclude_discovered=architect)
    if not targets:
        message = "Select at least one target specification DOCX file."
        raise _attach_safe_error_diagnostic(
            ValueError(message),
            code="input_target_required",
            message=message,
        )

    architect_key = os.path.normcase(str(architect))
    for target in targets:
        if not target.is_file():
            raise _attach_safe_error_diagnostic(
                FileNotFoundError(
                    f"Target specification does not exist: {target}"
                ),
                code="input_target_missing",
                message="A selected target specification does not exist.",
            )
        if target.suffix.lower() != ".docx":
            raise _attach_safe_error_diagnostic(
                ValueError(
                    f"Target specification must be a .docx file: {target}"
                ),
                code="input_target_not_docx",
                message="Every target specification must be a .docx file.",
            )
        if target.name.startswith("~$"):
            raise _attach_safe_error_diagnostic(
                ValueError(f"Target is a Word temporary lock file: {target}"),
                code="input_target_lock_file",
                message="A selected target is a Word temporary lock file.",
            )
        if _is_formatted_output(target):
            raise _attach_safe_error_diagnostic(
                ValueError(f"Target is already a formatted output: {target}"),
                code="input_target_already_formatted",
                message="A selected target is already a formatted output.",
            )
        if os.path.normcase(str(target)) == architect_key:
            message = "The architect template cannot also be a target specification."
            raise _attach_safe_error_diagnostic(
                ValueError(message),
                code="input_architect_is_target",
                message=message,
            )

    destination = Path(output_dir).expanduser().resolve()
    try:
        destination.mkdir(parents=True, exist_ok=True)
    except FileExistsError as exc:
        raise _attach_safe_error_diagnostic(
            NotADirectoryError(
                f"Output location is not a directory: {destination}"
            ),
            code="output_not_directory",
            message="Output location is not a directory.",
        ) from exc
    except OSError as exc:
        message = "Output directory could not be created."
        raise _attach_safe_error_diagnostic(
            OSError(message),
            code="output_create_failed",
            message=message,
        ) from exc
    if not destination.is_dir():
        raise _attach_safe_error_diagnostic(
            NotADirectoryError(
                f"Output location is not a directory: {destination}"
            ),
            code="output_not_directory",
            message="Output location is not a directory.",
        )
    try:
        with tempfile.NamedTemporaryFile(
            prefix=".spec-formatter-write-check-",
            dir=destination,
        ):
            pass
    except OSError as exc:
        raise _attach_safe_error_diagnostic(
            PermissionError(f"Output directory is not writable: {destination}"),
            code="output_not_writable",
            message="Output directory is not writable.",
        ) from exc
    return architect, targets, destination


def _manifest_matches_current_engine(
    manifest: template_analysis.BundleManifest,
    *,
    model: str,
    prompt_dir: Path,
    classifier: Optional[TemplateClassifier],
) -> bool:
    producer = manifest.producer
    expected_provider = "injected" if classifier is not None else "anthropic"
    expected_model = (
        getattr(classifier, "__name__", "custom-classifier")
        if classifier is not None
        else model
    )
    prompt_hashes = {
        "master_prompt_sha256": hashlib.sha256(
            (prompt_dir / "master_prompt.txt").read_text(encoding="utf-8").encode("utf-8")
        ).hexdigest(),
        "run_instruction_sha256": hashlib.sha256(
            (prompt_dir / "run_instruction_prompt.txt")
            .read_text(encoding="utf-8")
            .encode("utf-8")
        ).hexdigest(),
    }
    return (
        producer.get("name") == "spec-template-normalizer"
        and producer.get("version") == template_analysis.PIPELINE_VERSION
        and producer.get("classifier")
        == {"provider": expected_provider, "model": expected_model}
        and producer.get("prompts") == prompt_hashes
    )


def _stable_source_sha256(path: Path) -> str:
    before = path.stat()
    digest = template_analysis.sha256_file(path)
    after = path.stat()
    if (before.st_size, before.st_mtime_ns) != (after.st_size, after.st_mtime_ns):
        raise RuntimeError(
            "The architect template changed while it was being checked. "
            "Finish saving it and run again."
        )
    return digest


def _snapshot_input(source: Path, destination: Path) -> str:
    """Copy one stable input version and return its SHA-256 digest."""

    before = source.stat()
    destination.parent.mkdir(parents=True, exist_ok=True)
    copied_digest = hashlib.sha256()
    try:
        with source.open("rb") as reader, destination.open("xb") as writer:
            while chunk := reader.read(1024 * 1024):
                writer.write(chunk)
                copied_digest.update(chunk)
    except Exception:
        destination.unlink(missing_ok=True)
        raise

    after = source.stat()
    source_digest = template_analysis.sha256_file(source)
    final = source.stat()
    if (
        (before.st_size, before.st_mtime_ns) != (after.st_size, after.st_mtime_ns)
        or (after.st_size, after.st_mtime_ns) != (final.st_size, final.st_mtime_ns)
        or copied_digest.hexdigest() != source_digest
    ):
        destination.unlink(missing_ok=True)
        raise RuntimeError(
            f"{source.name} changed while it was being snapshotted. "
            "Finish saving it and run again."
        )
    return copied_digest.hexdigest()


def _find_cached_profile(
    cache_dir: Path,
    source_sha256: str,
    progress: Optional[ProgressCallback],
    *,
    model: str,
    prompt_dir: Path,
    classifier: Optional[TemplateClassifier],
) -> Optional[TemplateProfile]:
    if not cache_dir.is_dir():
        return None
    pattern = f"*--{source_sha256[:12]}--*.phase1"
    candidates = sorted(
        cache_dir.glob(pattern),
        key=lambda item: item.stat().st_mtime_ns,
        reverse=True,
    )
    for candidate in candidates:
        try:
            manifest = template_analysis.validate_bundle_directory(
                candidate,
                expected_source_sha256=source_sha256,
            )
            if not _manifest_matches_current_engine(
                manifest,
                model=model,
                prompt_dir=prompt_dir,
                classifier=classifier,
            ):
                continue
            return TemplateProfile(candidate, source_sha256, reused=True)
        except Exception as exc:
            _emit(
                progress,
                f"Ignoring an invalid cached template profile ({candidate.name}): {exc}",
            )
    return None


def prepare_template_profile(
    architect_template: Path,
    cache_dir: Path,
    api_key: str,
    *,
    force_analysis: bool = False,
    model: str = template_analysis.DEFAULT_MODEL,
    prompt_dir: Optional[Path] = None,
    progress: Optional[ProgressCallback] = None,
    classifier: Optional[TemplateClassifier] = None,
    analyzer: TemplateAnalyzer = template_analysis.run_phase1,
) -> TemplateProfile:
    """Return a current, strictly validated profile for *architect_template*."""

    architect = Path(architect_template).resolve()
    cache_root = Path(cache_dir).resolve() / _PROFILE_CACHE_NAMESPACE
    source_sha256 = _stable_source_sha256(architect)
    effective_prompt_dir = (
        Path(prompt_dir).resolve()
        if prompt_dir is not None
        else Path(__file__).resolve().parents[1]
    )

    # Injected classifiers are primarily an offline/test extension. Their
    # implementation identity is not captured strongly enough for safe reuse.
    if not force_analysis and classifier is None:
        cached = _find_cached_profile(
            cache_root,
            source_sha256,
            progress,
            model=model,
            prompt_dir=effective_prompt_dir,
            classifier=classifier,
        )
        if cached is not None:
            _emit(progress, "Reusing the validated architect template analysis.")
            return cached

    if classifier is None and not isinstance(api_key, str):
        raise ValueError("Anthropic API key must be text.")
    if classifier is None and not api_key.strip():
        raise ValueError(
            "An Anthropic API key is required to analyze a new architect template."
        )

    cache_root.mkdir(parents=True, exist_ok=True)
    _emit(progress, "Analyzing the architect template...")
    analyzer_kwargs: dict[str, Any] = {
        "source_docx": architect,
        "output_root": cache_root,
        "api_key": api_key,
        "model": model,
        "progress": lambda message: _emit(
            progress,
            _friendly_template_progress(message),
        ),
    }
    analyzer_kwargs["prompt_dir"] = effective_prompt_dir
    if classifier is not None:
        analyzer_kwargs["classifier"] = classifier
    phase1_result = analyzer(**analyzer_kwargs)
    manifest = template_analysis.validate_bundle_directory(
        phase1_result.bundle_dir,
        expected_source_sha256=source_sha256,
    )
    if not _manifest_matches_current_engine(
        manifest,
        model=model,
        prompt_dir=effective_prompt_dir,
        classifier=classifier,
    ):
        raise ValueError("Template analysis produced an incompatible profile bundle.")
    return TemplateProfile(phase1_result.bundle_dir, source_sha256, reused=False)


def _safe_filename_fragment(value: str) -> str:
    cleaned = re.sub(r"[^A-Za-z0-9_-]+", "-", value).strip("-_")
    return cleaned[:32] or "source"


def _utc_now() -> datetime:
    return datetime.now(timezone.utc)


def _iso_utc(value: datetime) -> str:
    return value.astimezone(timezone.utc).isoformat().replace("+00:00", "Z")


def _create_run_directory(output_root: Path, conversion_mode: str) -> tuple[str, Path]:
    """Create and return a collision-resistant, human-sortable run directory."""

    stamp = _utc_now().strftime("%Y%m%dT%H%M%S.%fZ")
    mode = _safe_filename_fragment(conversion_mode)
    for _attempt in range(10):
        run_id = uuid.uuid4().hex[:12]
        run_dir = output_root / f"{stamp}_{mode}_{run_id}"
        try:
            run_dir.mkdir(parents=False, exist_ok=False)
            return run_id, run_dir
        except FileExistsError:  # pragma: no cover - UUID collision defense
            continue
    raise FileExistsError("Could not allocate a unique formatter run directory.")


def _atomic_write_bytes(path: Path, payload: bytes) -> None:
    """Durably stage *payload* beside *path*, then publish it atomically."""

    path.parent.mkdir(parents=True, exist_ok=True)
    partial = path.parent / f".meta-{uuid.uuid4().hex[:12]}.tmp"
    try:
        with partial.open("xb") as writer:
            writer.write(payload)
            writer.flush()
            os.fsync(writer.fileno())
        os.replace(partial, path)
    except Exception:
        partial.unlink(missing_ok=True)
        raise


def _atomic_write_json(path: Path, payload: Mapping[str, Any]) -> None:
    encoded = json.dumps(
        payload,
        indent=2,
        sort_keys=True,
        ensure_ascii=False,
    ).encode("utf-8") + b"\n"
    _atomic_write_bytes(path, encoded)


def _publish_output(source: Path, destination: Path) -> str:
    """Copy from short staging, then atomically publish inside the run folder."""

    destination.parent.mkdir(parents=True, exist_ok=True)
    partial = destination.parent / f".publish-{uuid.uuid4().hex[:12]}.tmp.docx"
    digest = hashlib.sha256()
    try:
        with source.open("rb") as reader, partial.open("xb") as writer:
            while chunk := reader.read(1024 * 1024):
                writer.write(chunk)
                digest.update(chunk)
            writer.flush()
            os.fsync(writer.fileno())
        os.replace(partial, destination)
        published_digest = _stable_source_sha256(destination)
        if published_digest != digest.hexdigest():
            destination.unlink(missing_ok=True)
            raise RuntimeError("Published output checksum does not match staged output.")
        source.unlink(missing_ok=True)
        return published_digest
    except Exception:
        partial.unlink(missing_ok=True)
        raise


def _normalize_audit_summary(value: Any) -> dict[str, int]:
    summary = _empty_audit_summary()
    if not isinstance(value, Mapping):
        return summary
    for key in summary:
        count = value.get(key, 0)
        if isinstance(count, int) and not isinstance(count, bool) and count >= 0:
            summary[key] = count
    return summary


def _normalize_numbering_checks(value: Any) -> dict[str, Any]:
    if not isinstance(value, Mapping):
        return {}
    normalized: dict[str, Any] = {}
    for key, item in value.items():
        if not isinstance(key, str):
            continue
        if item is None or isinstance(item, (int, float, bool)):
            normalized[key] = item
        elif (
            isinstance(item, str)
            and key in {"conversion_mode", "policy", "status"}
            and re.fullmatch(r"[A-Za-z0-9_.:+\-]{1,80}", item) is not None
        ):
            normalized[key] = item
        elif isinstance(item, (list, tuple)) and all(
            element is None or isinstance(element, (int, float, bool))
            for element in item
        ):
            normalized[key] = list(item)
    return normalized


def _normalize_audit_details(value: Any) -> dict[str, Any]:
    """Keep JSON-safe, text-free structural audit data from the processor."""

    if not isinstance(value, Mapping):
        return {}

    safe_string_keys = frozenset(
        {
            "category",
            "code",
            "conversion_mode",
            "csi_role",
            "disposition",
            "kind",
            "method",
            "numbering_provenance",
            "policy",
            "property",
            "reason",
            "role",
            "source_kind",
            "status",
            "style_id",
            "target_kind",
        }
    )
    safe_reason_codes = frozenset(
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
    omitted = object()

    def normalize(item: Any, *, key: Optional[str] = None) -> Any:
        if item is None or isinstance(item, (bool, int, float)):
            return item
        if isinstance(item, str):
            # Only contract-defined structural identifiers may reach disk.
            # Unknown free-form strings can be model-authored and therefore
            # may echo document text even when their containing key sounds
            # harmless (for example an LLM-authored ``reason``).
            if key not in safe_string_keys:
                return omitted
            if key == "reason" and item not in safe_reason_codes:
                return "unspecified"
            if re.fullmatch(r"[A-Za-z0-9_.:/#@+\-]{1,160}", item) is None:
                return omitted
            return item
        if isinstance(item, Mapping):
            normalized: dict[str, Any] = {}
            for nested_key, nested in item.items():
                if not isinstance(nested_key, str) or _audit_key_may_contain_document_text(
                    nested_key
                ):
                    continue
                normalized_value = normalize(nested, key=nested_key)
                if normalized_value is not omitted:
                    normalized[nested_key] = normalized_value
            return normalized
        if isinstance(item, (list, tuple)):
            normalized_items = [normalize(nested, key=key) for nested in item]
            return [nested for nested in normalized_items if nested is not omitted]
        return omitted

    normalized_root = normalize(value)
    return normalized_root if isinstance(normalized_root, dict) else {}


def _audit_key_may_contain_document_text(key: str) -> bool:
    folded = key.casefold()
    return any(
        token in folded
        for token in ("text", "preview", "excerpt", "content", "paragraph_xml")
    )


def _sha256_text_file(path: Path) -> Optional[str]:
    if not path.is_file():
        return None
    return hashlib.sha256(path.read_bytes()).hexdigest()


def _target_prompt_fingerprints() -> dict[str, str]:
    prompt_root = Path(__file__).parent / "style_application" / "core" / "prompts"
    fingerprints: dict[str, str] = {}
    for filename in ("phase2_master_prompt.txt", "phase2_run_instruction.txt"):
        digest = _sha256_text_file(prompt_root / filename)
        if digest is not None:
            fingerprints[f"{Path(filename).stem}_sha256"] = digest
    return fingerprints


def _redact(value: Optional[str], secrets: Sequence[str]) -> Optional[str]:
    if value is None:
        return None
    redacted = value
    for secret in secrets:
        if secret:
            redacted = redacted.replace(secret, "[REDACTED]")
    return redacted


_DOCUMENT_DATA_FIELD_RX = re.compile(
    r"(?i)\b(?:content|excerpt|paragraph(?:_xml)?|preview|text)\s*[:=]"
)
_OOXML_FRAGMENT_RX = re.compile(
    r"<\??/?(?:a|m|mc|o|pic|r|v|w|w10|w14|wp|wps):|<\?xml",
    re.IGNORECASE,
)
_EVENT_TIMESTAMP_RX = re.compile(
    r"^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}(?:\.\d+)?Z\s+"
)
_TARGET_EVENT_RX = re.compile(r"^(Target .+?: )(.*)$")
_SAFE_REDACTION_RX = re.compile(
    r"^\[(?:document content omitted|untrusted detail omitted; sha256=[0-9a-f]{12})\]$"
)

_KNOWN_SAFE_ERROR_MESSAGES = {
    "template_section_shell_conflict": (
        "Architect template has conflicting section shells; use one canonical "
        "page layout and default/even/first header-footer mapping."
    ),
    "template_default_section_conflict": (
        "Architect template default section conflicts with its section chain."
    ),
    "template_duplicate_section_index": (
        "Architect template section chain has duplicate section_index values."
    ),
    "template_api_key_required": (
        "An Anthropic API key is required to analyze a new architect template."
    ),
    "target_api_key_required": (
        "Anthropic API key is required when unresolved paragraphs exist."
    ),
    "target_output_missing": (
        "Style application reported success without an output DOCX."
    ),
    "target_formatting_failed": "Target formatting failed.",
    "input_architect_lock_file": (
        "Select the saved architect DOCX, not Word's temporary lock file."
    ),
    "input_target_required": (
        "Select at least one target specification DOCX file."
    ),
    "input_architect_is_target": (
        "The architect template cannot also be a target specification."
    ),
    "invalid_conversion_mode": (
        "conversion_mode must be one of: csi_to_canadian, format_only"
    ),
    "invalid_max_workers": "max_workers must be an integer.",
    "output_create_failed": "Output directory could not be created.",
    "api_key_invalid_type": "Anthropic API key must be text.",
}

# These internal validation errors append a user-supplied filesystem path to
# the exception for direct Python callers. Public diagnostics retain only the
# stable code and path-free remediation text below.
_PATH_BEARING_SAFE_ERROR_PREFIXES = (
    (
        "Architect template does not exist:",
        SafeErrorDiagnostic(
            code="input_architect_missing",
            message="Architect template does not exist.",
        ),
    ),
    (
        "Architect template must be a .docx file:",
        SafeErrorDiagnostic(
            code="input_architect_not_docx",
            message="Architect template must be a .docx file.",
        ),
    ),
    (
        "Target specification does not exist:",
        SafeErrorDiagnostic(
            code="input_target_missing",
            message="A selected target specification does not exist.",
        ),
    ),
    (
        "Target specification must be a .docx file:",
        SafeErrorDiagnostic(
            code="input_target_not_docx",
            message="Every target specification must be a .docx file.",
        ),
    ),
    (
        "Target is a Word temporary lock file:",
        SafeErrorDiagnostic(
            code="input_target_lock_file",
            message="A selected target is a Word temporary lock file.",
        ),
    ),
    (
        "Target is already a formatted output:",
        SafeErrorDiagnostic(
            code="input_target_already_formatted",
            message="A selected target is already a formatted output.",
        ),
    ),
    (
        "Output location is not a directory:",
        SafeErrorDiagnostic(
            code="output_not_directory",
            message="Output location is not a directory.",
        ),
    ),
    (
        "Output directory is not writable:",
        SafeErrorDiagnostic(
            code="output_not_writable",
            message="Output directory is not writable.",
        ),
    ),
)

_WRAPPED_SAFE_ERROR_CODES = frozenset(
    {
        "template_section_shell_conflict",
        "template_default_section_conflict",
        "template_duplicate_section_index",
    }
)


def _known_safe_error_diagnostic(value: Optional[str]) -> Optional[SafeErrorDiagnostic]:
    """Recognize exact internal diagnostics, including safe wrapped variants."""

    if not value:
        return None
    for prefix, diagnostic in _PATH_BEARING_SAFE_ERROR_PREFIXES:
        if value.startswith(prefix) and value[len(prefix) :].strip():
            return diagnostic
    for code, message in _KNOWN_SAFE_ERROR_MESSAGES.items():
        # The preflight layer may add stable context around a lower-level
        # diagnostic.  Persist only the canonical marker, never that wrapper.
        if value == message or (
            code in _WRAPPED_SAFE_ERROR_CODES and message in value
        ):
            return SafeErrorDiagnostic(code=code, message=message)
    return None


_SAFE_OPERATIONAL_PREFIXES = (
    "Added ",
    "All architect fonts ",
    "Analyzing the architect template",
    "Applied ",
    "Application coverage:",
    "BEGIN ENVIRONMENT APPLICATION",
    "Building slim bundle",
    "Checking input files",
    "Checking the architect template",
    "Classification coverage:",
    "Classification checkpoint saved",
    "Classifying ",
    "Canadian conversion:",
    "Complete:",
    "Converting CSI hierarchy",
    "Created ",
    "END ENVIRONMENT APPLICATION",
    "Extracting DOCX",
    "Failed ",
    "Formatted ",
    "Identifying the template",
    "Imported ",
    "Importing ",
    "Inserted ",
    "Left ",
    "Namespaced architect style",
    "No architect ",
    "No docDefaults ",
    "No fontTable ",
    "No numbering ",
    "No theme ",
    "No token matches ",
    "Output:",
    "Patched sectPr[",
    "Patched tokens ",
    "Processing target ",
    "Queued ",
    "Reading the architect template",
    "Rebuilt ",
    "Remapped ",
    "Removed old ",
    "Replaced ",
    "Replacing ",
    "Reusing the validated architect template analysis",
    "Rewired ",
    "Skipped ",
    "Started ",
    "Stripped ",
    "Suppressed ",
    "Target has no numbering",
    "Updated ",
    "Validating the architect formatting profile",
    "Validating the template profile",
    "Wrote ",
)


def _is_safe_operational_line(line: str) -> bool:
    candidate = re.sub(
        r"^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}(?:\.\d+)?Z\s+",
        "",
        line.strip(),
    )
    if not candidate:
        return True
    if _SAFE_REDACTION_RX.fullmatch(candidate):
        return True
    if re.fullmatch(r"=+", candidate) or re.match(r"^\[\d+/\d+\] ", candidate):
        return True
    if re.match(r"^numId \d+ -> \d+ \(abstractNum \d+ -> \d+\)$", candidate):
        return True
    return candidate.startswith(_SAFE_OPERATIONAL_PREFIXES)


def _sanitize_run_text(
    value: Optional[str],
    secrets: Sequence[str],
    *,
    allow_operational: bool = False,
) -> Optional[str]:
    """Redact secrets and reject payload-shaped document text from run artifacts.

    Target logs remain line-for-line useful for normal operational messages.
    A line that looks like OOXML or an explicit text/content field is replaced
    wholesale because retaining the surrounding exception is not worth
    persisting body text from a customer document.
    """

    redacted = _redact(value, secrets)
    if redacted is None:
        return None
    safe_lines: list[str] = []
    for raw_line in redacted.splitlines() or [redacted]:
        bounded = raw_line[:4096]
        timestamp_match = _EVENT_TIMESTAMP_RX.match(bounded)
        timestamp = timestamp_match.group(0) if timestamp_match else ""
        candidate = bounded[len(timestamp) :]

        target_match = _TARGET_EVENT_RX.fullmatch(candidate)
        if target_match is not None:
            detail = _sanitize_run_text(
                target_match.group(2),
                secrets,
                allow_operational=allow_operational,
            )
            safe_lines.append(f"{timestamp}{target_match.group(1)}{detail or ''}")
            continue

        if _OOXML_FRAGMENT_RX.search(candidate) or _DOCUMENT_DATA_FIELD_RX.search(
            candidate
        ):
            safe_lines.append(f"{timestamp}[document content omitted]")
            continue

        known = _known_safe_error_diagnostic(candidate)
        if known is not None:
            safe_lines.append(f"{timestamp}ERROR [{known.code}]: {known.message}")
            continue

        if allow_operational and _is_safe_operational_line(candidate):
            safe_lines.append(bounded)
            continue
        fingerprint = hashlib.sha256(candidate.encode("utf-8")).hexdigest()[:12]
        safe_lines.append(
            f"{timestamp}[untrusted detail omitted; sha256={fingerprint}]"
        )
    return "\n".join(safe_lines)


def safe_error_diagnostic(
    value: Optional[str | BaseException],
    secrets: Sequence[str] = (),
) -> Optional[SafeErrorDiagnostic]:
    """Return a stable safe error code/message for artifacts and user summaries."""

    if value is None:
        return None
    if isinstance(value, BaseException):
        code = getattr(value, "safe_error_code", None)
        message = getattr(value, "safe_error_message", None)
        if isinstance(code, str) and isinstance(message, str):
            return SafeErrorDiagnostic(code=code, message=message)
        raw_value = str(value)
    else:
        raw_value = value
    known = _known_safe_error_diagnostic(raw_value)
    if known is not None:
        return known
    return SafeErrorDiagnostic(
        code="untrusted_error",
        message=(
            _sanitize_run_text(raw_value, secrets)
            or "[untrusted detail omitted]"
        ),
    )


def _plan_output_paths(
    targets: Sequence[Path],
    output_dir: Path,
    conversion_mode: str = FORMAT_ONLY,
) -> dict[Path, Path]:
    conversion_mode = validate_conversion_mode(conversion_mode)
    suffix = "_CANADIAN_FORMATTED.docx" if conversion_mode == CSI_TO_CANADIAN else "_FORMATTED.docx"
    stem_counts: dict[str, int] = {}
    for target in targets:
        key = target.stem.casefold()
        stem_counts[key] = stem_counts.get(key, 0) + 1

    planned: dict[Path, Path] = {}
    used_names: set[str] = set()
    for target in targets:
        stem = target.stem
        if stem_counts[stem.casefold()] == 1:
            proposed = f"{stem}{suffix}"
        else:
            parent = _safe_filename_fragment(target.parent.name)
            digest = hashlib.sha1(str(target).encode("utf-8")).hexdigest()[:8]
            proposed = f"{stem}__{parent}-{digest}{suffix}"
        filename = _bounded_output_component(proposed, suffix=suffix)
        folded = filename.casefold()
        if folded in used_names:
            raise ValueError(f"Could not create unique output name for {target}")
        used_names.add(folded)
        planned[target] = output_dir / filename
    return planned


def _utf16_code_units(value: str) -> int:
    return len(value.encode("utf-16-le")) // 2


def _truncate_utf16(value: str, max_units: int) -> str:
    if max_units < 0:
        raise ValueError("max_units must be non-negative")
    kept: list[str] = []
    used = 0
    for character in value:
        units = _utf16_code_units(character)
        if used + units > max_units:
            break
        kept.append(character)
        used += units
    return "".join(kept)


def _bounded_output_component(proposed: str, *, suffix: str) -> str:
    """Return a deterministic Windows-safe output filename component."""

    if _utf16_code_units(proposed) <= _MAX_OUTPUT_COMPONENT_UTF16_UNITS:
        return proposed
    digest = hashlib.sha256(proposed.encode("utf-8")).hexdigest()[:12]
    tail = f"__{digest}{suffix}"
    available = _MAX_OUTPUT_COMPONENT_UTF16_UNITS - _utf16_code_units(tail)
    if available <= 0:  # pragma: no cover - fixed formatter suffixes are short
        raise ValueError("Formatted output suffix is too long for Windows.")
    stem = proposed[: -len(suffix)] if proposed.endswith(suffix) else proposed
    bounded = f"{_truncate_utf16(stem, available)}{tail}"
    if _utf16_code_units(bounded) > _MAX_OUTPUT_COMPONENT_UTF16_UNITS:
        raise AssertionError("Output filename bound calculation failed")
    return bounded


def _validate_output_plan(
    architect: Path,
    targets: Sequence[Path],
    planned_outputs: dict[Path, Path],
) -> None:
    inputs = (architect, *targets)
    input_by_key = {os.path.normcase(str(path)): path for path in inputs}
    for source, output in planned_outputs.items():
        conflicting_input = input_by_key.get(os.path.normcase(str(output)))
        if conflicting_input is not None:
            raise ValueError(
                f"Formatted output for {source.name} would overwrite an input file: "
                f"{conflicting_input}"
            )


def _format_one_target(
    target: Path,
    final_output: Path,
    staging_dir: Path,
    shared: SharedConfig,
    api_key: str,
    model: str,
    processor: TargetProcessor,
    conversion_mode: str,
    on_started: Optional[Callable[[], None]] = None,
) -> TargetFormatResult:
    start = time.monotonic()
    processor_log: tuple[str, ...] = ()
    conversion_report: Optional[CanadianConversionReport] = None
    snapshot_sha256: Optional[str] = None
    audit_summary = _empty_audit_summary()
    audit: dict[str, Any] = {}
    numbering_checks: dict[str, Any] = {}
<<<<<<< HEAD
    stage: Optional[str] = None
=======
    # Structured pipeline-side phase events (snapshot/publish) interleaved with
    # the engine's own phase events.  Carries counts/timings only, never text.
    diag_events: list[dict[str, Any]] = []
>>>>>>> 769bf0ca3a9a2744c852a1a23a7b3e5f88efb5b3
    try:
        stage = "processing"
        if on_started is not None:
            on_started()
        # Deliberately avoid carrying user-controlled filenames into the work
        # tree.  This keeps Windows paths short even for deeply nested inputs.
        snapshot = staging_dir / "source.docx"
        with diag.timed(diag_events, "target", "snapshot"):
            snapshot_sha256 = _snapshot_input(target, snapshot)
        result = processor(
            docx_path=snapshot,
            arch_registry=shared.arch_registry,
            env_registry=shared.env_registry,
            arch_styles_xml=shared.arch_styles_xml,
            available_roles=shared.available_roles,
            api_key=api_key,
            output_dir=staging_dir / "output",
            source_tokens=shared.source_tokens,
            arch_root=shared.arch_root,
            model=model,
            role_specs=shared.role_specs,
            conversion_mode=conversion_mode,
        )
<<<<<<< HEAD
        # The processor writes into an isolated staging directory. Its final
        # path diagnostics therefore point at files that are deleted when the
        # job temp directory closes. The public result already carries the
        # durable published path, so never replay or persist staging paths.
        staging_marker = os.path.normcase(str(staging_dir))
        processor_log = tuple(
            line
            for line in result.log
            if not str(line).lstrip().startswith("Output:")
            and staging_marker not in os.path.normcase(str(line))
        )
=======
        processor_log = tuple(result.log)
        diag_events.extend(getattr(result, "diagnostics", None) or [])
>>>>>>> 769bf0ca3a9a2744c852a1a23a7b3e5f88efb5b3
        conversion_report = result.conversion_report
        audit_summary = _normalize_audit_summary(
            getattr(result, "audit_summary", None)
        )
        audit = _normalize_audit_details(getattr(result, "audit", None))
        numbering_checks = _normalize_numbering_checks(
            getattr(result, "numbering_checks", None)
        )
        stage = getattr(result, "stage", None)
        if not result.success:
            return TargetFormatResult(
                source_path=target,
                success=False,
                output_path=None,
                log=processor_log,
                error=result.error or "Target formatting failed.",
                duration_seconds=result.duration_seconds,
                conversion_report=conversion_report,
                source_sha256=snapshot_sha256,
                audit_summary=audit_summary,
                audit=audit,
                numbering_checks=numbering_checks,
<<<<<<< HEAD
                stage=stage,
=======
                diagnostics=tuple(diag_events),
>>>>>>> 769bf0ca3a9a2744c852a1a23a7b3e5f88efb5b3
            )
        stage = "publication"
        if result.output_path is None or not result.output_path.is_file():
            raise RuntimeError("Style application reported success without an output DOCX.")
        if _stable_source_sha256(target) != snapshot_sha256:
            raise RuntimeError(
                f"{target.name} changed during formatting. Finish saving it and run again."
            )
        with diag.timed(diag_events, "target", "publish"):
            output_sha256 = _publish_output(result.output_path, final_output)
        return TargetFormatResult(
            source_path=target,
            success=True,
            output_path=final_output,
            log=processor_log,
            error=None,
            duration_seconds=result.duration_seconds,
            conversion_report=conversion_report,
            source_sha256=snapshot_sha256,
            output_sha256=output_sha256,
            audit_summary=audit_summary,
            audit=audit,
            numbering_checks=numbering_checks,
<<<<<<< HEAD
            stage=getattr(result, "stage", None) or "complete",
=======
            diagnostics=tuple(diag_events),
>>>>>>> 769bf0ca3a9a2744c852a1a23a7b3e5f88efb5b3
        )
    except Exception as exc:
        return TargetFormatResult(
            source_path=target,
            success=False,
            output_path=None,
            log=processor_log + (f"FAILED: {exc}",),
            error=str(exc),
            duration_seconds=time.monotonic() - start,
            conversion_report=conversion_report,
            source_sha256=snapshot_sha256,
            audit_summary=audit_summary,
            audit=audit,
            numbering_checks=numbering_checks,
<<<<<<< HEAD
            stage=stage,
=======
            diagnostics=tuple(diag_events),
>>>>>>> 769bf0ca3a9a2744c852a1a23a7b3e5f88efb5b3
        )


def _profile_provenance(profile: TemplateProfile) -> dict[str, Any]:
    manifest = template_analysis.validate_bundle_directory(
        profile.bundle_dir,
        expected_source_sha256=profile.source_sha256,
    )
    producer = getattr(manifest, "producer", {})
    return {
        "bundle_dir": str(profile.bundle_dir),
        "bundle_id": getattr(manifest, "bundle_id", profile.bundle_dir.name),
        "created_utc": getattr(manifest, "created_utc", None),
        "source_sha256": profile.source_sha256,
        "reused": profile.reused,
        "contract_version": _PROFILE_CONTRACT_VERSION,
        "producer": dict(producer) if isinstance(producer, Mapping) else {},
    }


def _redact_json(value: Any, secrets: Sequence[str]) -> Any:
    if isinstance(value, str):
        return _redact(value, secrets)
    if isinstance(value, Mapping):
        # Redact secrets in KEYS as well as values: JSON object keys are a
        # distinct channel that a plain value walk would leave untouched.
        return {
            (_redact(str(key), secrets) or str(key)): _redact_json(item, secrets)
            for key, item in value.items()
        }
    if isinstance(value, (list, tuple)):
        return [_redact_json(item, secrets) for item in value]
    return value


def _write_diagnostics_log(
    run_dir: Path,
    recorder: diag.DiagnosticsRecorder,
    secrets: Sequence[str],
) -> Path:
    """Publish the structured diagnostics stream as one redacted JSONL file.

    Every event's ``fields`` were already reduced to safe scalars/identifiers
    by the recorder; the secret redaction here is defense in depth so an API
    key can never survive even if one slipped into a structural field.
    """

    diagnostics_path = run_dir / "diagnostics.jsonl"
    lines = [
        json.dumps(_redact_json(event, secrets), ensure_ascii=False, sort_keys=True)
        for event in recorder.iter_dicts()
    ]
    payload = ("\n".join(lines) + "\n") if lines else ""
    _atomic_write_bytes(diagnostics_path, payload.encode("utf-8"))
    return diagnostics_path


def _write_run_artifacts(
    *,
    run_id: str,
    conversion_mode: str,
    output_root: Path,
    run_dir: Path,
    architect: Path,
    profile: TemplateProfile,
    template_model: str,
    target_model: str,
    started_utc: datetime,
    finished_utc: datetime,
    targets: Sequence[TargetFormatResult],
    events: Sequence[str],
    secrets: Sequence[str],
    recorder: diag.DiagnosticsRecorder,
) -> tuple[tuple[TargetFormatResult, ...], Path, Path]:
    """Publish per-target audits, diagnostics, run.log, and the run manifest."""

    audited_results: list[TargetFormatResult] = []
    for index, item in enumerate(targets, start=1):
        error_diagnostic = safe_error_diagnostic(item.error, secrets)
        identity = (item.source_sha256 or hashlib.sha256(
            str(item.source_path).encode("utf-8")
        ).hexdigest())[:12]
        audit_path = run_dir / f"target-{index:04d}-{identity}.audit.json"
        conversion = (
            _normalize_audit_details(item.conversion_report.as_dict())
            if item.conversion_report is not None
            else None
        )
        audit_payload = {
            "schema_version": _RUN_AUDIT_VERSION,
            "run_id": run_id,
            "conversion_mode": conversion_mode,
            "source": {
                "path": str(item.source_path),
                "sha256": item.source_sha256,
            },
            "output": (
                {
                    "path": str(item.output_path),
                    "sha256": item.output_sha256,
                }
                if item.output_path is not None
                else None
            ),
            "success": item.success,
            "stage": item.stage,
            "duration_seconds": round(item.duration_seconds, 6),
            "error_code": (
                error_diagnostic.code if error_diagnostic is not None else None
            ),
            "error": (
                error_diagnostic.message if error_diagnostic is not None else None
            ),
            "disposition_counts": dict(item.audit_summary),
            "numbering_checks": _redact_json(item.numbering_checks, secrets),
            "application_audit": _redact_json(item.audit, secrets),
            "conversion_report": _redact_json(conversion, secrets),
            "diagnostics": _redact_json(
                [diag.sanitize_event(event) for event in item.diagnostics],
                secrets,
            ),
        }
        _atomic_write_json(audit_path, audit_payload)
        audited_results.append(replace(item, audit_path=audit_path))

    total_counts = _empty_audit_summary()
    for item in audited_results:
        for key in total_counts:
            total_counts[key] += item.audit_summary.get(key, 0)

    log_lines = [
        _sanitize_run_text(event, secrets, allow_operational=True) or ""
        for event in events
    ]
    for item in audited_results:
        error_diagnostic = safe_error_diagnostic(item.error, secrets)
        log_lines.append(
            f"TARGET {item.source_path.name}: "
            f"{'succeeded' if item.success else 'failed'} "
            f"({item.duration_seconds:.3f}s)"
        )
        if item.stage:
            log_lines.append(f"  STAGE: {item.stage}")
        counts = item.audit_summary
        log_lines.append(
            "  AUDIT COUNTS: "
            f"styled={counts.get('styled', 0)}, "
            f"ignored={counts.get('ignored', 0)}, "
            f"out_of_scope={counts.get('out_of_scope', 0)}, "
            f"unresolved={counts.get('unresolved', 0)}"
        )
        if item.output_path is not None:
            log_lines.append(f"  OUTPUT: {item.output_path}")
        if error_diagnostic is not None:
            log_lines.append(
                f"  ERROR [{error_diagnostic.code}]: {error_diagnostic.message}"
            )
        if item.audit_path is not None:
            log_lines.append(f"  AUDIT: {item.audit_path.name}")
    run_log_path = run_dir / "run.log"
    _atomic_write_bytes(
        run_log_path,
        ("\n".join(log_lines).rstrip() + "\n").encode("utf-8"),
    )

    diagnostics_path = _write_diagnostics_log(run_dir, recorder, secrets)

    succeeded = sum(1 for item in audited_results if item.success)
    failed = len(audited_results) - succeeded
    profile_metadata = _profile_provenance(profile)
    target_records: list[dict[str, Any]] = []
    for item in audited_results:
        error_diagnostic = safe_error_diagnostic(item.error, secrets)
        target_records.append(
            {
                "source_path": str(item.source_path),
                "source_sha256": item.source_sha256,
                "success": item.success,
                "stage": item.stage,
                "output_path": str(item.output_path) if item.output_path else None,
                "output_sha256": item.output_sha256,
                "audit_path": str(item.audit_path) if item.audit_path else None,
                "duration_seconds": round(item.duration_seconds, 6),
                "error_code": (
                    error_diagnostic.code if error_diagnostic is not None else None
                ),
                "error": (
                    error_diagnostic.message if error_diagnostic is not None else None
                ),
                "disposition_counts": dict(item.audit_summary),
                "numbering_checks": _redact_json(item.numbering_checks, secrets),
            }
        )
    manifest_path = run_dir / "run.json"
    manifest = {
        "schema_version": _RUN_MANIFEST_VERSION,
        "run_id": run_id,
        "conversion_mode": conversion_mode,
        "status": (
            "succeeded"
            if failed == 0
            else "failed"
            if succeeded == 0
            else "partial_failure"
        ),
        "started_utc": _iso_utc(started_utc),
        "finished_utc": _iso_utc(finished_utc),
        "duration_seconds": round((finished_utc - started_utc).total_seconds(), 6),
        "application": {
            "name": "spec-template-normalizer",
            "version": APPLICATION_VERSION,
            "template_pipeline_version": template_analysis.PIPELINE_VERSION,
            "application_policy_version": APPLICATION_POLICY_VERSION,
            "profile_contract_version": _PROFILE_CONTRACT_VERSION,
        },
        "diagnostics": _redact_json(recorder.summary(), secrets),
        "paths": {
            "output_root": str(output_root),
            "run_dir": str(run_dir),
            "run_manifest": str(manifest_path),
            "run_log": str(run_log_path),
            "diagnostics_log": str(diagnostics_path),
        },
        "architect_template": {
            "path": str(architect),
            "sha256": profile.source_sha256,
        },
        "template_profile": profile_metadata,
        "models": {
            "template": template_model,
            "target": target_model,
        },
        "prompt_fingerprints": {
            "template": profile_metadata.get("producer", {}).get("prompts", {}),
            "target": _target_prompt_fingerprints(),
        },
        "summary": {
            "targets": len(audited_results),
            "succeeded": succeeded,
            "failed": failed,
            "dispositions": total_counts,
        },
        "targets": target_records,
    }
    _atomic_write_json(manifest_path, manifest)
    return tuple(audited_results), manifest_path, diagnostics_path


def _write_initialization_failure_artifacts(
    *,
    run_id: str,
    conversion_mode: str,
    output_root: Path,
    run_dir: Path,
    architect: Path,
    targets: Sequence[Path],
    template_model: str,
    target_model: str,
    started_utc: datetime,
    events: Sequence[str],
    error: Exception,
    secrets: Sequence[str],
    recorder: diag.DiagnosticsRecorder,
) -> Path:
    """Persist a complete failed-run record when preparation cannot finish."""

    finished_utc = _utc_now()
    error_diagnostic = safe_error_diagnostic(error, secrets)
    if error_diagnostic is None:  # pragma: no cover - ``error`` is concrete
        error_diagnostic = SafeErrorDiagnostic(
            code="untrusted_error",
            message="[untrusted detail omitted]",
        )
    architect_hash: Optional[str]
    try:
        architect_hash = _stable_source_sha256(architect)
    except Exception:
        architect_hash = None

    target_records: list[dict[str, Any]] = []
    for index, target in enumerate(targets, start=1):
        try:
            source_hash = _stable_source_sha256(target)
        except Exception:
            source_hash = None
        identity = (source_hash or hashlib.sha256(str(target).encode("utf-8")).hexdigest())[
            :12
        ]
        audit_path = run_dir / f"target-{index:04d}-{identity}.audit.json"
        audit_payload = {
            "schema_version": _RUN_AUDIT_VERSION,
            "run_id": run_id,
            "conversion_mode": conversion_mode,
            "phase": "not_started",
            "stage": "not_started",
            "source": {"path": str(target), "sha256": source_hash},
            "output": None,
            "success": False,
            "duration_seconds": 0.0,
            "error_type": type(error).__name__,
            "error_code": error_diagnostic.code,
            "error": error_diagnostic.message,
            "disposition_counts": _empty_audit_summary(),
            "numbering_checks": {},
            "application_audit": {},
            "conversion_report": None,
        }
        _atomic_write_json(audit_path, audit_payload)
        target_records.append(
            {
                "source_path": str(target),
                "source_sha256": source_hash,
                "success": False,
                "stage": "not_started",
                "output_path": None,
                "output_sha256": None,
                "audit_path": str(audit_path),
                "duration_seconds": 0.0,
                "error_type": type(error).__name__,
                "error_code": error_diagnostic.code,
                "error": error_diagnostic.message,
                "disposition_counts": _empty_audit_summary(),
                "numbering_checks": {},
            }
        )

    run_log_path = run_dir / "run.log"
    log_lines = [
        _sanitize_run_text(event, secrets, allow_operational=True) or ""
        for event in events
    ]
    log_lines.append(
        "RUN FAILED DURING INITIALIZATION "
        f"[{error_diagnostic.code}]: {error_diagnostic.message}"
    )
    _atomic_write_bytes(
        run_log_path,
        ("\n".join(log_lines).rstrip() + "\n").encode("utf-8"),
    )

    diagnostics_path = _write_diagnostics_log(run_dir, recorder, secrets)

    manifest_path = run_dir / "run.json"
    manifest = {
        "schema_version": _RUN_MANIFEST_VERSION,
        "run_id": run_id,
        "conversion_mode": conversion_mode,
        "status": "failed",
        "failure_phase": "initialization",
        "started_utc": _iso_utc(started_utc),
        "finished_utc": _iso_utc(finished_utc),
        "duration_seconds": round((finished_utc - started_utc).total_seconds(), 6),
        "application": {
            "name": "spec-template-normalizer",
            "version": APPLICATION_VERSION,
            "template_pipeline_version": template_analysis.PIPELINE_VERSION,
            "application_policy_version": APPLICATION_POLICY_VERSION,
            "profile_contract_version": _PROFILE_CONTRACT_VERSION,
        },
        "diagnostics": _redact_json(recorder.summary(), secrets),
        "paths": {
            "output_root": str(output_root),
            "run_dir": str(run_dir),
            "run_manifest": str(manifest_path),
            "run_log": str(run_log_path),
            "diagnostics_log": str(diagnostics_path),
        },
        "architect_template": {"path": str(architect), "sha256": architect_hash},
        "template_profile": None,
        "models": {"template": template_model, "target": target_model},
        "prompt_fingerprints": {"target": _target_prompt_fingerprints()},
        "error_type": type(error).__name__,
        "error_code": error_diagnostic.code,
        "error": error_diagnostic.message,
        "summary": {
            "targets": len(target_records),
            "succeeded": 0,
            "failed": len(target_records),
            "dispositions": _empty_audit_summary(),
        },
        "targets": target_records,
    }
    _atomic_write_json(manifest_path, manifest)
    return manifest_path


def format_specifications(
    architect_template: Path,
    target_specs: Iterable[Path],
    output_dir: Path,
    api_key: str,
    *,
    cache_dir: Optional[Path] = None,
    force_template_analysis: bool = False,
    max_workers: int = 3,
    template_model: str = template_analysis.DEFAULT_MODEL,
    target_model: str = "claude-sonnet-5",
    conversion_mode: str = FORMAT_ONLY,
    diagnostics_level: str = "info",
    template_prompt_dir: Optional[Path] = None,
    template_classifier: Optional[TemplateClassifier] = None,
    progress: Optional[ProgressCallback] = None,
    progress_event: Optional[ProgressEventCallback] = None,
    _template_analyzer: TemplateAnalyzer = template_analysis.run_phase1,
    _config_loader: Callable[[Path], SharedConfig] = load_and_validate_shared_config,
    _target_processor: TargetProcessor = process_single_file,
) -> FormatRunResult:
    """Format one or more target specs using an architect's template.

    This is the canonical public API for the unified application.  Inputs are
    validated before any classifier work begins.  The architect template is
    analyzed once for this run (or reused from a matching validated profile),
    then every target is processed independently so one bad target does not
    discard successful outputs. ``conversion_mode`` selects either formatting
    only or fail-closed CSI-to-Canadian hierarchy conversion in the same run.
<<<<<<< HEAD
    The legacy ``progress`` callback continues to receive plain strings;
    ``progress_event`` additionally receives the UTC occurrence time.
=======

    ``diagnostics_level`` (``debug``/``info``/``warning``/``error``) controls how
    much of the structured diagnostics stream is persisted to
    ``diagnostics.jsonl`` and rolled up into ``run.json``.  The
    ``SPEC_FORMATTER_DIAGNOSTICS_LEVEL`` environment variable overrides it so a
    field build can be asked for verbose diagnostics without a code change.
    Diagnostics never contain secrets or document text.
>>>>>>> 769bf0ca3a9a2744c852a1a23a7b3e5f88efb5b3
    """

    started_utc = _utc_now()
    events: list[str] = []
<<<<<<< HEAD
    pending_events: queue.SimpleQueue[tuple[datetime, str]] = queue.SimpleQueue()
    event_order_lock = threading.Lock()
    event_owner_thread = threading.get_ident()
    last_event_at: Optional[datetime] = None
=======
    resolved_level = diag.level_from_name(
        os.environ.get("SPEC_FORMATTER_DIAGNOSTICS_LEVEL", "") or diagnostics_level,
        default=diag.INFO,
    )
    recorder = diag.DiagnosticsRecorder(min_level=resolved_level)
>>>>>>> 769bf0ca3a9a2744c852a1a23a7b3e5f88efb5b3

    def enqueue_event(
        message: str,
        *,
        occurred_at: Optional[datetime] = None,
    ) -> datetime:
        """Serialize event timestamps across the caller and target workers."""

        nonlocal last_event_at
        with event_order_lock:
            event_time = occurred_at or _utc_now()
            if last_event_at is not None and event_time < last_event_at:
                event_time = last_event_at
            last_event_at = event_time
            pending_events.put((event_time, message))
        return event_time

    def drain_reported_events(*, emit_callbacks: bool = True) -> None:
        """Publish queued events from the calling thread in occurrence order."""

        while True:
            try:
                event_time, message = pending_events.get_nowait()
            except queue.Empty:
                return
            events.append(f"{_iso_utc(event_time)} {message}")
            if emit_callbacks:
                _emit(progress, message)
                _emit_progress_event(progress_event, message, event_time)

    def report(
        message: str,
        *,
        occurred_at: Optional[datetime] = None,
    ) -> None:
        enqueue_event(message, occurred_at=occurred_at)
        # Worker events are drained by the calling thread's submit/wait loop,
        # preserving the historical callback thread affinity.
        if threading.get_ident() == event_owner_thread:
            drain_reported_events()

    report("Checking input files...")
    try:
        conversion_mode = validate_conversion_mode(conversion_mode)
    except ValueError as exc:
        _attach_safe_error_diagnostic(
            exc,
            code="invalid_conversion_mode",
            message="conversion_mode must be one of: csi_to_canadian, format_only",
        )
        raise
    if not isinstance(api_key, str):
        message = "Anthropic API key must be text."
        raise _attach_safe_error_diagnostic(
            ValueError(message),
            code="api_key_invalid_type",
            message=message,
        )
    if not isinstance(max_workers, int) or isinstance(max_workers, bool):
        message = "max_workers must be an integer."
        raise _attach_safe_error_diagnostic(
            ValueError(message),
            code="invalid_max_workers",
            message=message,
        )
    normalized_api_key = api_key.strip()
    architect, targets, destination = _validate_inputs(
        architect_template,
        tuple(target_specs),
        output_dir,
    )
    workers = max(1, min(max_workers, _MAX_WORKERS, len(targets)))
    profile_cache = (
        Path(cache_dir).expanduser().resolve()
        if cache_dir is not None
        else default_template_cache_dir().expanduser().resolve()
    )
    run_id, run_dir = _create_run_directory(destination, conversion_mode)
    recorder.info(
        "pipeline",
        "run_start",
        targets=len(targets),
        workers=workers,
        mode=conversion_mode,
    )
    try:
        with recorder.timer("pipeline", "template_analysis") as phase:
            profile = prepare_template_profile(
                architect,
                profile_cache,
                normalized_api_key,
                force_analysis=force_template_analysis,
                model=template_model,
                prompt_dir=template_prompt_dir,
                progress=report,
                classifier=template_classifier,
                analyzer=_template_analyzer,
            )
            phase.set(reused=profile.reused)
        if _stable_source_sha256(architect) != profile.source_sha256:
            raise RuntimeError(
                "The architect template changed during this run. Finish saving it and run again."
            )

        report("Validating the template profile...")
        with recorder.timer("pipeline", "config_load"):
            shared = _config_loader(profile.bundle_dir)
        planned_outputs = _plan_output_paths(targets, run_dir, conversion_mode)
        _validate_output_plan(architect, targets, planned_outputs)
    except Exception as exc:
<<<<<<< HEAD
        drain_reported_events()
=======
        recorder.error("pipeline", "init_failed", error_type=type(exc).__name__.lower())
>>>>>>> 769bf0ca3a9a2744c852a1a23a7b3e5f88efb5b3
        manifest_path = _write_initialization_failure_artifacts(
            run_id=run_id,
            conversion_mode=conversion_mode,
            output_root=destination,
            run_dir=run_dir,
            architect=architect,
            targets=targets,
            template_model=template_model,
            target_model=target_model,
            started_utc=started_utc,
            events=events,
            error=exc,
            secrets=(normalized_api_key,),
            recorder=recorder,
        )
        try:
            setattr(exc, "run_dir", run_dir)
            setattr(exc, "manifest_path", manifest_path)
        except Exception:  # pragma: no cover - unusual immutable exception type
            pass
        raise
    results_by_target: dict[Path, TargetFormatResult] = {}

    with tempfile.TemporaryDirectory(
        prefix="sf-",
    ) as job_temp:
        job_root = Path(job_temp)
        with ThreadPoolExecutor(max_workers=workers) as executor:
<<<<<<< HEAD
            futures: dict[Future[TargetFormatResult], Path] = {}

=======
            futures: dict[Future[TargetFormatResult], tuple[int, Path]] = {}
>>>>>>> 769bf0ca3a9a2744c852a1a23a7b3e5f88efb5b3
            for index, target in enumerate(targets):
                staging_dir = job_root / f"t{index:04d}"
                report(f"Queued {index + 1} of {len(targets)}: {target.name}")
                future = executor.submit(
                    _format_one_target,
                    target,
                    planned_outputs[target],
                    staging_dir,
                    shared,
                    normalized_api_key,
                    target_model,
                    _target_processor,
                    conversion_mode,
                    lambda index=index, target=target: report(
                        f"Processing target {index + 1} of {len(targets)}: "
                        f"{target.name}"
                    ),
                )
<<<<<<< HEAD
                futures[future] = target
                # A very fast worker may have started before ``submit``
                # returns. Surface that event before queuing the next target.
                drain_reported_events()

            completed = 0
            pending = set(futures)
            while pending:
                drain_reported_events()
                done, pending = wait(
                    pending,
                    timeout=0.05,
                    return_when=FIRST_COMPLETED,
=======
                futures[future] = (index + 1, target)

            completed = 0
            for future in as_completed(futures):
                target_number, target = futures[future]
                try:
                    result = future.result()
                except Exception as exc:  # pragma: no cover - defensive boundary
                    result = TargetFormatResult(
                        source_path=target,
                        success=False,
                        output_path=None,
                        log=(f"FAILED: {exc}",),
                        error=str(exc),
                        duration_seconds=0.0,
                    )
                results_by_target[target] = result
                completed += 1
                recorder.ingest(result.diagnostics, target=target_number)
                counts = result.audit_summary
                recorder.record(
                    diag.INFO if result.success else diag.WARNING,
                    "pipeline",
                    "target_done",
                    target=target_number,
                    success=result.success,
                    duration_ms=round(result.duration_seconds * 1000.0, 3),
                    styled=counts.get("styled", 0),
                    ignored=counts.get("ignored", 0),
                    out_of_scope=counts.get("out_of_scope", 0),
                    unresolved=counts.get("unresolved", 0),
                )
                for line in result.log:
                    report(f"Target {target.name}: {line}")
                report(
                    f"Target {target.name}: audit styled={counts.get('styled', 0)}, "
                    f"ignored={counts.get('ignored', 0)}, "
                    f"out_of_scope={counts.get('out_of_scope', 0)}, "
                    f"unresolved={counts.get('unresolved', 0)}"
>>>>>>> 769bf0ca3a9a2744c852a1a23a7b3e5f88efb5b3
                )
                drain_reported_events()
                for future in done:
                    target = futures[future]
                    try:
                        result = future.result()
                    except Exception as exc:  # pragma: no cover - defensive boundary
                        result = TargetFormatResult(
                            source_path=target,
                            success=False,
                            output_path=None,
                            log=(f"FAILED: {exc}",),
                            error=str(exc),
                            duration_seconds=0.0,
                            stage="processing",
                        )
                    results_by_target[target] = result
                    completed += 1
                    for line in result.log:
                        if str(line).lstrip().startswith("FAILED:"):
                            continue
                        safe_line = _sanitize_run_text(
                            str(line),
                            (normalized_api_key,),
                            allow_operational=True,
                        )
                        for safe_part in (safe_line or "").splitlines():
                            if not safe_part or _SAFE_REDACTION_RX.fullmatch(safe_part):
                                continue
                            report(f"Target {target.name}: {safe_part}")
                    status = "Formatted" if result.success else "Failed"
                    report(
                        f"{status} {completed} of {len(targets)}: {target.name}"
                    )
            drain_reported_events()

    ordered_results = tuple(results_by_target[target] for target in targets)
    succeeded = sum(1 for item in ordered_results if item.success)
    failed = len(ordered_results) - succeeded
    recorder.info(
        "pipeline",
        "run_complete",
        targets=len(ordered_results),
        succeeded=succeeded,
        failed=failed,
    )
    complete_message = f"Complete: {succeeded} succeeded, {failed} failed."
    complete_occurred_at = enqueue_event(complete_message)
    # Keep the durable log's terminal event while preserving the historical
    # guarantee that UI completion is emitted only after artifacts publish.
    drain_reported_events(emit_callbacks=False)
    finished_utc = _utc_now()
    audited_results, manifest_path, diagnostics_path = _write_run_artifacts(
        run_id=run_id,
        conversion_mode=conversion_mode,
        output_root=destination,
        run_dir=run_dir,
        architect=architect,
        profile=profile,
        template_model=template_model,
        target_model=target_model,
        started_utc=started_utc,
        finished_utc=finished_utc,
        targets=ordered_results,
        events=events,
        secrets=(normalized_api_key,),
        recorder=recorder,
    )
    run_result = FormatRunResult(
        template_profile=profile,
        output_dir=run_dir,
        targets=audited_results,
        run_id=run_id,
        conversion_mode=conversion_mode,
        output_root=destination,
        run_dir=run_dir,
        manifest_path=manifest_path,
        diagnostics_path=diagnostics_path,
    )
    _emit(progress, complete_message)
    _emit_progress_event(progress_event, complete_message, complete_occurred_at)
    return run_result


__all__ = [
    "FormatRunResult",
    "CSI_TO_CANADIAN",
    "FORMAT_ONLY",
    "ProgressEventCallback",
    "SafeErrorDiagnostic",
    "TargetFormatResult",
    "TemplateProfile",
    "collect_target_specs",
    "default_template_cache_dir",
    "format_specifications",
    "prepare_template_profile",
    "safe_error_diagnostic",
]
