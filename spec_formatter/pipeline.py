"""One-call architect-template analysis and target-spec formatting.

The validated template bundle remains an internal integrity boundary.  Callers
provide the architect DOCX and target DOCX files; they never need to create,
locate, or transfer a bundle themselves.
"""

from __future__ import annotations

import hashlib
import json
import os
import re
import tempfile
import time
import uuid
import warnings
from concurrent.futures import Future, ThreadPoolExecutor, as_completed
from dataclasses import dataclass, field, replace
from datetime import datetime, timezone
from pathlib import Path
from typing import Any, Callable, Iterable, Mapping, Optional, Sequence

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
TemplateClassifier = Callable[..., dict[str, Any]]
TemplateAnalyzer = Callable[..., template_analysis.Phase1Result]
TargetProcessor = Callable[..., BatchResult]

_FORMATTED_SUFFIXES = (
    "_FORMATTED.DOCX",
    "_CANADIAN_FORMATTED.DOCX",
    "_PHASE2_FORMATTED.DOCX",
)
_MAX_WORKERS = 6
_RUN_MANIFEST_VERSION = 1
_RUN_AUDIT_VERSION = 1
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
        raise FileNotFoundError(f"Architect template does not exist: {architect}")
    if architect.suffix.lower() != ".docx":
        raise ValueError(f"Architect template must be a .docx file: {architect}")
    if architect.name.startswith("~$"):
        raise ValueError("Select the saved architect DOCX, not Word's temporary lock file.")

    targets = collect_target_specs(target_specs, exclude_discovered=architect)
    if not targets:
        raise ValueError("Select at least one target specification DOCX file.")

    architect_key = os.path.normcase(str(architect))
    for target in targets:
        if not target.is_file():
            raise FileNotFoundError(f"Target specification does not exist: {target}")
        if target.suffix.lower() != ".docx":
            raise ValueError(f"Target specification must be a .docx file: {target}")
        if target.name.startswith("~$"):
            raise ValueError(f"Target is a Word temporary lock file: {target}")
        if _is_formatted_output(target):
            raise ValueError(f"Target is already a formatted output: {target}")
        if os.path.normcase(str(target)) == architect_key:
            raise ValueError("The architect template cannot also be a target specification.")

    destination = Path(output_dir).expanduser().resolve()
    destination.mkdir(parents=True, exist_ok=True)
    if not destination.is_dir():
        raise NotADirectoryError(f"Output location is not a directory: {destination}")
    try:
        with tempfile.NamedTemporaryFile(
            prefix=".spec-formatter-write-check-",
            dir=destination,
        ):
            pass
    except OSError as exc:
        raise PermissionError(f"Output directory is not writable: {destination}") from exc
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
    "Classifications saved:",
    "Classifying ",
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
        if _OOXML_FRAGMENT_RX.search(raw_line) or _DOCUMENT_DATA_FIELD_RX.search(
            raw_line
        ):
            safe_lines.append("[document content omitted]")
            continue
        bounded = raw_line[:4096]
        if allow_operational and _is_safe_operational_line(bounded):
            safe_lines.append(bounded)
            continue
        fingerprint = hashlib.sha256(bounded.encode("utf-8")).hexdigest()[:12]
        safe_lines.append(f"[untrusted detail omitted; sha256={fingerprint}]")
    return "\n".join(safe_lines)


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
) -> TargetFormatResult:
    start = time.monotonic()
    processor_log: tuple[str, ...] = ()
    conversion_report: Optional[CanadianConversionReport] = None
    snapshot_sha256: Optional[str] = None
    audit_summary = _empty_audit_summary()
    audit: dict[str, Any] = {}
    numbering_checks: dict[str, Any] = {}
    try:
        # Deliberately avoid carrying user-controlled filenames into the work
        # tree.  This keeps Windows paths short even for deeply nested inputs.
        snapshot = staging_dir / "source.docx"
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
        processor_log = tuple(result.log)
        conversion_report = result.conversion_report
        audit_summary = _normalize_audit_summary(
            getattr(result, "audit_summary", None)
        )
        audit = _normalize_audit_details(getattr(result, "audit", None))
        numbering_checks = _normalize_numbering_checks(
            getattr(result, "numbering_checks", None)
        )
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
            )
        if result.output_path is None or not result.output_path.is_file():
            raise RuntimeError("Style application reported success without an output DOCX.")
        if _stable_source_sha256(target) != snapshot_sha256:
            raise RuntimeError(
                f"{target.name} changed during formatting. Finish saving it and run again."
            )
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
        return {str(key): _redact_json(item, secrets) for key, item in value.items()}
    if isinstance(value, (list, tuple)):
        return [_redact_json(item, secrets) for item in value]
    return value


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
) -> tuple[tuple[TargetFormatResult, ...], Path]:
    """Publish per-target audits, run.log, and finally the run manifest."""

    audited_results: list[TargetFormatResult] = []
    for index, item in enumerate(targets, start=1):
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
            "duration_seconds": round(item.duration_seconds, 6),
            "error": _sanitize_run_text(item.error, secrets),
            "disposition_counts": dict(item.audit_summary),
            "numbering_checks": _redact_json(item.numbering_checks, secrets),
            "application_audit": _redact_json(item.audit, secrets),
            "conversion_report": _redact_json(conversion, secrets),
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
        log_lines.append(
            f"TARGET {item.source_path.name}: "
            f"{'succeeded' if item.success else 'failed'} "
            f"({item.duration_seconds:.3f}s)"
        )
        for line in item.log:
            log_lines.append(
                f"  {_sanitize_run_text(str(line), secrets, allow_operational=True)}"
            )
        if item.error:
            log_lines.append(f"  ERROR: {_sanitize_run_text(item.error, secrets)}")
        if item.audit_path is not None:
            log_lines.append(f"  AUDIT: {item.audit_path.name}")
    run_log_path = run_dir / "run.log"
    _atomic_write_bytes(
        run_log_path,
        ("\n".join(log_lines).rstrip() + "\n").encode("utf-8"),
    )

    succeeded = sum(1 for item in audited_results if item.success)
    failed = len(audited_results) - succeeded
    profile_metadata = _profile_provenance(profile)
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
            "version": template_analysis.PIPELINE_VERSION,
            "application_policy_version": APPLICATION_POLICY_VERSION,
            "profile_contract_version": _PROFILE_CONTRACT_VERSION,
        },
        "paths": {
            "output_root": str(output_root),
            "run_dir": str(run_dir),
            "run_manifest": str(manifest_path),
            "run_log": str(run_log_path),
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
        "targets": [
            {
                "source_path": str(item.source_path),
                "source_sha256": item.source_sha256,
                "success": item.success,
                "output_path": str(item.output_path) if item.output_path else None,
                "output_sha256": item.output_sha256,
                "audit_path": str(item.audit_path) if item.audit_path else None,
                "duration_seconds": round(item.duration_seconds, 6),
                "error": _sanitize_run_text(item.error, secrets),
                "disposition_counts": dict(item.audit_summary),
                "numbering_checks": _redact_json(item.numbering_checks, secrets),
            }
            for item in audited_results
        ],
    }
    _atomic_write_json(manifest_path, manifest)
    return tuple(audited_results), manifest_path


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
) -> Path:
    """Persist a complete failed-run record when preparation cannot finish."""

    finished_utc = _utc_now()
    safe_error = _sanitize_run_text(str(error), secrets)
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
            "source": {"path": str(target), "sha256": source_hash},
            "output": None,
            "success": False,
            "duration_seconds": 0.0,
            "error_type": type(error).__name__,
            "error": safe_error,
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
                "output_path": None,
                "output_sha256": None,
                "audit_path": str(audit_path),
                "duration_seconds": 0.0,
                "error_type": type(error).__name__,
                "error": safe_error,
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
        f"RUN FAILED DURING INITIALIZATION: "
        f"{_sanitize_run_text(str(error), secrets)}"
    )
    _atomic_write_bytes(
        run_log_path,
        ("\n".join(log_lines).rstrip() + "\n").encode("utf-8"),
    )

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
            "version": template_analysis.PIPELINE_VERSION,
            "application_policy_version": APPLICATION_POLICY_VERSION,
            "profile_contract_version": _PROFILE_CONTRACT_VERSION,
        },
        "paths": {
            "output_root": str(output_root),
            "run_dir": str(run_dir),
            "run_manifest": str(manifest_path),
            "run_log": str(run_log_path),
        },
        "architect_template": {"path": str(architect), "sha256": architect_hash},
        "template_profile": None,
        "models": {"template": template_model, "target": target_model},
        "prompt_fingerprints": {"target": _target_prompt_fingerprints()},
        "error_type": type(error).__name__,
        "error": safe_error,
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
    template_prompt_dir: Optional[Path] = None,
    template_classifier: Optional[TemplateClassifier] = None,
    progress: Optional[ProgressCallback] = None,
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
    """

    started_utc = _utc_now()
    events: list[str] = []

    def report(message: str) -> None:
        events.append(f"{_iso_utc(_utc_now())} {message}")
        _emit(progress, message)

    report("Checking input files...")
    conversion_mode = validate_conversion_mode(conversion_mode)
    if not isinstance(api_key, str):
        raise ValueError("Anthropic API key must be text.")
    normalized_api_key = api_key.strip()
    architect, targets, destination = _validate_inputs(
        architect_template,
        tuple(target_specs),
        output_dir,
    )
    workers = max(1, min(int(max_workers), _MAX_WORKERS, len(targets)))
    profile_cache = (
        Path(cache_dir).expanduser().resolve()
        if cache_dir is not None
        else default_template_cache_dir().expanduser().resolve()
    )
    run_id, run_dir = _create_run_directory(destination, conversion_mode)
    try:
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
        if _stable_source_sha256(architect) != profile.source_sha256:
            raise RuntimeError(
                "The architect template changed during this run. Finish saving it and run again."
            )

        report("Validating the template profile...")
        shared = _config_loader(profile.bundle_dir)
        planned_outputs = _plan_output_paths(targets, run_dir, conversion_mode)
        _validate_output_plan(architect, targets, planned_outputs)
    except Exception as exc:
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
            futures: dict[Future[TargetFormatResult], Path] = {}
            for index, target in enumerate(targets):
                staging_dir = job_root / f"t{index:04d}"
                report(f"Started {index + 1} of {len(targets)}: {target.name}")
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
                )
                futures[future] = target

            completed = 0
            for future in as_completed(futures):
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
                    )
                results_by_target[target] = result
                completed += 1
                for line in result.log:
                    report(f"Target {target.name}: {line}")
                counts = result.audit_summary
                report(
                    f"Target {target.name}: audit styled={counts.get('styled', 0)}, "
                    f"ignored={counts.get('ignored', 0)}, "
                    f"out_of_scope={counts.get('out_of_scope', 0)}, "
                    f"unresolved={counts.get('unresolved', 0)}"
                )
                status = "Formatted" if result.success else "Failed"
                report(f"{status} {completed} of {len(targets)}: {target.name}")

    ordered_results = tuple(results_by_target[target] for target in targets)
    succeeded = sum(1 for item in ordered_results if item.success)
    failed = len(ordered_results) - succeeded
    complete_message = f"Complete: {succeeded} succeeded, {failed} failed."
    events.append(f"{_iso_utc(_utc_now())} {complete_message}")
    finished_utc = _utc_now()
    audited_results, manifest_path = _write_run_artifacts(
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
    )
    _emit(progress, complete_message)
    return run_result


__all__ = [
    "FormatRunResult",
    "CSI_TO_CANADIAN",
    "FORMAT_ONLY",
    "TargetFormatResult",
    "TemplateProfile",
    "collect_target_specs",
    "default_template_cache_dir",
    "format_specifications",
    "prepare_template_profile",
]
