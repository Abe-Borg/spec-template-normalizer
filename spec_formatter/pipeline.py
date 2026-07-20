"""One-call architect-template analysis and target-spec formatting.

The validated template bundle remains an internal integrity boundary.  Callers
provide the architect DOCX and target DOCX files; they never need to create,
locate, or transfer a bundle themselves.
"""

from __future__ import annotations

import hashlib
import os
import re
import tempfile
import time
import warnings
from concurrent.futures import Future, ThreadPoolExecutor, as_completed
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Callable, Iterable, Optional, Sequence

from . import template_analysis
from .style_application.batch_runner import (
    BatchResult,
    SharedConfig,
    load_and_validate_shared_config,
    process_single_file,
)
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


@dataclass(frozen=True)
class FormatRunResult:
    """Consolidated result returned by :func:`format_specifications`."""

    template_profile: TemplateProfile
    output_dir: Path
    targets: tuple[TargetFormatResult, ...]

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


def collect_target_specs(inputs: Iterable[Path]) -> tuple[Path, ...]:
    """Expand DOCX files and folders into a stable, deduplicated target list.

    Folder discovery is intentionally non-recursive and excludes Word lock
    files plus outputs from current and legacy versions of the formatter.
    """

    discovered: list[Path] = []
    for raw_path in inputs:
        path = Path(raw_path).expanduser()
        if path.is_dir():
            discovered.extend(
                candidate
                for candidate in path.glob("*.docx")
                if not candidate.name.startswith("~$")
                and not _is_formatted_output(candidate)
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

    targets = collect_target_specs(target_specs)
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
    cache_root = Path(cache_dir).resolve()
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
            filename = f"{stem}{suffix}"
        else:
            parent = _safe_filename_fragment(target.parent.name)
            digest = hashlib.sha1(str(target).encode("utf-8")).hexdigest()[:8]
            filename = f"{stem}__{parent}-{digest}{suffix}"
        folded = filename.casefold()
        if folded in used_names:
            raise ValueError(f"Could not create unique output name for {target}")
        used_names.add(folded)
        planned[target] = output_dir / filename
    return planned


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
    try:
        snapshot = staging_dir / "source" / target.name
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
        if not result.success:
            return TargetFormatResult(
                source_path=target,
                success=False,
                output_path=None,
                log=processor_log,
                error=result.error or "Target formatting failed.",
                duration_seconds=result.duration_seconds,
                conversion_report=conversion_report,
            )
        if result.output_path is None or not result.output_path.is_file():
            raise RuntimeError("Style application reported success without an output DOCX.")
        if _stable_source_sha256(target) != snapshot_sha256:
            raise RuntimeError(
                f"{target.name} changed during formatting. Finish saving it and run again."
            )
        final_output.parent.mkdir(parents=True, exist_ok=True)
        os.replace(result.output_path, final_output)
        return TargetFormatResult(
            source_path=target,
            success=True,
            output_path=final_output,
            log=processor_log,
            error=None,
            duration_seconds=result.duration_seconds,
            conversion_report=conversion_report,
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
        )


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

    _emit(progress, "Checking input files...")
    conversion_mode = validate_conversion_mode(conversion_mode)
    if not isinstance(api_key, str):
        raise ValueError("Anthropic API key must be text.")
    normalized_api_key = api_key.strip()
    architect, targets, destination = _validate_inputs(
        architect_template,
        tuple(target_specs),
        output_dir,
    )
    planned_outputs = _plan_output_paths(targets, destination, conversion_mode)
    _validate_output_plan(architect, targets, planned_outputs)
    workers = max(1, min(int(max_workers), _MAX_WORKERS, len(targets)))
    profile_cache = (
        Path(cache_dir).expanduser().resolve()
        if cache_dir is not None
        else default_template_cache_dir().expanduser().resolve()
    )
    profile = prepare_template_profile(
        architect,
        profile_cache,
        normalized_api_key,
        force_analysis=force_template_analysis,
        model=template_model,
        prompt_dir=template_prompt_dir,
        progress=progress,
        classifier=template_classifier,
        analyzer=_template_analyzer,
    )
    if _stable_source_sha256(architect) != profile.source_sha256:
        raise RuntimeError(
            "The architect template changed during this run. Finish saving it and run again."
        )

    _emit(progress, "Validating the template profile...")
    shared = _config_loader(profile.bundle_dir)
    results_by_target: dict[Path, TargetFormatResult] = {}

    with tempfile.TemporaryDirectory(
        prefix=".spec-formatter-job-",
        dir=destination,
    ) as job_temp:
        job_root = Path(job_temp)
        with ThreadPoolExecutor(max_workers=workers) as executor:
            futures: dict[Future[TargetFormatResult], Path] = {}
            for index, target in enumerate(targets):
                staging_dir = job_root / f"target-{index:04d}"
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
                status = "Formatted" if result.success else "Failed"
                _emit(progress, f"{status} {completed} of {len(targets)}: {target.name}")

    ordered_results = tuple(results_by_target[target] for target in targets)
    run_result = FormatRunResult(profile, destination, ordered_results)
    _emit(
        progress,
        f"Complete: {run_result.succeeded} succeeded, {run_result.failed} failed.",
    )
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
