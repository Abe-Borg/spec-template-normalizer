from __future__ import annotations

import hashlib
import json
import os
import tempfile
import threading
from datetime import datetime
from pathlib import Path
from types import SimpleNamespace

import pytest

import spec_formatter
from spec_formatter import pipeline
from spec_formatter.style_application.batch_runner import BatchResult, SharedConfig
from spec_formatter.style_application.core.csi_to_canadian import (
    CanadianConversionReport,
    ConversionIssue,
    MarkerEdit,
)


def _write_input(path: Path, contents: bytes) -> Path:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_bytes(contents)
    return path


def _install_fake_bundle_validator(monkeypatch: pytest.MonkeyPatch) -> None:
    """Validate the tiny profile directories published by the fake analyzer."""

    def validate_bundle_directory(
        bundle_dir: Path,
        *,
        expected_source_sha256: str,
    ) -> SimpleNamespace:
        recorded_hash = (Path(bundle_dir) / "source.sha256").read_text(encoding="ascii")
        if recorded_hash != expected_source_sha256:
            raise ValueError("profile source hash does not match")
        prompt_dir = Path(pipeline.__file__).resolve().parents[1]
        return SimpleNamespace(
            producer={
                "name": "spec-template-normalizer",
                "version": pipeline.template_analysis.PIPELINE_VERSION,
                "engine_fingerprint": pipeline.template_analysis.ENGINE_SOURCE_DIGEST,
                "classifier": {
                    "provider": "anthropic",
                    "model": pipeline.template_analysis.DEFAULT_MODEL,
                },
                "prompts": {
                    "master_prompt_sha256": hashlib.sha256(
                        (prompt_dir / "master_prompt.txt")
                        .read_text(encoding="utf-8")
                        .encode("utf-8")
                    ).hexdigest(),
                    "run_instruction_sha256": hashlib.sha256(
                        (prompt_dir / "run_instruction_prompt.txt")
                        .read_text(encoding="utf-8")
                        .encode("utf-8")
                    ).hexdigest(),
                },
            }
        )

    monkeypatch.setattr(
        pipeline.template_analysis,
        "validate_bundle_directory",
        validate_bundle_directory,
    )


def _fake_dependencies(
    monkeypatch: pytest.MonkeyPatch,
    *,
    failing_target: str | None = None,
) -> tuple[dict[str, list], object, object, object]:
    _install_fake_bundle_validator(monkeypatch)
    calls: dict[str, list] = {
        "analyzer": [],
        "config_loader": [],
        "processor": [],
    }

    def analyzer(**kwargs):
        calls["analyzer"].append(kwargs)
        source_hash = pipeline.template_analysis.sha256_file(kwargs["source_docx"])
        bundle_dir = (
            Path(kwargs["output_root"])
            / (
                f"architect--{source_hash[:12]}--"
                f"offline-test-{len(calls['analyzer'])}.phase1"
            )
        )
        bundle_dir.mkdir(parents=True)
        (bundle_dir / "source.sha256").write_text(source_hash, encoding="ascii")
        return SimpleNamespace(bundle_dir=bundle_dir)

    def config_loader(bundle_dir: Path) -> SharedConfig:
        calls["config_loader"].append(Path(bundle_dir))
        return SharedConfig(
            arch_registry={"BODY": "ArchitectBody"},
            env_registry={"test": True},
            arch_styles_xml="<w:styles/>",
            available_roles=["BODY"],
            source_tokens={},
            arch_root=Path(bundle_dir),
            role_specs={"BODY": {"style_id": "ArchitectBody"}},
        )

    def processor(**kwargs) -> BatchResult:
        source = Path(kwargs["docx_path"])
        calls["processor"].append(source)
        source_bytes = source.read_bytes()
        if failing_target and source_bytes.startswith(Path(failing_target).stem.encode()):
            return BatchResult(
                filename=source.name,
                success=False,
                output_path=None,
                log=["simulated target failure"],
                error="simulated target failure",
                duration_seconds=0.01,
                audit_summary={
                    "styled": 0,
                    "ignored": 0,
                    "out_of_scope": 0,
                    "unresolved": 1,
                },
                numbering_checks={"preserved": False},
            )

        staging_dir = Path(kwargs["output_dir"])
        staging_dir.mkdir(parents=True, exist_ok=True)
        staged_output = staging_dir / f"{source.stem}_PHASE2_FORMATTED.docx"
        staged_output.write_bytes(b"formatted:" + source_bytes)
        return BatchResult(
            filename=source.name,
            success=True,
            output_path=staged_output,
            log=[
                "Applied classifications, stability verified",
                f"Output: {staged_output}",
            ],
            error=None,
            duration_seconds=0.02,
            audit_summary={
                "styled": 1,
                "ignored": 2,
                "out_of_scope": 3,
                "unresolved": 0,
            },
            audit={
                "schema_version": 1,
                "paragraph_indices": [0],
                "out_of_scope": [
                    {
                        "paragraph_index": 2,
                        "reason": "table",
                        "original_text_preview": "never persist paragraph text",
                    }
                ],
            },
            numbering_checks={"preserved": True, "checked": 1},
        )

    return calls, analyzer, config_loader, processor


def _run_with_fakes(
    architect: Path,
    targets: list[Path],
    output_dir: Path,
    *,
    analyzer,
    config_loader,
    processor,
    cache_dir: Path | None = None,
    progress=None,
    progress_event=None,
):
    if cache_dir is None:
        cache_dir = output_dir.parent / "test-profile-cache"
    return pipeline.format_specifications(
        architect,
        targets,
        output_dir,
        api_key="offline-test-key",
        cache_dir=cache_dir,
        max_workers=3,
        progress=progress,
        progress_event=progress_event,
        _template_analyzer=analyzer,
        _config_loader=config_loader,
        _target_processor=processor,
    )


def test_one_call_analyzes_template_once_for_multiple_targets_and_preserves_inputs(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    architect = _write_input(tmp_path / "architect.docx", b"architect-original")
    targets = [
        _write_input(tmp_path / "mechanical.docx", b"mechanical-original"),
        _write_input(tmp_path / "electrical.docx", b"electrical-original"),
    ]
    originals = {path: path.read_bytes() for path in [architect, *targets]}
    calls, analyzer, config_loader, processor = _fake_dependencies(monkeypatch)

    result = _run_with_fakes(
        architect,
        targets,
        tmp_path / "formatted",
        analyzer=analyzer,
        config_loader=config_loader,
        processor=processor,
    )

    assert result.success is True
    assert result.succeeded == 2
    assert result.failed == 0
    assert result.template_profile.reused is False
    assert len(calls["analyzer"]) == 1
    assert len(calls["config_loader"]) == 1
    assert {path.name for path in calls["processor"]} == {"source.docx"}
    assert len({path.parent for path in calls["processor"]}) == 2
    assert all(
        os.path.commonpath([path, Path(tempfile.gettempdir())])
        == str(Path(tempfile.gettempdir()))
        for path in calls["processor"]
    )
    assert {path.name for path in result.output_paths} == {
        "mechanical_FORMATTED.docx",
        "electrical_FORMATTED.docx",
    }
    assert all(path.is_file() for path in result.output_paths)
    assert result.output_root == (tmp_path / "formatted").resolve()
    assert result.output_dir == result.run_dir
    assert result.run_dir is not None and result.run_dir.parent == result.output_root
    assert result.manifest_path == result.run_dir / "run.json"
    assert result.manifest_path.is_file()
    assert (result.run_dir / "run.log").is_file()
    assert all(item.audit_path and item.audit_path.is_file() for item in result.targets)
    assert {path: path.read_bytes() for path in originals} == originals


def test_validated_template_profile_is_reused_on_a_later_run(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    architect = _write_input(tmp_path / "architect.docx", b"stable-architect")
    target = _write_input(tmp_path / "target.docx", b"stable-target")
    cache_dir = tmp_path / "profile-cache"
    calls, analyzer, config_loader, processor = _fake_dependencies(monkeypatch)
    progress_messages: list[str] = []

    first = _run_with_fakes(
        architect,
        [target],
        tmp_path / "formatted",
        cache_dir=cache_dir,
        analyzer=analyzer,
        config_loader=config_loader,
        processor=processor,
        progress=progress_messages.append,
    )
    second = pipeline.format_specifications(
        architect,
        [target],
        tmp_path / "formatted",
        api_key="",
        cache_dir=cache_dir,
        progress=progress_messages.append,
        _template_analyzer=analyzer,
        _config_loader=config_loader,
        _target_processor=processor,
    )

    assert first.template_profile.reused is False
    assert second.template_profile.reused is True
    assert first.template_profile.bundle_dir == second.template_profile.bundle_dir
    assert first.template_profile.bundle_dir.parent == (
        cache_dir.resolve() / pipeline._PROFILE_CACHE_NAMESPACE
    )
    assert first.run_dir != second.run_dir
    assert first.run_dir is not None and second.run_dir is not None
    assert first.run_dir.parent == second.run_dir.parent == (tmp_path / "formatted").resolve()
    assert len(calls["analyzer"]) == 1
    assert len(calls["config_loader"]) == 2
    assert len(calls["processor"]) == 2
    assert "Reusing the validated architect template analysis." in progress_messages
    assert architect.read_bytes() == b"stable-architect"
    assert target.read_bytes() == b"stable-target"


def test_run_manifest_log_and_audits_capture_provenance_without_api_key(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    architect = _write_input(tmp_path / "architect.docx", b"architect-original")
    target = _write_input(tmp_path / "target.docx", b"target-original")
    calls, analyzer, config_loader, processor = _fake_dependencies(monkeypatch)
    api_key = "never-persist-this-secret"

    result = pipeline.format_specifications(
        architect,
        [target],
        tmp_path / "formatted",
        api_key=api_key,
        cache_dir=tmp_path / "profile-cache",
        _template_analyzer=analyzer,
        _config_loader=config_loader,
        _target_processor=processor,
    )

    assert result.manifest_path is not None
    manifest_text = result.manifest_path.read_text(encoding="utf-8")
    run_log_text = (result.run_dir / "run.log").read_text(encoding="utf-8")
    assert api_key not in manifest_text
    assert api_key not in run_log_text
    manifest = json.loads(manifest_text)
    assert manifest["schema_version"] == 2
    assert manifest["run_id"] == result.run_id
    assert manifest["conversion_mode"] == pipeline.FORMAT_ONLY
    assert manifest["application"]["version"] == spec_formatter.__version__
    assert (
        manifest["application"]["template_pipeline_version"]
        == pipeline.template_analysis.PIPELINE_VERSION
    )
    assert (
        manifest["template_profile"]["producer"]["version"]
        == pipeline.template_analysis.PIPELINE_VERSION
    )
    assert manifest["paths"]["output_root"] == str(result.output_root)
    assert manifest["paths"]["run_dir"] == str(result.run_dir)
    assert manifest["architect_template"]["sha256"] == hashlib.sha256(
        b"architect-original"
    ).hexdigest()
    assert manifest["summary"] == {
        "targets": 1,
        "succeeded": 1,
        "failed": 0,
        "dispositions": {
            "styled": 1,
            "ignored": 2,
            "out_of_scope": 3,
            "unresolved": 0,
        },
    }
    target_result = result.targets[0]
    assert target_result.source_sha256 == hashlib.sha256(b"target-original").hexdigest()
    assert target_result.output_sha256 == hashlib.sha256(
        b"formatted:target-original"
    ).hexdigest()
    assert target_result.audit_path is not None
    audit = json.loads(target_result.audit_path.read_text(encoding="utf-8"))
    assert audit["schema_version"] == 2
    assert audit["disposition_counts"] == target_result.audit_summary
    assert audit["application_audit"]["paragraph_indices"] == [0]
    assert "original_text_preview" not in manifest_text
    assert "never persist paragraph text" not in target_result.audit_path.read_text(
        encoding="utf-8"
    )
    assert audit["numbering_checks"] == {"checked": 1, "preserved": True}
    assert run_log_text.count("Applied classifications, stability verified") == 1
    assert run_log_text.count("TARGET target.docx: succeeded") == 1
    assert f"OUTPUT: {target_result.output_path}" in run_log_text
    assert "sf-" not in run_log_text
    assert "[untrusted detail omitted" not in run_log_text
    assert not any(
        str(line).lstrip().startswith("Output:") for line in target_result.log
    )
    assert manifest["targets"][0]["stage"] == "complete"


def test_run_publishes_structured_diagnostics_alongside_the_manifest(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    architect = _write_input(tmp_path / "architect.docx", b"architect-original")
    target = _write_input(tmp_path / "target.docx", b"target-original")
    calls, analyzer, config_loader, processor = _fake_dependencies(monkeypatch)

    result = pipeline.format_specifications(
        architect,
        [target],
        tmp_path / "formatted",
        api_key="offline-test-key",
        cache_dir=tmp_path / "profile-cache",
        diagnostics_level="debug",
        _template_analyzer=analyzer,
        _config_loader=config_loader,
        _target_processor=processor,
    )

    # A diagnostics stream is published next to run.json and surfaced on the result.
    assert result.diagnostics_path == result.run_dir / "diagnostics.jsonl"
    assert result.diagnostics_path.is_file()
    events = [
        json.loads(line)
        for line in result.diagnostics_path.read_text(encoding="utf-8").splitlines()
        if line.strip()
    ]
    assert events, "expected at least one diagnostics event"
    for event in events:
        assert set(("seq", "ts", "level", "component", "event", "fields")) <= set(event)
    by_event = {event["event"]: event for event in events}
    assert "run_start" in by_event
    assert by_event["run_start"]["fields"]["targets"] == 1
    assert "run_complete" in by_event
    assert by_event["run_complete"]["fields"] == {
        "targets": 1,
        "succeeded": 1,
        "failed": 0,
    }
    # Sequence numbers are strictly increasing.
    assert [event["seq"] for event in events] == sorted(event["seq"] for event in events)

    manifest = json.loads(result.manifest_path.read_text(encoding="utf-8"))
    diagnostics = manifest["diagnostics"]
    assert diagnostics["level"] == "DEBUG"
    assert diagnostics["log"] == "diagnostics.jsonl"
    assert diagnostics["event_count"] == len(events)
    assert diagnostics["errors"] == 0
    assert "phase_durations_ms" in diagnostics
    assert manifest["paths"]["diagnostics_log"] == str(result.diagnostics_path)
    # The additive block must not perturb the exact-equality summary contract.
    assert manifest["summary"] == {
        "targets": 1,
        "succeeded": 1,
        "failed": 0,
        "dispositions": {"styled": 1, "ignored": 2, "out_of_scope": 3, "unresolved": 0},
    }


def test_diagnostics_level_honours_environment_override(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    architect = _write_input(tmp_path / "architect.docx", b"architect-original")
    target = _write_input(tmp_path / "target.docx", b"target-original")
    calls, analyzer, config_loader, processor = _fake_dependencies(monkeypatch)
    monkeypatch.setenv("SPEC_FORMATTER_DIAGNOSTICS_LEVEL", "warning")

    result = pipeline.format_specifications(
        architect,
        [target],
        tmp_path / "formatted",
        api_key="offline-test-key",
        cache_dir=tmp_path / "profile-cache",
        diagnostics_level="info",
        _template_analyzer=analyzer,
        _config_loader=config_loader,
        _target_processor=processor,
    )

    manifest = json.loads(result.manifest_path.read_text(encoding="utf-8"))
    # The environment override wins over the call argument, so quieter INFO
    # phase events are filtered out of the persisted stream.
    assert manifest["diagnostics"]["level"] == "WARNING"
    diag_lines = [
        line
        for line in result.diagnostics_path.read_text(encoding="utf-8").splitlines()
        if line.strip()
    ]
    assert diag_lines == []


def test_initialization_failure_still_publishes_a_diagnostics_stream(
    tmp_path: Path,
) -> None:
    architect = _write_input(tmp_path / "architect.docx", b"architect-original")
    target = _write_input(tmp_path / "target.docx", b"target-original")
    api_key = "never-persist-this-secret"

    def failing_analyzer(**_kwargs):
        raise RuntimeError(f"template text=<w:t>CONFIDENTIAL {api_key}</w:t>")

    with pytest.raises(RuntimeError) as caught:
        pipeline.format_specifications(
            architect,
            [target],
            tmp_path / "formatted",
            api_key=api_key,
            cache_dir=tmp_path / "profile-cache",
            _template_analyzer=failing_analyzer,
        )

    run_dir = caught.value.run_dir
    diagnostics_path = run_dir / "diagnostics.jsonl"
    assert diagnostics_path.is_file()
    text = diagnostics_path.read_text(encoding="utf-8")
    assert api_key not in text
    assert "CONFIDENTIAL" not in text
    events = [json.loads(line) for line in text.splitlines() if line.strip()]
    by_event = {event["event"]: event for event in events}
    assert "init_failed" in by_event
    # Only the exception type name is recorded, never its (leaky) message.
    assert by_event["init_failed"]["fields"]["error_type"] == "runtimeerror"
    manifest = json.loads((run_dir / "run.json").read_text(encoding="utf-8"))
    assert manifest["diagnostics"]["errors"] >= 1
    assert manifest["paths"]["diagnostics_log"] == str(diagnostics_path)


def test_run_artifacts_strip_document_text_from_logs_errors_and_audits(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    architect = _write_input(tmp_path / "architect.docx", b"architect-original")
    target = _write_input(tmp_path / "target.docx", b"target-original")
    calls, analyzer, config_loader, _processor = _fake_dependencies(monkeypatch)
    api_key = "never-persist-this-secret"
    report = CanadianConversionReport(
        paragraphs_examined=1,
        paragraphs_converted=1,
        literal_markers_removed=1,
        automatic_numbering_retargeted=0,
        unnumbered_paragraphs_numbered=0,
        edits=(
            MarkerEdit(
                paragraph_index=4,
                role="PART",
                source_kind="literal",
                target_kind="automatic",
                source_marker="PART 9 - CONFIDENTIAL",
                target_marker="Part 9",
            ),
        ),
        warnings=(
            ConversionIssue(
                paragraph_index=4,
                code="marker_warning",
                message="TOP SECRET PARAGRAPH",
                text_preview="TOP SECRET PARAGRAPH",
            ),
        ),
    )

    def processor(**kwargs) -> BatchResult:
        return BatchResult(
            filename=Path(kwargs["docx_path"]).name,
            success=False,
            output_path=None,
            log=[
                "Classifying target",
                f'<w:p><w:r><w:t>SECRET DOC TEXT {api_key}</w:t></w:r></w:p>',
            ],
            error=f"paragraph_xml=<w:p>SECRET DOC TEXT {api_key}</w:p>",
            duration_seconds=0.01,
            conversion_report=report,
            audit={
                "ignored_paragraphs": [
                    {
                        "paragraph_index": 4,
                        "reason": "TOP SECRET PARAGRAPH",
                        "original_text_preview": "TOP SECRET PARAGRAPH",
                    }
                ],
            },
            numbering_checks={"policy": "format_only", "checked": 1},
            # A hostile processor tries to smuggle secrets/document text through
            # the structured diagnostics channel; both write paths must scrub it.
            diagnostics=[
                {
                    "level": "info",
                    "component": "target",
                    "event": "apply_classifications",
                    "fields": {
                        "modified": 3,
                        "leaked_secret": f"SECRET DOC TEXT {api_key}",
                        "preview": "TOP SECRET PARAGRAPH",
                    },
                }
            ],
        )

    result = pipeline.format_specifications(
        architect,
        [target],
        tmp_path / "formatted",
        api_key=api_key,
        cache_dir=tmp_path / "profile-cache",
        _template_analyzer=analyzer,
        _config_loader=config_loader,
        _target_processor=processor,
    )

    assert result.failed == 1
    assert result.manifest_path is not None
    assert result.targets[0].audit_path is not None
    assert result.diagnostics_path is not None and result.diagnostics_path.is_file()
    artifact_texts = [
        result.manifest_path.read_text(encoding="utf-8"),
        (result.run_dir / "run.log").read_text(encoding="utf-8"),
        result.targets[0].audit_path.read_text(encoding="utf-8"),
        result.diagnostics_path.read_text(encoding="utf-8"),
    ]
    for artifact_text in artifact_texts:
        assert api_key not in artifact_text
        assert "SECRET DOC TEXT" not in artifact_text
        assert "TOP SECRET PARAGRAPH" not in artifact_text
        assert "PART 9 - CONFIDENTIAL" not in artifact_text
        assert "leaked_secret" not in artifact_text

    manifest = json.loads(artifact_texts[0])
    assert manifest["targets"][0]["error_code"] == "untrusted_error"
    assert manifest["targets"][0]["error"] == "[document content omitted]"
    audit = json.loads(artifact_texts[2])
    assert audit["application_audit"]["ignored_paragraphs"] == [
        {"paragraph_index": 4, "reason": "unspecified"}
    ]
    assert audit["conversion_report"]["warnings"] == [
        {"paragraph_index": 4, "code": "marker_warning"}
    ]
    assert audit["conversion_report"]["edits"] == [
        {
            "paragraph_index": 4,
            "role": "PART",
            "source_kind": "literal",
            "target_kind": "automatic",
        }
    ]
    assert "[document content omitted]" in artifact_texts[1]
    assert artifact_texts[1].count("Classifying target") == 1


def test_diagnostics_channel_resists_key_and_token_exfiltration(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    architect = _write_input(tmp_path / "architect.docx", b"architect-original")
    target = _write_input(tmp_path / "target.docx", b"target-original")
    calls, analyzer, config_loader, _processor = _fake_dependencies(monkeypatch)
    api_key = "sk-ant-api03-" + "Z9y8X7w6" * 12

    def processor(**kwargs) -> BatchResult:
        source = Path(kwargs["docx_path"])
        staging_dir = Path(kwargs["output_dir"])
        staging_dir.mkdir(parents=True, exist_ok=True)
        staged_output = staging_dir / f"{source.stem}_PHASE2_FORMATTED.docx"
        staged_output.write_bytes(b"formatted:" + source.read_bytes())
        return BatchResult(
            filename=source.name,
            success=True,
            output_path=staged_output,
            log=["Applied classifications, stability verified"],
            error=None,
            duration_seconds=0.02,
            audit_summary={"styled": 1, "ignored": 0, "out_of_scope": 0, "unresolved": 0},
            # A hostile processor tries every diagnostics smuggling channel:
            # a secret as a dict KEY, a document token as a KEY, and a heading
            # tokenised across arbitrary keys and a list.
            diagnostics=[
                {
                    "level": "info",
                    "component": "target",
                    "event": "apply_classifications",
                    "fields": {
                        "modified": 3,
                        "num_id_remaps": {api_key: 2, "CONFIDENTIAL": 4},
                        api_key: 1,
                        "w0": "ACME",
                        "w1": "MERGER",
                        "leak_list": ["TOPSECRET", "PROJECT"],
                    },
                }
            ],
        )

    result = pipeline.format_specifications(
        architect,
        [target],
        tmp_path / "formatted",
        api_key=api_key,
        cache_dir=tmp_path / "profile-cache",
        diagnostics_level="debug",
        _template_analyzer=analyzer,
        _config_loader=config_loader,
        _target_processor=processor,
    )

    artifact_texts = [
        result.manifest_path.read_text(encoding="utf-8"),
        (result.run_dir / "run.log").read_text(encoding="utf-8"),
        result.targets[0].audit_path.read_text(encoding="utf-8"),
        result.diagnostics_path.read_text(encoding="utf-8"),
    ]
    for artifact_text in artifact_texts:
        assert api_key not in artifact_text
        assert "CONFIDENTIAL" not in artifact_text
        assert "ACME" not in artifact_text
        assert "MERGER" not in artifact_text
        assert "TOPSECRET" not in artifact_text
    # The one legitimate structural datum still comes through.
    diag_events = [
        json.loads(line)
        for line in artifact_texts[3].splitlines()
        if line.strip()
    ]
    apply_events = [e for e in diag_events if e["event"] == "apply_classifications"]
    assert apply_events
    fields = apply_events[0]["fields"]
    assert fields["modified"] == 3
    # Any container that held smuggled data was emptied, not merely value-scrubbed.
    assert fields.get("num_id_remaps", {}) == {}
    assert fields.get("leak_list", []) == []


def test_run_text_sanitizer_rejects_unstructured_model_authored_details() -> None:
    confidential = "deterministic override got confidential sprinkler layout"

    sanitized = pipeline._sanitize_run_text(confidential, ())

    assert confidential not in sanitized
    assert sanitized.startswith("[untrusted detail omitted; sha256=")
    assert pipeline._sanitize_run_text(
        "Applied classifications, stability verified",
        (),
        allow_operational=True,
    ) == "Applied classifications, stability verified"


def test_safe_error_diagnostic_extracts_only_allowlisted_internal_message() -> None:
    canonical = (
        "Architect template has conflicting section shells; use one canonical "
        "page layout and default/even/first header-footer mapping."
    )
    wrapped = (
        "Preflight validation failed (page_layout semantic validation failed: "
        f"{canonical}) trailing payload must not persist"
    )

    diagnostic = pipeline.safe_error_diagnostic(wrapped)
    unknown = pipeline.safe_error_diagnostic(
        "unknown ValueError with confidential paragraph prose"
    )

    assert diagnostic == pipeline.SafeErrorDiagnostic(
        code="template_section_shell_conflict",
        message=canonical,
    )
    assert "Preflight" not in diagnostic.message
    assert "trailing payload" not in diagnostic.message
    assert unknown is not None
    assert unknown.code == "untrusted_error"
    assert "confidential paragraph prose" not in unknown.message
    assert unknown.message.startswith("[untrusted detail omitted; sha256=")

    additional_known = {
        "template_default_section_conflict": (
            "Architect template default section conflicts with its section chain."
        ),
        "template_duplicate_section_index": (
            "Architect template section chain has duplicate section_index values."
        ),
        "template_api_key_required": (
            "An Anthropic API key is required to analyze a new architect template."
        ),
    }
    for code, message in additional_known.items():
        assert pipeline.safe_error_diagnostic(message) == pipeline.SafeErrorDiagnostic(
            code=code,
            message=message,
        )


@pytest.mark.parametrize(
    ("raw_error", "expected_code", "expected_message"),
    [
        (
            r"Architect template does not exist: C:\Private\Client A\template.docx",
            "input_architect_missing",
            "Architect template does not exist.",
        ),
        (
            r"Target specification does not exist: C:\Private\Client A\secret.docx",
            "input_target_missing",
            "A selected target specification does not exist.",
        ),
        (
            "Select at least one target specification DOCX file.",
            "input_target_required",
            "Select at least one target specification DOCX file.",
        ),
        (
            "conversion_mode must be one of: csi_to_canadian, format_only",
            "invalid_conversion_mode",
            "conversion_mode must be one of: csi_to_canadian, format_only",
        ),
        (
            "max_workers must be an integer.",
            "invalid_max_workers",
            "max_workers must be an integer.",
        ),
        (
            "Output directory could not be created.",
            "output_create_failed",
            "Output directory could not be created.",
        ),
    ],
)
def test_safe_error_diagnostic_exposes_path_free_input_remediation(
    raw_error: str,
    expected_code: str,
    expected_message: str,
) -> None:
    diagnostic = pipeline.safe_error_diagnostic(raw_error)

    assert diagnostic == pipeline.SafeErrorDiagnostic(
        code=expected_code,
        message=expected_message,
    )
    assert "Private" not in diagnostic.message
    assert "secret.docx" not in diagnostic.message


def test_progress_uses_queued_then_actual_processing_events(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    architect = _write_input(tmp_path / "architect.docx", b"architect-original")
    target = _write_input(tmp_path / "target.docx", b"target-original")
    _calls, analyzer, config_loader, processor = _fake_dependencies(monkeypatch)
    progress_messages: list[str] = []

    _run_with_fakes(
        architect,
        [target],
        tmp_path / "formatted",
        analyzer=analyzer,
        config_loader=config_loader,
        processor=processor,
        progress=progress_messages.append,
    )

    assert "Queued 1 of 1: target.docx" in progress_messages
    assert "Processing target 1 of 1: target.docx" in progress_messages
    assert not any(message.startswith("Started ") for message in progress_messages)


def test_target_progress_callbacks_keep_calling_thread_affinity(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    architect = _write_input(tmp_path / "architect.docx", b"architect-original")
    targets = [
        _write_input(tmp_path / f"target-{index}.docx", f"target-{index}".encode())
        for index in range(4)
    ]
    _calls, analyzer, config_loader, processor = _fake_dependencies(monkeypatch)
    calling_thread = threading.get_ident()
    callback_threads: list[int] = []

    _run_with_fakes(
        architect,
        targets,
        tmp_path / "formatted",
        analyzer=analyzer,
        config_loader=config_loader,
        processor=processor,
        progress=lambda _message: callback_threads.append(threading.get_ident()),
    )

    assert callback_threads
    assert set(callback_threads) == {calling_thread}


def test_timestamped_progress_events_are_ordered_and_match_persisted_log(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    architect = _write_input(tmp_path / "architect.docx", b"architect-original")
    targets = [
        _write_input(tmp_path / f"target-{index}.docx", f"target-{index}".encode())
        for index in range(4)
    ]
    _calls, analyzer, config_loader, processor = _fake_dependencies(monkeypatch)
    calling_thread = threading.get_ident()
    observed: list[tuple[str, datetime, int]] = []

    result = _run_with_fakes(
        architect,
        targets,
        tmp_path / "formatted",
        analyzer=analyzer,
        config_loader=config_loader,
        processor=processor,
        progress_event=lambda message, occurred_at: observed.append(
            (message, occurred_at, threading.get_ident())
        ),
    )

    messages = [message for message, _occurred_at, _thread_id in observed]
    event_times = [occurred_at for _message, occurred_at, _thread_id in observed]
    assert event_times == sorted(event_times)
    assert {thread_id for _message, _occurred_at, thread_id in observed} == {
        calling_thread
    }
    for index, target in enumerate(targets, start=1):
        queued = f"Queued {index} of {len(targets)}: {target.name}"
        processing = f"Processing target {index} of {len(targets)}: {target.name}"
        assert messages.index(queued) < messages.index(processing)

    run_log_lines = (result.run_dir / "run.log").read_text(
        encoding="utf-8"
    ).splitlines()
    persisted_events = [
        line for line in run_log_lines if pipeline._EVENT_TIMESTAMP_RX.match(line)
    ]
    persisted_times = [
        datetime.fromisoformat(line.split(" ", 1)[0].replace("Z", "+00:00"))
        for line in persisted_events
    ]
    assert persisted_times == sorted(persisted_times)
    assert [line.split(" ", 1)[1] for line in persisted_events] == messages


def test_pre_contract_cache_entry_is_ignored_once(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    architect = _write_input(tmp_path / "architect.docx", b"stable-architect")
    target = _write_input(tmp_path / "target.docx", b"stable-target")
    cache_dir = tmp_path / "profile-cache"
    source_hash = pipeline.template_analysis.sha256_file(architect)
    old_bundle = cache_dir / f"architect--{source_hash[:12]}--old.phase1"
    old_bundle.mkdir(parents=True)
    (old_bundle / "source.sha256").write_text(source_hash, encoding="ascii")
    calls, analyzer, config_loader, processor = _fake_dependencies(monkeypatch)

    result = _run_with_fakes(
        architect,
        [target],
        tmp_path / "formatted",
        cache_dir=cache_dir,
        analyzer=analyzer,
        config_loader=config_loader,
        processor=processor,
    )

    assert result.template_profile.reused is False
    assert len(calls["analyzer"]) == 1
    assert result.template_profile.bundle_dir.parent == (
        cache_dir.resolve() / pipeline._PROFILE_CACHE_NAMESPACE
    )
    assert old_bundle.is_dir()


def test_tampered_cached_template_profile_is_rejected_and_freshly_analyzed(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    architect = _write_input(tmp_path / "architect.docx", b"stable-architect")
    target = _write_input(tmp_path / "target.docx", b"stable-target")
    cache_dir = tmp_path / "profile-cache"
    calls, analyzer, config_loader, processor = _fake_dependencies(monkeypatch)
    progress_messages: list[str] = []

    first = _run_with_fakes(
        architect,
        [target],
        tmp_path / "formatted",
        cache_dir=cache_dir,
        analyzer=analyzer,
        config_loader=config_loader,
        processor=processor,
    )
    (first.template_profile.bundle_dir / "source.sha256").write_text(
        "0" * 64,
        encoding="ascii",
    )

    second = _run_with_fakes(
        architect,
        [target],
        tmp_path / "formatted",
        cache_dir=cache_dir,
        analyzer=analyzer,
        config_loader=config_loader,
        processor=processor,
        progress=progress_messages.append,
    )

    assert second.template_profile.reused is False
    assert second.template_profile.bundle_dir != first.template_profile.bundle_dir
    assert len(calls["analyzer"]) == 2
    assert len(calls["config_loader"]) == 2
    assert len(calls["processor"]) == 2
    assert any(
        message.startswith("Ignoring an invalid cached template profile")
        for message in progress_messages
    )
    assert architect.read_bytes() == b"stable-architect"
    assert target.read_bytes() == b"stable-target"


def test_changed_architect_bytes_invalidate_the_cached_profile(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    architect = _write_input(tmp_path / "architect.docx", b"architect-version-one")
    target = _write_input(tmp_path / "target.docx", b"stable-target")
    cache_dir = tmp_path / "profile-cache"
    calls, analyzer, config_loader, processor = _fake_dependencies(monkeypatch)

    first = _run_with_fakes(
        architect,
        [target],
        tmp_path / "formatted",
        cache_dir=cache_dir,
        analyzer=analyzer,
        config_loader=config_loader,
        processor=processor,
    )
    architect.write_bytes(b"architect-version-two")
    second = _run_with_fakes(
        architect,
        [target],
        tmp_path / "formatted",
        cache_dir=cache_dir,
        analyzer=analyzer,
        config_loader=config_loader,
        processor=processor,
    )

    assert first.template_profile.reused is False
    assert second.template_profile.reused is False
    assert first.template_profile.source_sha256 != second.template_profile.source_sha256
    assert first.template_profile.bundle_dir != second.template_profile.bundle_dir
    assert len(calls["analyzer"]) == 2


def test_input_preflight_fails_before_analyzer_or_other_pipeline_work(
    tmp_path: Path,
) -> None:
    architect = _write_input(tmp_path / "architect.docx", b"architect-original")
    missing_target = tmp_path / "missing.docx"
    output_dir = tmp_path / "formatted"
    cache_dir = tmp_path / "profile-cache"
    calls: list[str] = []

    def unexpected_dependency(*args, **kwargs):
        calls.append("called")
        raise AssertionError("pipeline dependency ran before input preflight completed")

    with pytest.raises(
        FileNotFoundError,
        match="Target specification does not exist",
    ) as raised:
        pipeline.format_specifications(
            architect,
            [missing_target],
            output_dir,
            api_key="offline-test-key",
            cache_dir=cache_dir,
            _template_analyzer=unexpected_dependency,
            _config_loader=unexpected_dependency,
            _target_processor=unexpected_dependency,
        )

    assert calls == []
    assert not output_dir.exists()
    assert not cache_dir.exists()
    assert architect.read_bytes() == b"architect-original"
    diagnostic = pipeline.safe_error_diagnostic(raised.value)
    assert diagnostic == pipeline.SafeErrorDiagnostic(
        code="input_target_missing",
        message="A selected target specification does not exist.",
    )
    assert str(missing_target) not in diagnostic.message


def test_isolated_run_output_cannot_overwrite_the_architect_input(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    architect = _write_input(
        tmp_path / "mechanical_FORMATTED.docx",
        b"architect-original",
    )
    target = _write_input(tmp_path / "mechanical.docx", b"target-original")
    calls, analyzer, config_loader, processor = _fake_dependencies(monkeypatch)

    result = _run_with_fakes(
        architect,
        [target],
        tmp_path,
        analyzer=analyzer,
        config_loader=config_loader,
        processor=processor,
    )

    assert result.success
    assert result.output_paths[0] != architect
    assert result.output_paths[0].parent == result.run_dir
    assert architect.read_bytes() == b"architect-original"
    assert target.read_bytes() == b"target-original"


def test_target_failure_is_isolated_and_successful_output_is_retained(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    architect = _write_input(tmp_path / "architect.docx", b"architect-original")
    good = _write_input(tmp_path / "good.docx", b"good-original")
    bad = _write_input(tmp_path / "bad.docx", b"bad-original")
    originals = {path: path.read_bytes() for path in (architect, good, bad)}
    calls, analyzer, config_loader, processor = _fake_dependencies(
        monkeypatch,
        failing_target=bad.name,
    )

    result = _run_with_fakes(
        architect,
        [good, bad],
        tmp_path / "formatted",
        analyzer=analyzer,
        config_loader=config_loader,
        processor=processor,
    )
    by_name = {item.source_path.name: item for item in result.targets}

    assert result.success is False
    assert result.succeeded == 1
    assert result.failed == 1
    assert by_name["good.docx"].success is True
    assert by_name["good.docx"].output_path is not None
    assert by_name["good.docx"].output_path.read_bytes() == b"formatted:good-original"
    assert by_name["bad.docx"].success is False
    assert by_name["bad.docx"].output_path is None
    assert by_name["bad.docx"].error == "simulated target failure"
    assert result.output_paths == (by_name["good.docx"].output_path,)
    assert {path: path.read_bytes() for path in originals} == originals


def test_known_target_failure_persists_actionable_code_message_and_stage(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    architect = _write_input(tmp_path / "architect.docx", b"architect-original")
    target = _write_input(tmp_path / "target.docx", b"target-original")
    _calls, analyzer, config_loader, _processor = _fake_dependencies(monkeypatch)
    canonical = (
        "Architect template has conflicting section shells; use one canonical "
        "page layout and default/even/first header-footer mapping."
    )
    wrapped = (
        "Preflight validation failed (page_layout semantic validation failed: "
        f"{canonical}) PRIVATE WRAPPER DETAIL"
    )

    def processor(**kwargs) -> BatchResult:
        return BatchResult(
            filename=Path(kwargs["docx_path"]).name,
            success=False,
            output_path=None,
            log=[f"FAILED: {wrapped}"],
            error=wrapped,
            duration_seconds=0.01,
            stage="page_layout",
        )

    result = _run_with_fakes(
        architect,
        [target],
        tmp_path / "formatted",
        analyzer=analyzer,
        config_loader=config_loader,
        processor=processor,
    )

    manifest = json.loads(result.manifest_path.read_text(encoding="utf-8"))
    audit = json.loads(result.targets[0].audit_path.read_text(encoding="utf-8"))
    run_log = (result.run_dir / "run.log").read_text(encoding="utf-8")
    for record in (manifest["targets"][0], audit):
        assert record["error_code"] == "template_section_shell_conflict"
        assert record["error"] == canonical
        assert record["stage"] == "page_layout"
    assert "ERROR [template_section_shell_conflict]" in run_log
    assert run_log.count("ERROR [template_section_shell_conflict]") == 1
    assert canonical in run_log
    assert "PRIVATE WRAPPER DETAIL" not in run_log


def test_target_processor_receives_an_isolated_snapshot(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    architect = _write_input(tmp_path / "architect.docx", b"architect-original")
    target = _write_input(tmp_path / "target.docx", b"target-original")
    calls, analyzer, config_loader, processor = _fake_dependencies(monkeypatch)
    received_paths: list[Path] = []

    def mutating_processor(**kwargs) -> BatchResult:
        received = Path(kwargs["docx_path"])
        received_paths.append(received)
        received.write_bytes(b"processor-overwrote-snapshot")
        return processor(**kwargs)

    result = _run_with_fakes(
        architect,
        [target],
        tmp_path / "formatted",
        analyzer=analyzer,
        config_loader=config_loader,
        processor=mutating_processor,
    )

    assert result.success is True
    assert len(received_paths) == 1
    assert received_paths[0] != target
    assert calls["processor"] == received_paths
    assert result.output_paths[0].read_bytes() == (
        b"formatted:processor-overwrote-snapshot"
    )
    assert architect.read_bytes() == b"architect-original"
    assert target.read_bytes() == b"target-original"


def test_same_stem_targets_receive_distinct_traceable_output_names(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    architect = _write_input(tmp_path / "architect.docx", b"architect-original")
    mechanical = _write_input(
        tmp_path / "mechanical" / "section.docx",
        b"mechanical-section",
    )
    electrical = _write_input(
        tmp_path / "electrical" / "section.docx",
        b"electrical-section",
    )
    calls, analyzer, config_loader, processor = _fake_dependencies(monkeypatch)

    result = _run_with_fakes(
        architect,
        [mechanical, electrical],
        tmp_path / "formatted",
        analyzer=analyzer,
        config_loader=config_loader,
        processor=processor,
    )
    output_names = {path.name for path in result.output_paths}

    assert result.success is True
    assert len(output_names) == 2
    assert all(name.startswith("section__") for name in output_names)
    assert all(name.endswith("_FORMATTED.docx") for name in output_names)
    assert any("mechanical" in name for name in output_names)
    assert any("electrical" in name for name in output_names)
    assert mechanical.read_bytes() == b"mechanical-section"
    assert electrical.read_bytes() == b"electrical-section"


def test_long_windows_output_components_are_truncated_and_hashed(tmp_path: Path) -> None:
    stem = "section-" + ("x" * 242)
    target = tmp_path / f"{stem}.docx"

    first = pipeline._plan_output_paths(
        [target],
        tmp_path / "formatted",
        pipeline.FORMAT_ONLY,
    )[target]
    second = pipeline._plan_output_paths(
        [target],
        tmp_path / "formatted",
        pipeline.FORMAT_ONLY,
    )[target]

    assert first.name == second.name
    assert first.name.endswith("_FORMATTED.docx")
    assert pipeline._utf16_code_units(first.name) <= 240
    assert stem not in first.name
    assert "__" in first.stem


def test_folder_discovery_excludes_lock_files_and_current_or_legacy_outputs(
    tmp_path: Path,
) -> None:
    specs = tmp_path / "specs"
    target = _write_input(specs / "target.docx", b"target")
    _write_input(specs / "target_FORMATTED.docx", b"current-output")
    _write_input(specs / "target_PHASE2_FORMATTED.docx", b"legacy-output")
    _write_input(specs / "target_CANADIAN_FORMATTED.docx", b"canadian-output")
    _write_input(specs / "another_formatted.docx", b"case-insensitive-output")
    _write_input(specs / "~$target.docx", b"word-lock-file")
    _write_input(specs / "nested" / "nested.docx", b"non-recursive")

    discovered = pipeline.collect_target_specs([specs, target])

    assert discovered == (target.resolve(),)


def test_folder_discovery_excludes_architect_but_explicit_selection_is_rejected(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    specs = tmp_path / "specs"
    architect = _write_input(specs / "architect.docx", b"architect-original")
    target = _write_input(specs / "target.docx", b"target-original")
    calls, analyzer, config_loader, processor = _fake_dependencies(monkeypatch)

    result = _run_with_fakes(
        architect,
        [specs],
        tmp_path / "formatted",
        analyzer=analyzer,
        config_loader=config_loader,
        processor=processor,
    )

    assert result.success
    assert [item.source_path for item in result.targets] == [target.resolve()]
    with pytest.raises(
        ValueError,
        match="architect template cannot also be a target",
    ):
        pipeline.format_specifications(
            architect,
            [specs, architect],
            tmp_path / "formatted-explicit",
            api_key="offline-test-key",
            cache_dir=tmp_path / "profile-cache-explicit",
            _template_analyzer=analyzer,
            _config_loader=config_loader,
            _target_processor=processor,
        )


def test_canadian_mode_reaches_target_processor_and_uses_distinct_output_name(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    architect = _write_input(tmp_path / "architect.docx", b"architect-original")
    target = _write_input(tmp_path / "target.docx", b"target-original")
    calls, analyzer, config_loader, processor = _fake_dependencies(monkeypatch)
    received_modes: list[str] = []

    def capturing_processor(**kwargs) -> BatchResult:
        received_modes.append(kwargs["conversion_mode"])
        return processor(**kwargs)

    result = pipeline.format_specifications(
        architect,
        [target],
        tmp_path / "formatted",
        api_key="offline-test-key",
        cache_dir=tmp_path / "profile-cache",
        max_workers=1,
        conversion_mode=pipeline.CSI_TO_CANADIAN,
        _template_analyzer=analyzer,
        _config_loader=config_loader,
        _target_processor=capturing_processor,
    )

    assert result.success
    assert received_modes == [pipeline.CSI_TO_CANADIAN]
    assert result.output_paths[0].name == "target_CANADIAN_FORMATTED.docx"
    assert target.read_bytes() == b"target-original"


def test_invalid_conversion_mode_fails_before_analysis_or_filesystem_writes(
    tmp_path: Path,
) -> None:
    architect = _write_input(tmp_path / "architect.docx", b"architect-original")
    target = _write_input(tmp_path / "target.docx", b"target-original")
    calls: list[str] = []

    def unexpected_dependency(*args, **kwargs):
        calls.append("called")
        raise AssertionError("pipeline work started for an invalid conversion mode")

    with pytest.raises(ValueError, match="conversion_mode") as raised:
        pipeline.format_specifications(
            architect,
            [target],
            tmp_path / "formatted",
            api_key="offline-test-key",
            conversion_mode="not-a-real-mode",
            _template_analyzer=unexpected_dependency,
            _config_loader=unexpected_dependency,
            _target_processor=unexpected_dependency,
        )

    assert calls == []
    assert not (tmp_path / "formatted").exists()
    assert pipeline.safe_error_diagnostic(raised.value) == (
        pipeline.SafeErrorDiagnostic(
            code="invalid_conversion_mode",
            message="conversion_mode must be one of: csi_to_canadian, format_only",
        )
    )


@pytest.mark.parametrize("invalid_workers", [True, 1.5, "three", None])
def test_invalid_max_workers_has_stable_diagnostic_before_filesystem_writes(
    tmp_path: Path,
    invalid_workers: object,
) -> None:
    architect = _write_input(tmp_path / "architect.docx", b"architect-original")
    target = _write_input(tmp_path / "target.docx", b"target-original")

    with pytest.raises(ValueError, match="max_workers must be an integer") as raised:
        pipeline.format_specifications(
            architect,
            [target],
            tmp_path / "formatted",
            api_key="offline-test-key",
            max_workers=invalid_workers,  # type: ignore[arg-type]
        )

    assert not (tmp_path / "formatted").exists()
    assert pipeline.safe_error_diagnostic(raised.value) == (
        pipeline.SafeErrorDiagnostic(
            code="invalid_max_workers",
            message="max_workers must be an integer.",
        )
    )


def test_output_directory_creation_failure_has_path_free_diagnostic(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    architect = _write_input(tmp_path / "architect.docx", b"architect-original")
    target = _write_input(tmp_path / "target.docx", b"target-original")
    output_dir = (tmp_path / "Private Client" / "formatted").resolve()
    original_mkdir = Path.mkdir

    def fail_output_mkdir(path: Path, *args, **kwargs) -> None:
        if path == output_dir:
            raise PermissionError(f"denied: {path}")
        original_mkdir(path, *args, **kwargs)

    monkeypatch.setattr(Path, "mkdir", fail_output_mkdir)

    with pytest.raises(OSError, match="could not be created") as raised:
        pipeline.format_specifications(
            architect,
            [target],
            output_dir,
            api_key="offline-test-key",
        )

    diagnostic = pipeline.safe_error_diagnostic(raised.value)
    assert diagnostic == pipeline.SafeErrorDiagnostic(
        code="output_create_failed",
        message="Output directory could not be created.",
    )
    assert str(output_dir) not in diagnostic.message
    assert isinstance(raised.value.__cause__, PermissionError)


def test_template_initialization_failure_still_publishes_failed_run_artifacts(
    tmp_path: Path,
) -> None:
    architect = _write_input(tmp_path / "architect.docx", b"architect-original")
    target = _write_input(tmp_path / "target.docx", b"target-original")
    api_key = "never-persist-this-secret"

    def failing_analyzer(**_kwargs):
        raise RuntimeError(
            f"template text=<w:t>CONFIDENTIAL BODY {api_key}</w:t>"
        )

    with pytest.raises(RuntimeError) as caught:
        pipeline.format_specifications(
            architect,
            [target],
            tmp_path / "formatted",
            api_key=api_key,
            cache_dir=tmp_path / "profile-cache",
            _template_analyzer=failing_analyzer,
        )

    run_dir = caught.value.run_dir
    manifest_path = caught.value.manifest_path
    assert run_dir.parent == (tmp_path / "formatted").resolve()
    assert manifest_path == run_dir / "run.json"
    assert (run_dir / "run.log").is_file()
    manifest_text = manifest_path.read_text(encoding="utf-8")
    audit_path = next(run_dir.glob("target-*.audit.json"))
    combined = manifest_text + (run_dir / "run.log").read_text(
        encoding="utf-8"
    ) + audit_path.read_text(encoding="utf-8")
    assert api_key not in combined
    assert "CONFIDENTIAL BODY" not in combined
    manifest = json.loads(manifest_text)
    assert manifest["status"] == "failed"
    assert manifest["failure_phase"] == "initialization"
    assert manifest["summary"]["failed"] == 1
    assert manifest["targets"][0]["audit_path"] == str(audit_path)


def test_publication_failure_preserves_processor_log_and_conversion_report(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    target = _write_input(tmp_path / "target.docx", b"source")
    staging = tmp_path / "staging"
    report = CanadianConversionReport(
        paragraphs_examined=1,
        paragraphs_converted=1,
        literal_markers_removed=1,
        automatic_numbering_retargeted=0,
        unnumbered_paragraphs_numbered=0,
        edits=(),
        warnings=(),
    )

    def processor(**kwargs) -> BatchResult:
        output = Path(kwargs["output_dir"]) / "target_CANADIAN_FORMATTED.docx"
        output.parent.mkdir(parents=True, exist_ok=True)
        output.write_bytes(b"formatted")
        return BatchResult(
            filename="target.docx",
            success=True,
            output_path=output,
            log=["conversion complete"],
            error=None,
            duration_seconds=0.1,
            conversion_report=report,
        )

    monkeypatch.setattr(
        pipeline.os,
        "replace",
        lambda *_args, **_kwargs: (_ for _ in ()).throw(OSError("publish failed")),
    )
    result = pipeline._format_one_target(
        target,
        tmp_path / "final.docx",
        staging,
        SharedConfig({}, {}, "", [], {}, tmp_path, {}),
        "",
        "test-model",
        processor,
        pipeline.CSI_TO_CANADIAN,
    )

    assert result.success is False
    assert result.conversion_report is report
    assert result.log == ("conversion complete", "FAILED: publish failed")
    assert result.stage == "publication"


def test_processor_exception_records_processing_stage(tmp_path: Path) -> None:
    target = _write_input(tmp_path / "target.docx", b"source")

    def processor(**_kwargs):
        raise RuntimeError("injected processor failure")

    result = pipeline._format_one_target(
        target,
        tmp_path / "final.docx",
        tmp_path / "staging",
        SharedConfig({}, {}, "", [], {}, tmp_path, {}),
        "",
        "test-model",
        processor,
        pipeline.FORMAT_ONLY,
    )

    assert result.success is False
    assert result.stage == "processing"


def test_cached_profile_with_a_stale_engine_fingerprint_is_not_reused(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    prompt_dir = Path(pipeline.__file__).resolve().parents[1]
    producer = {
        "name": "spec-template-normalizer",
        "version": pipeline.template_analysis.PIPELINE_VERSION,
        "engine_fingerprint": pipeline.template_analysis.ENGINE_SOURCE_DIGEST,
        "classifier": {
            "provider": "anthropic",
            "model": pipeline.template_analysis.DEFAULT_MODEL,
        },
        "prompts": pipeline._prompt_fingerprints(prompt_dir)
        if hasattr(pipeline, "_prompt_fingerprints")
        else {
            "master_prompt_sha256": hashlib.sha256(
                (prompt_dir / "master_prompt.txt").read_text(encoding="utf-8").encode("utf-8")
            ).hexdigest(),
            "run_instruction_sha256": hashlib.sha256(
                (prompt_dir / "run_instruction_prompt.txt").read_text(encoding="utf-8").encode("utf-8")
            ).hexdigest(),
        },
    }
    current = SimpleNamespace(producer=dict(producer))
    stale = SimpleNamespace(producer={**producer, "engine_fingerprint": "deadbeefdeadbeef"})
    common = dict(
        model=pipeline.template_analysis.DEFAULT_MODEL,
        prompt_dir=prompt_dir,
        classifier=None,
    )

    assert pipeline._manifest_matches_current_engine(current, **common) is True
    assert pipeline._manifest_matches_current_engine(stale, **common) is False


def test_stale_profiles_for_one_template_are_pruned_to_the_newest_two(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    architect = _write_input(tmp_path / "architect.docx", b"architect-original")
    calls, analyzer, _config_loader, _processor = _fake_dependencies(monkeypatch)
    cache_dir = tmp_path / "profile-cache"
    messages: list[str] = []

    profiles = [
        pipeline.prepare_template_profile(
            architect,
            cache_dir,
            "key",
            force_analysis=True,
            analyzer=analyzer,
            progress=messages.append,
        )
        for _ in range(3)
    ]

    namespace = cache_dir / pipeline._PROFILE_CACHE_NAMESPACE
    remaining = sorted(p.name for p in namespace.glob("*.phase1"))
    assert len(calls["analyzer"]) == 3
    assert len(remaining) == 2
    # The profile just selected is always kept, and so is the next newest.
    assert profiles[-1].bundle_dir.name in remaining
    assert profiles[-2].bundle_dir.name in remaining
    assert profiles[0].bundle_dir.name not in remaining
    assert any("Removed 1 older cached profile" in line for line in messages)
    assert not any(profiles[0].bundle_dir.name in line for line in messages)


def test_profile_provenance_is_captured_at_selection_not_revalidated_at_publication(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    architect = _write_input(tmp_path / "architect.docx", b"architect-original")
    target = _write_input(tmp_path / "target.docx", b"target-original")
    _calls, analyzer, config_loader, processor = _fake_dependencies(monkeypatch)
    validator = pipeline.template_analysis.validate_bundle_directory
    validations: list[Path] = []

    def counting_validator(bundle_dir, *, expected_source_sha256):
        validations.append(Path(bundle_dir))
        return validator(bundle_dir, expected_source_sha256=expected_source_sha256)

    monkeypatch.setattr(pipeline.template_analysis, "validate_bundle_directory", counting_validator)

    result = _run_with_fakes(
        architect,
        [target],
        tmp_path / "formatted",
        analyzer=analyzer,
        config_loader=config_loader,
        processor=processor,
    )

    assert result.template_profile.provenance is not None
    assert result.template_profile.provenance["producer"]["name"] == "spec-template-normalizer"
    # One validation when the fresh profile is selected; none while writing run.json.
    assert len(validations) == 1
    manifest = json.loads(result.manifest_path.read_text(encoding="utf-8"))
    assert manifest["template_profile"]["producer"]["name"] == "spec-template-normalizer"
    assert manifest["template_profile"]["reused"] is False


def test_publication_failure_still_writes_a_failed_manifest_with_target_outcomes(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    architect = _write_input(tmp_path / "architect.docx", b"architect-original")
    target = _write_input(tmp_path / "target.docx", b"target-original")
    _calls, analyzer, config_loader, processor = _fake_dependencies(monkeypatch)

    def failing_writer(**_kwargs):
        raise OSError("disk full while writing audits")

    monkeypatch.setattr(pipeline, "_write_run_artifacts", failing_writer)

    with pytest.raises(OSError) as raised:
        _run_with_fakes(
            architect,
            [target],
            tmp_path / "formatted",
            analyzer=analyzer,
            config_loader=config_loader,
            processor=processor,
        )

    run_dir = raised.value.run_dir
    manifest_path = raised.value.manifest_path
    assert run_dir.is_dir() and manifest_path == run_dir / "run.json"
    manifest = json.loads(manifest_path.read_text(encoding="utf-8"))
    assert manifest["status"] == "failed"
    assert manifest["failure_phase"] == "publication"
    assert "disk full" not in manifest_path.read_text(encoding="utf-8")
    record = manifest["targets"][0]
    assert record["success"] is True
    assert record["stage"] == "complete"
    assert record["output_path"] == str(run_dir / "target_FORMATTED.docx")
    assert (run_dir / "target_FORMATTED.docx").is_file()
    # The fallback audit and summary describe the same outcome as the
    # target record; they used to claim not_started / all failed.
    audit = json.loads(Path(record["audit_path"]).read_text(encoding="utf-8"))
    assert audit["success"] is True
    assert audit["stage"] == "complete"
    assert audit["output"]["path"] == record["output_path"]
    expected_counts = {"styled": 1, "ignored": 2, "out_of_scope": 3, "unresolved": 0}
    assert audit["disposition_counts"] == expected_counts
    assert record["disposition_counts"] == expected_counts
    assert record["numbering_checks"] == {"preserved": True, "checked": 1}
    assert manifest["summary"] == {
        "targets": 1,
        "succeeded": 1,
        "failed": 0,
        "dispositions": expected_counts,
    }
    assert "RUN FAILED DURING PUBLICATION" in (run_dir / "run.log").read_text(encoding="utf-8")
    assert not (run_dir / ".staging").exists()


def test_staging_directory_is_removed_when_the_run_ends(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    architect = _write_input(tmp_path / "architect.docx", b"architect-original")
    target = _write_input(tmp_path / "target.docx", b"target-original")
    _calls, analyzer, config_loader, processor = _fake_dependencies(monkeypatch)
    observed: list[bool] = []
    original = pipeline._atomic_write_bytes

    def observing_write(path, payload):
        original(path, payload)
        observed.append((path.parent / ".staging").is_dir())

    monkeypatch.setattr(pipeline, "_atomic_write_bytes", observing_write)

    result = _run_with_fakes(
        architect,
        [target],
        tmp_path / "formatted",
        analyzer=analyzer,
        config_loader=config_loader,
        processor=processor,
    )

    assert observed and all(observed)
    assert not (result.run_dir / ".staging").exists()
    assert sorted(p.name for p in result.run_dir.iterdir() if p.name.startswith(".")) == []


def test_cached_profile_rejection_message_carries_no_raw_detail(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    architect = _write_input(tmp_path / "architect.docx", b"architect-original")
    target = _write_input(tmp_path / "target.docx", b"target-original")
    _calls, analyzer, config_loader, processor = _fake_dependencies(monkeypatch)
    cache_dir = tmp_path / "profile-cache"
    first = _run_with_fakes(
        architect,
        [target],
        tmp_path / "formatted",
        analyzer=analyzer,
        config_loader=config_loader,
        processor=processor,
        cache_dir=cache_dir,
    )
    private = r"C:\\Users\\Private\\secret-bundle"
    validator = pipeline.template_analysis.validate_bundle_directory

    def rejecting_validator(bundle_dir, *, expected_source_sha256):
        if Path(bundle_dir) == first.template_profile.bundle_dir:
            raise ValueError(f"manifest tampered at {private}")
        return validator(bundle_dir, expected_source_sha256=expected_source_sha256)

    monkeypatch.setattr(pipeline.template_analysis, "validate_bundle_directory", rejecting_validator)
    messages: list[str] = []

    _run_with_fakes(
        architect,
        [target],
        tmp_path / "formatted",
        analyzer=analyzer,
        config_loader=config_loader,
        processor=processor,
        cache_dir=cache_dir,
        progress=messages.append,
    )

    rejection = [line for line in messages if "Ignoring an invalid cached template profile" in line]
    assert rejection
    assert all(private not in line and "tampered" not in line for line in rejection)


def test_run_artifacts_prefer_the_engine_error_code_over_classified_text(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    from spec_formatter.style_application.core.errors import ERROR_REMEDIATIONS

    architect = _write_input(tmp_path / "architect.docx", b"architect-original")
    target = _write_input(tmp_path / "target.docx", b"target-original")
    _calls, analyzer, config_loader, _processor = _fake_dependencies(monkeypatch)
    raw_detail = "Paragraph 412 begins with ambiguous decimal text 'CONFIDENTIAL CLAUSE'"

    def coded_processor(**kwargs) -> BatchResult:
        return BatchResult(
            filename=Path(kwargs["docx_path"]).name,
            success=False,
            output_path=None,
            log=["FAILED: " + raw_detail],
            error=raw_detail,
            duration_seconds=0.01,
            stage="disposition_verification",
            error_code="classification_coverage_incomplete",
            safe_error=ERROR_REMEDIATIONS["classification_coverage_incomplete"],
        )

    result = _run_with_fakes(
        architect,
        [target],
        tmp_path / "formatted",
        analyzer=analyzer,
        config_loader=config_loader,
        processor=coded_processor,
    )

    item = result.targets[0]
    assert item.success is False
    assert item.error_code == "classification_coverage_incomplete"
    manifest_text = result.manifest_path.read_text(encoding="utf-8")
    audit_text = item.audit_path.read_text(encoding="utf-8")
    assert "CONFIDENTIAL CLAUSE" not in manifest_text
    assert "CONFIDENTIAL CLAUSE" not in audit_text
    manifest = json.loads(manifest_text)
    record = manifest["targets"][0]
    assert record["error_code"] == "classification_coverage_incomplete"
    assert record["error"] == ERROR_REMEDIATIONS["classification_coverage_incomplete"]
    assert record["stage"] == "disposition_verification"
    audit = json.loads(audit_text)
    assert audit["error_code"] == "classification_coverage_incomplete"
    assert audit["error"] == record["error"]


def test_target_error_diagnostic_falls_back_to_text_classification_without_a_code() -> None:
    from spec_formatter.pipeline import TargetFormatResult, _target_error_diagnostic

    coded = TargetFormatResult(
        source_path=Path("t.docx"), success=False, output_path=None, log=(),
        error="detail with CONFIDENTIAL text", duration_seconds=0.0,
        error_code="canadian_target_hierarchy",
    )
    uncoded = TargetFormatResult(
        source_path=Path("t.docx"), success=False, output_path=None, log=(),
        error="detail with CONFIDENTIAL text", duration_seconds=0.0,
    )
    unknown = TargetFormatResult(
        source_path=Path("t.docx"), success=False, output_path=None, log=(),
        error="detail with CONFIDENTIAL text", duration_seconds=0.0,
        error_code="not_a_real_code",
    )

    assert _target_error_diagnostic(coded, ()).code == "canadian_target_hierarchy"
    assert "CONFIDENTIAL" not in _target_error_diagnostic(coded, ()).message
    assert _target_error_diagnostic(uncoded, ()).code == "untrusted_error"
    assert _target_error_diagnostic(unknown, ()).code == "untrusted_error"


def test_package_exports_every_name_the_gui_imports() -> None:
    import spec_formatter

    for name in (
        "SafeErrorDiagnostic",
        "safe_error_diagnostic",
        "collect_target_specs",
        "prepare_template_profile",
        "format_specifications",
        "default_template_cache_dir",
    ):
        assert name in spec_formatter.__all__
        assert getattr(spec_formatter, name) is getattr(pipeline, name)


class _InlineExecutor:
    """A ThreadPoolExecutor stand-in whose futures are complete before submit returns.

    It makes the completion race deterministic: every ``done`` signal is queued
    (and drained by the submission loop) before the completion loop starts.
    """

    def __init__(self, max_workers=None):
        self.max_workers = max_workers

    def __enter__(self):
        return self

    def __exit__(self, *exc_info):
        return False

    def submit(self, fn, *args, **kwargs):
        from concurrent.futures import Future

        future: Future = Future()
        try:
            future.set_result(fn(*args, **kwargs))
        except BaseException as exc:  # pragma: no cover - surfaced by the pipeline
            future.set_exception(exc)
        return future


def test_completion_signals_drained_during_submission_are_not_lost(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    """A target that finishes before the completion loop starts is still collected.

    ``report`` drains the signal queue from the owner thread, so a target that
    completes while the next one is being queued has its ``done`` signal
    consumed early. The run used to block forever waiting for a signal that
    would never be queued again.
    """

    architect = _write_input(tmp_path / "architect.docx", b"architect-original")
    targets = [
        _write_input(tmp_path / f"target{index}.docx", f"target{index}-original".encode())
        for index in range(3)
    ]
    _calls, analyzer, config_loader, processor = _fake_dependencies(monkeypatch)
    monkeypatch.setattr(pipeline, "ThreadPoolExecutor", _InlineExecutor)
    outcome: dict = {}

    def run() -> None:
        try:
            outcome["result"] = _run_with_fakes(
                architect,
                targets,
                tmp_path / "formatted",
                analyzer=analyzer,
                config_loader=config_loader,
                processor=processor,
            )
        except BaseException as exc:  # pragma: no cover - reported below
            outcome["error"] = exc

    worker = threading.Thread(target=run, daemon=True)
    worker.start()
    worker.join(timeout=30)

    assert not worker.is_alive(), "the run hung waiting for an already-drained completion"
    assert "error" not in outcome, outcome.get("error")
    result = outcome["result"]
    assert result.success and len(result.targets) == 3
    assert [item.success for item in result.targets] == [True, True, True]
    assert result.manifest_path.is_file()


def test_architect_and_target_usage_reach_run_json_at_every_verbosity(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    """Usage recorded by the classifiers must survive into the canonical run.

    A fresh architect analysis reports its counts on Phase1Result, which
    prepare_template_profile used to drop on the floor, so the canonical
    entry point published no architect tokens at all. Target usage lived
    only on an INFO event, so a run at warning level published none either.
    """
    for level in ("info", "warning"):
        architect = _write_input(tmp_path / f"arch-{level}.docx", b"architect-original")
        target = _write_input(tmp_path / f"target-{level}.docx", b"target-original")
        calls, analyzer, config_loader, processor = _fake_dependencies(monkeypatch)

        def analyzer_with_usage(**kwargs):
            base = analyzer(**kwargs)
            return SimpleNamespace(
                bundle_dir=base.bundle_dir,
                usage={
                    "requests_attempted": 2,
                    "input_tokens": 40000,
                    "output_tokens": 900,
                    "usage_complete": True,
                },
            )

        base_processor = processor

        def processor_with_usage(**kwargs):
            result = base_processor(**kwargs)
            result.usage = {
                "requests_attempted": 1,
                "input_tokens": 1200,
                "output_tokens": 60,
                "usage_complete": True,
            }
            return result

        result = pipeline.format_specifications(
            architect,
            [target],
            tmp_path / f"formatted-{level}",
            api_key="offline-test-key",
            cache_dir=tmp_path / f"cache-{level}",
            diagnostics_level=level,
            _template_analyzer=analyzer_with_usage,
            _config_loader=config_loader,
            _target_processor=processor_with_usage,
        )

        manifest = json.loads(Path(result.manifest_path).read_text(encoding="utf-8"))
        usage = manifest["diagnostics"]["usage"]
        assert usage["architect"]["input_tokens"] == 40000, level
        assert usage["architect"]["output_tokens"] == 900, level
        assert usage["target"]["input_tokens"] == 1200, level
        assert usage["target"]["usage_complete"] is True, level
