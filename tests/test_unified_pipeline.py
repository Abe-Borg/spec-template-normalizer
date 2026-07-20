from __future__ import annotations

import hashlib
from pathlib import Path
from types import SimpleNamespace

import pytest

from spec_formatter import pipeline
from spec_formatter.style_application.batch_runner import BatchResult, SharedConfig


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
        if source.name == failing_target:
            return BatchResult(
                filename=source.name,
                success=False,
                output_path=None,
                log=["simulated target failure"],
                error="simulated target failure",
                duration_seconds=0.01,
            )

        staging_dir = Path(kwargs["output_dir"])
        staging_dir.mkdir(parents=True, exist_ok=True)
        staged_output = staging_dir / f"{source.stem}_PHASE2_FORMATTED.docx"
        staged_output.write_bytes(b"formatted:" + source.read_bytes())
        return BatchResult(
            filename=source.name,
            success=True,
            output_path=staged_output,
            log=["formatted offline"],
            error=None,
            duration_seconds=0.02,
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
    assert {path.name for path in calls["processor"]} == {
        "mechanical.docx",
        "electrical.docx",
    }
    assert {path.name for path in result.output_paths} == {
        "mechanical_FORMATTED.docx",
        "electrical_FORMATTED.docx",
    }
    assert all(path.is_file() for path in result.output_paths)
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
    assert len(calls["analyzer"]) == 1
    assert len(calls["config_loader"]) == 2
    assert len(calls["processor"]) == 2
    assert "Reusing the validated architect template analysis." in progress_messages
    assert architect.read_bytes() == b"stable-architect"
    assert target.read_bytes() == b"stable-target"


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

    with pytest.raises(FileNotFoundError, match="Target specification does not exist"):
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


def test_output_plan_cannot_overwrite_the_architect_input(
    tmp_path: Path,
) -> None:
    architect = _write_input(
        tmp_path / "mechanical_FORMATTED.docx",
        b"architect-original",
    )
    target = _write_input(tmp_path / "mechanical.docx", b"target-original")
    calls: list[str] = []

    def unexpected_dependency(*args, **kwargs):
        calls.append("called")
        raise AssertionError("pipeline work started despite an unsafe output plan")

    with pytest.raises(ValueError, match="would overwrite an input file"):
        pipeline.format_specifications(
            architect,
            [target],
            tmp_path,
            api_key="offline-test-key",
            _template_analyzer=unexpected_dependency,
            _config_loader=unexpected_dependency,
            _target_processor=unexpected_dependency,
        )

    assert calls == []
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


def test_folder_discovery_excludes_lock_files_and_current_or_legacy_outputs(
    tmp_path: Path,
) -> None:
    specs = tmp_path / "specs"
    target = _write_input(specs / "target.docx", b"target")
    _write_input(specs / "target_FORMATTED.docx", b"current-output")
    _write_input(specs / "target_PHASE2_FORMATTED.docx", b"legacy-output")
    _write_input(specs / "another_formatted.docx", b"case-insensitive-output")
    _write_input(specs / "~$target.docx", b"word-lock-file")
    _write_input(specs / "nested" / "nested.docx", b"non-recursive")

    discovered = pipeline.collect_target_specs([specs, target])

    assert discovered == (target.resolve(),)
