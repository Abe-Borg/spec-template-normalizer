"""Shipped resources resolve from one root, frozen or from a checkout."""

from __future__ import annotations

import sys
from pathlib import Path

import pytest

import phase1_pipeline
from spec_formatter import pipeline, resources

REPO_ROOT = Path(__file__).resolve().parents[1]


def _freeze(monkeypatch, root: Path) -> None:
    """Make the process look like a PyInstaller bundle rooted at *root*."""

    monkeypatch.setattr(sys, "frozen", True, raising=False)
    monkeypatch.setattr(sys, "_MEIPASS", str(root), raising=False)


def test_checkout_resource_root_is_the_repository():
    assert not resources.is_frozen()
    assert resources.resource_root() == REPO_ROOT
    for name in resources.ARCHITECT_PROMPT_FILES:
        assert (resources.architect_prompt_dir() / name).is_file(), name
    for name in resources.TARGET_PROMPT_FILES:
        assert (resources.target_prompt_dir() / name).is_file(), name
    assert resources.resource_path("LICENSE").is_file()
    assert resources.resource_path("THIRD_PARTY_NOTICES.md").is_file()


def test_frozen_resource_root_is_the_bundle_directory(tmp_path: Path, monkeypatch):
    _freeze(monkeypatch, tmp_path)
    bundle = tmp_path.resolve()
    assert resources.is_frozen()
    assert resources.resource_root() == bundle
    assert resources.architect_prompt_dir() == bundle
    assert resources.target_prompt_dir() == (
        bundle / "spec_formatter" / "style_application" / "core" / "prompts"
    )
    assert resources.resource_path("LICENSE") == bundle / "LICENSE"


def test_frozen_flag_without_a_bundle_directory_is_not_frozen(monkeypatch):
    monkeypatch.setattr(sys, "frozen", True, raising=False)
    monkeypatch.delattr(sys, "_MEIPASS", raising=False)
    assert not resources.is_frozen()
    assert resources.resource_root() == REPO_ROOT


def test_architect_analysis_reads_its_prompts_from_the_resource_root(
    tmp_path: Path, monkeypatch
):
    bundle = tmp_path / "bundle"
    bundle.mkdir()
    _freeze(monkeypatch, bundle)
    source = tmp_path / "template.docx"
    source.write_bytes(b"never opened: the prompt files are read first")

    with pytest.raises(FileNotFoundError, match="Missing required prompt file") as excinfo:
        phase1_pipeline.run_phase1(
            source,
            tmp_path / "profiles",
            "",
            classifier=lambda *args, **kwargs: None,
        )
    assert str(bundle.resolve() / "master_prompt.txt") in str(excinfo.value)


def test_profile_preparation_hands_the_analyzer_the_resource_root(
    tmp_path: Path, monkeypatch
):
    bundle = tmp_path / "bundle"
    bundle.mkdir()
    _freeze(monkeypatch, bundle)
    architect = tmp_path / "architect.docx"
    architect.write_bytes(b"hashed, never opened as a package here")
    seen: dict = {}

    def analyzer(**kwargs):
        seen.update(kwargs)
        raise RuntimeError("stop before analysis")

    with pytest.raises(RuntimeError, match="stop before analysis"):
        pipeline.prepare_template_profile(
            architect,
            tmp_path / "cache",
            "test-key",
            analyzer=analyzer,
        )
    assert seen["prompt_dir"] == bundle.resolve()


def test_target_prompt_fingerprints_follow_the_resource_root(tmp_path: Path, monkeypatch):
    checkout = pipeline._target_prompt_fingerprints()
    assert set(checkout) == {
        "phase2_master_prompt_sha256",
        "phase2_run_instruction_sha256",
    }
    _freeze(monkeypatch, tmp_path)
    assert pipeline._target_prompt_fingerprints() == {}
