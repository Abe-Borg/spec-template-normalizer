"""Stable engine error codes and the closed stage sets."""

from __future__ import annotations

import re
from pathlib import Path

import pytest

from spec_formatter.style_application import batch_runner
from spec_formatter.style_application.batch_runner import ApplicationStageError
from spec_formatter.style_application.core.errors import (
    ENGINE_STAGES,
    ERROR_REMEDIATIONS,
    PIPELINE_STAGES,
    RUNNER_STAGES,
    EngineError,
    attach_engine_error,
    remediation_for,
)

REPO_ROOT = Path(__file__).resolve().parents[2]
W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"


def test_engine_error_keeps_detail_and_carries_a_fixed_remediation():
    error = EngineError("classification_coverage_incomplete", "missing coverage for [4]")

    assert isinstance(error, ValueError)
    assert str(error) == "missing coverage for [4]"
    assert error.code == error.safe_error_code == "classification_coverage_incomplete"
    assert error.safe_error_message == ERROR_REMEDIATIONS["classification_coverage_incomplete"]
    assert remediation_for("classification_coverage_incomplete") == error.safe_error_message
    assert remediation_for("no_such_code") is None
    with pytest.raises(KeyError):
        EngineError("no_such_code", "x")
    with pytest.raises(KeyError):
        attach_engine_error(RuntimeError("x"), "no_such_code")


def test_remediations_are_fixed_sentences_without_placeholders():
    assert 15 <= len(ERROR_REMEDIATIONS) <= 20
    for code, sentence in ERROR_REMEDIATIONS.items():
        assert re.fullmatch(r"[a-z][a-z0-9_]+", code)
        assert sentence.strip() and "{" not in sentence and "}" not in sentence
        assert sentence.endswith(".")


def test_every_recorded_stage_is_in_the_closed_sets():
    source = (REPO_ROOT / "spec_formatter" / "style_application" / "batch_runner.py").read_text(
        encoding="utf-8"
    )
    engine_literals = set(re.findall(r'checkpoint\.stage = "([a-z_]+)"', source))
    engine_literals |= set(re.findall(r'stage="([a-z_]+)"', source))
    runner_literals = set(re.findall(r'^\s+stage = "([a-z_]+)"', source, flags=re.M))
    allowed = set(ENGINE_STAGES) | set(RUNNER_STAGES)
    assert engine_literals <= allowed, engine_literals - allowed
    assert runner_literals <= allowed, runner_literals - allowed
    assert set(RUNNER_STAGES) <= runner_literals | set(ENGINE_STAGES)

    pipeline_source = (REPO_ROOT / "spec_formatter" / "pipeline.py").read_text(encoding="utf-8")
    pipeline_literals = set(re.findall(r'stage[=:] ?"([a-z_]+)"', pipeline_source))
    assert pipeline_literals <= set(PIPELINE_STAGES) | set(ENGINE_STAGES), pipeline_literals


def test_attach_engine_error_keeps_the_original_exception_type():
    error = attach_engine_error(ImportError("no numbering importer"), "numbering_importer_unavailable")

    assert isinstance(error, ImportError)
    assert error.safe_error_code == "numbering_importer_unavailable"
    assert error.safe_error_message == ERROR_REMEDIATIONS["numbering_importer_unavailable"]


def _seed_hf_failure(tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> dict:
    extract_dir = tmp_path / "extract"
    word_dir = extract_dir / "word"
    word_dir.mkdir(parents=True)
    (word_dir / "styles.xml").write_text(f'<w:styles xmlns:w="{W_NS}"/>', encoding="utf-8")
    (word_dir / "document.xml").write_text(
        f'<w:document xmlns:w="{W_NS}"><w:body><w:p><w:r><w:t>Scope of work.</w:t></w:r></w:p>'
        "</w:body></w:document>",
        encoding="utf-8",
    )
    (word_dir / "footer1.xml").write_text(
        f'<w:ftr xmlns:w="{W_NS}"><w:p><w:r><w:t>SECTION 01 00 00</w:t></w:r></w:p></w:ftr>',
        encoding="utf-8",
    )
    source = tmp_path / "source.docx"
    source.write_bytes(b"source package")
    monkeypatch.setattr(
        batch_runner,
        "apply_environment_to_target",
        lambda **_kwargs: {"header_footer_import": {"part_names": {"word/footer1.xml"}}},
    )
    return {
        "docx_path": source,
        "extract_dir": extract_dir,
        "bundle": {
            "paragraphs": [{"paragraph_index": 0, "text": "Scope of work."}],
            "deterministic_classifications": [],
            "deterministic_ignored_paragraphs": [],
            "filter_report": {"paragraphs_out_of_scope": []},
        },
        "classifications": {
            "classifications": [{"paragraph_index": 0, "csi_role": "PARAGRAPH"}],
            "ignored_paragraphs": [],
        },
        "arch_registry": {"PARAGRAPH": "Body"},
        "env_registry": {},
        "arch_styles_xml": "<w:styles/>",
        "output_dir": tmp_path / "output",
        "log": [],
        "source_tokens": {"SectionID": "SECTION 01 00 00"},
        "arch_root": None,
        "role_specs": None,
        "conversion_mode": "format_only",
    }


def test_header_footer_failure_reaches_the_runner_with_its_code_and_stage(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
):
    kwargs = _seed_hf_failure(tmp_path, monkeypatch)

    with pytest.raises(ApplicationStageError, match="requires a recognisable target SectionID") as raised:
        batch_runner._apply_classified_target(**kwargs)

    error = raised.value
    assert error.stage == "header_footer_token_patch"
    assert error.safe_error_code == "header_footer_target_section_id_required"
    assert error.safe_error_message == ERROR_REMEDIATIONS["header_footer_target_section_id_required"]
    assert "Scope of work" not in error.safe_error_message


def test_process_single_file_returns_the_engine_error_code(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
):
    kwargs = _seed_hf_failure(tmp_path, monkeypatch)

    class FakeDecomposer:
        def __init__(self, _path: str) -> None:
            pass

        def extract(self, *, output_dir: Path) -> Path:
            del output_dir
            return kwargs["extract_dir"]

    monkeypatch.setattr(batch_runner, "DocxDecomposer", FakeDecomposer)
    monkeypatch.setattr(batch_runner, "build_phase2_slim_bundle", lambda *_a, **_k: kwargs["bundle"])
    monkeypatch.setattr(
        batch_runner,
        "classify_target_document",
        lambda **_kwargs: dict(kwargs["classifications"], notes=[]),
    )

    result = batch_runner.process_single_file(
        docx_path=kwargs["docx_path"],
        arch_registry=kwargs["arch_registry"],
        env_registry={},
        arch_styles_xml="<w:styles/>",
        available_roles=["PARAGRAPH"],
        api_key="key",
        output_dir=tmp_path / "output",
        source_tokens=kwargs["source_tokens"],
    )

    assert result.success is False
    assert result.stage == "header_footer_token_patch"
    assert result.error_code == "header_footer_target_section_id_required"
    assert result.safe_error == ERROR_REMEDIATIONS["header_footer_target_section_id_required"]
    assert "requires a recognisable target SectionID" in (result.error or "")
