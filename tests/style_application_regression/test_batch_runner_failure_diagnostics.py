import json
from pathlib import Path
from types import SimpleNamespace

import pytest

from spec_formatter.style_application import batch_runner
from spec_formatter.style_application.batch_runner import (
    ApplicationFailureDiagnostics,
    ApplicationStageError,
    PreparedFile,
)
from spec_formatter.style_application.core.csi_to_canadian import (
    CSI_TO_CANADIAN,
    CanadianConversionReport,
    ConversionIssue,
    MarkerEdit,
)


SECRET_TEXT = "CONFIDENTIAL TARGET PARAGRAPH"


def _seed_extract(tmp_path: Path) -> Path:
    extract_dir = tmp_path / "extract"
    word_dir = extract_dir / "word"
    word_dir.mkdir(parents=True)
    (word_dir / "styles.xml").write_text(
        '<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"/>',
        encoding="utf-8",
    )
    (word_dir / "document.xml").write_text(
        '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
        f"<w:body><w:p><w:r><w:t>{SECRET_TEXT}</w:t></w:r></w:p></w:body>"
        "</w:document>",
        encoding="utf-8",
    )
    return extract_dir


def _bundle_and_classifications() -> tuple[dict, dict]:
    bundle = {
        "paragraphs": [{"paragraph_index": 0, "text": SECRET_TEXT}],
        "deterministic_classifications": [],
        "deterministic_ignored_paragraphs": [],
        "filter_report": {
            "paragraphs_out_of_scope": [
                {
                    "paragraph_index": 9,
                    "reason": "table",
                    "original_text_preview": SECRET_TEXT,
                }
            ]
        },
    }
    classifications = {
        "classifications": [
            {
                "paragraph_index": 0,
                "csi_role": "PARAGRAPH",
                "original_text_preview": SECRET_TEXT,
            }
        ],
        "ignored_paragraphs": [],
    }
    return bundle, classifications


def _conversion_report() -> CanadianConversionReport:
    return CanadianConversionReport(
        paragraphs_examined=1,
        paragraphs_converted=1,
        literal_markers_removed=1,
        automatic_numbering_retargeted=0,
        unnumbered_paragraphs_numbered=0,
        edits=(
            MarkerEdit(
                paragraph_index=0,
                role="PARAGRAPH",
                source_kind="literal",
                target_kind="automatic",
                source_marker=SECRET_TEXT,
                target_marker="1.1",
            ),
        ),
        warnings=(
            ConversionIssue(
                paragraph_index=0,
                code="marker_warning",
                message=SECRET_TEXT,
                text_preview=SECRET_TEXT,
            ),
        ),
    )


def _application_kwargs(tmp_path: Path, extract_dir: Path) -> dict:
    bundle, classifications = _bundle_and_classifications()
    source = tmp_path / "source.docx"
    source.write_bytes(b"source package")
    return {
        "docx_path": source,
        "extract_dir": extract_dir,
        "bundle": bundle,
        "classifications": classifications,
        "arch_registry": {"PARAGRAPH": "Body"},
        "env_registry": {},
        "arch_styles_xml": "<w:styles/>",
        "output_dir": tmp_path / "output",
        "log": [],
        "source_tokens": None,
        "arch_root": None,
        "role_specs": None,
        "conversion_mode": CSI_TO_CANADIAN,
    }


def _stub_successful_conversion(monkeypatch: pytest.MonkeyPatch, report) -> None:
    monkeypatch.setattr(batch_runner, "extract_target_tokens", lambda *_args: {})
    monkeypatch.setattr(
        batch_runner,
        "apply_csi_to_canadian",
        lambda *_args, **_kwargs: report,
    )
    monkeypatch.setattr(
        batch_runner,
        "classifications_for_canadian_application",
        lambda classifications, _report: classifications,
    )


def test_environment_failure_exposes_text_free_conversion_and_classification_checkpoint(
    monkeypatch: pytest.MonkeyPatch,
    tmp_path: Path,
) -> None:
    extract_dir = _seed_extract(tmp_path)
    report = _conversion_report()
    _stub_successful_conversion(monkeypatch, report)

    def fail_environment(**_kwargs):
        raise RuntimeError("environment unavailable")

    monkeypatch.setattr(batch_runner, "apply_environment_to_target", fail_environment)

    with pytest.raises(ApplicationStageError) as raised:
        batch_runner._apply_classified_target(
            **_application_kwargs(tmp_path, extract_dir)
        )

    error = raised.value
    assert error.stage == "environment_application"
    assert error.audit_summary == {
        "styled": 1,
        "ignored": 0,
        "out_of_scope": 1,
        "unresolved": 0,
    }
    assert error.audit["classifications"] == [
        {"paragraph_index": 0, "csi_role": "PARAGRAPH"}
    ]
    assert error.audit["out_of_scope"] == [
        {"paragraph_index": 9, "reason": "table"}
    ]
    assert error.numbering_checks == {}
    assert error.conversion_report is not None
    assert error.conversion_report.paragraphs_converted == 1
    assert error.conversion_report.edits[0].source_marker is None
    assert error.conversion_report.warnings[0].text_preview == ""
    assert SECRET_TEXT not in json.dumps(error.diagnostics.as_dict())
    assert not (tmp_path / "output").exists()


def test_batch_result_preserves_late_numbering_checkpoint_without_publishing_docx(
    monkeypatch: pytest.MonkeyPatch,
    tmp_path: Path,
) -> None:
    extract_dir = _seed_extract(tmp_path)
    report = _conversion_report()
    kwargs = _application_kwargs(tmp_path, extract_dir)
    bundle = kwargs["bundle"]
    classifications = kwargs["classifications"]
    _stub_successful_conversion(monkeypatch, report)

    monkeypatch.setattr(
        batch_runner,
        "apply_environment_to_target",
        lambda **_kwargs: {"header_footer_import": {}},
    )
    monkeypatch.setattr(batch_runner, "HAS_NUMBERING_IMPORTER", False)
    monkeypatch.setattr(
        batch_runner,
        "_check_numbering_module_needed",
        lambda *_args: None,
    )
    monkeypatch.setattr(
        batch_runner,
        "remap_header_footer_numids",
        lambda *_args, **_kwargs: None,
    )
    monkeypatch.setattr(
        batch_runner,
        "import_arch_styles_into_target",
        lambda **_kwargs: SimpleNamespace(
            body_style_id_map={"Body": "Body"},
            style_id_map={},
        ),
    )
    monkeypatch.setattr(batch_runner, "snapshot_stability", lambda *_args: object())
    monkeypatch.setattr(
        batch_runner,
        "apply_phase2_classifications",
        lambda **_kwargs: SimpleNamespace(
            requested=1,
            modified=1,
            skipped_sectpr=[],
            allowed_rpr_properties_by_paragraph={},
            numbering_checks={
                "policy": CSI_TO_CANADIAN,
                "paragraphs_checked": 1,
                "body_text_preserved": True,
                "detail": SECRET_TEXT,
            },
        ),
    )
    monkeypatch.setattr(
        batch_runner,
        "verify_stability",
        lambda *_args, **_kwargs: None,
    )

    def fail_output(*_args, **_kwargs):
        raise RuntimeError("package validation failed")

    monkeypatch.setattr(batch_runner, "_build_and_patch_output", fail_output)

    prepared = PreparedFile(
        file_key="source",
        docx_path=kwargs["docx_path"],
        extract_dir=extract_dir,
        bundle=bundle,
        prep_log=["prepared"],
    )
    result = batch_runner._apply_batch_result(
        prepared,
        classifications,
        kwargs["arch_registry"],
        kwargs["env_registry"],
        kwargs["arch_styles_xml"],
        kwargs["output_dir"],
        conversion_mode=CSI_TO_CANADIAN,
    )

    assert result.success is False
    assert result.output_path is None
    assert result.stage == "output_publication"
    assert result.conversion_report is not None
    assert result.conversion_report.paragraphs_converted == 1
    assert result.audit_summary["styled"] == 1
    assert result.numbering_checks == {
        "policy": CSI_TO_CANADIAN,
        "paragraphs_checked": 1,
        "body_text_preserved": True,
    }
    assert SECRET_TEXT not in json.dumps(result.audit)
    assert SECRET_TEXT not in json.dumps(result.numbering_checks)
    assert not kwargs["output_dir"].exists()


def test_process_single_file_translates_staged_failure_diagnostics(
    monkeypatch: pytest.MonkeyPatch,
    tmp_path: Path,
) -> None:
    extract_dir = _seed_extract(tmp_path)
    bundle, classifications = _bundle_and_classifications()
    source = tmp_path / "source.docx"
    source.write_bytes(b"source package")
    report = CanadianConversionReport(
        paragraphs_examined=1,
        paragraphs_converted=1,
        literal_markers_removed=1,
        automatic_numbering_retargeted=0,
        unnumbered_paragraphs_numbered=0,
        edits=(),
        warnings=(),
    )
    diagnostics = ApplicationFailureDiagnostics(
        stage="environment_application",
        conversion_report=report,
        audit_summary={
            "styled": 1,
            "ignored": 0,
            "out_of_scope": 1,
            "unresolved": 0,
        },
        audit={
            "schema_version": 1,
            "summary": {"styled": 1},
            "classifications": [
                {"paragraph_index": 0, "csi_role": "PARAGRAPH"}
            ],
        },
        numbering_checks={"paragraphs_checked": 1},
    )

    class FakeDecomposer:
        def __init__(self, _path: str) -> None:
            pass

        def extract(self, *, output_dir: Path) -> Path:
            del output_dir
            return extract_dir

    monkeypatch.setattr(batch_runner, "DocxDecomposer", FakeDecomposer)
    monkeypatch.setattr(
        batch_runner,
        "build_phase2_slim_bundle",
        lambda *_args, **_kwargs: bundle,
    )
    monkeypatch.setattr(
        batch_runner,
        "classify_target_document",
        lambda **_kwargs: classifications,
    )

    def fail_application(**_kwargs):
        raise ApplicationStageError(
            "environment unavailable",
            diagnostics=diagnostics,
        )

    monkeypatch.setattr(batch_runner, "_apply_classified_target", fail_application)

    result = batch_runner.process_single_file(
        docx_path=source,
        arch_registry={"PARAGRAPH": "Body"},
        env_registry={},
        arch_styles_xml="<w:styles/>",
        available_roles=["PARAGRAPH"],
        api_key="offline-test-key",
        output_dir=tmp_path / "output",
        conversion_mode=CSI_TO_CANADIAN,
    )

    assert result.success is False
    assert result.output_path is None
    assert result.stage == "environment_application"
    assert result.conversion_report is report
    assert result.audit_summary["styled"] == 1
    assert result.audit["classifications"][0]["paragraph_index"] == 0
    assert result.numbering_checks == {"paragraphs_checked": 1}


def test_missing_target_api_key_reports_classification_preflight_stage(
    monkeypatch: pytest.MonkeyPatch,
    tmp_path: Path,
) -> None:
    extract_dir = _seed_extract(tmp_path)
    source = tmp_path / "source.docx"
    source.write_bytes(b"source package")

    class FakeDecomposer:
        def __init__(self, _path: str) -> None:
            pass

        def extract(self, *, output_dir: Path) -> Path:
            del output_dir
            return extract_dir

    monkeypatch.setattr(batch_runner, "DocxDecomposer", FakeDecomposer)
    monkeypatch.setattr(
        batch_runner,
        "build_phase2_slim_bundle",
        lambda *_args, **_kwargs: {
            "paragraphs": [{"paragraph_index": 0}],
            "deterministic_classifications": [],
        },
    )
    monkeypatch.setattr(
        batch_runner,
        "classify_target_document",
        lambda **_kwargs: pytest.fail("classification must not run without a key"),
    )

    result = batch_runner.process_single_file(
        docx_path=source,
        arch_registry={"PARAGRAPH": "Body"},
        env_registry={},
        arch_styles_xml="<w:styles/>",
        available_roles=["PARAGRAPH"],
        api_key="",
        output_dir=tmp_path / "output",
    )

    assert result.success is False
    assert result.stage == "classification_preflight"
    assert result.error == (
        "Anthropic API key is required when unresolved paragraphs exist."
    )
