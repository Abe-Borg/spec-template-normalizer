import json
from pathlib import Path
from types import SimpleNamespace

import pytest

from spec_formatter.style_application import batch_runner
from spec_formatter.style_application.batch_runner import (
    ApplicationFailureDiagnostics,
    ApplicationStageError,
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
            # Mirror every ``ApplyReport`` attribute the shared application
            # path reads while recording the ``apply_classifications`` phase.
            # A missing attribute aborts the target inside that phase with the
            # wrong stage, so the stub must stay complete.
            requested=1,
            modified=1,
            invalid_indices=[],
            skipped_sectpr=[],
            unmapped_roles=[],
            missing_style_ids=set(),
            stripped_direct_ppr=0,
            preserved_direct_ppr=0,
            preserved_automatic_numbering=0,
            suppressed_architect_numbering=0,
            stripped_run_fonts=0,
            ignored=0,
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

    result = batch_runner.process_single_file(
        docx_path=kwargs["docx_path"],
        arch_registry=kwargs["arch_registry"],
        env_registry=kwargs["env_registry"],
        arch_styles_xml=kwargs["arch_styles_xml"],
        available_roles=["PARAGRAPH"],
        api_key="offline-test-key",
        output_dir=kwargs["output_dir"],
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


def _seed_verification_kwargs(tmp_path: Path, classifications: dict) -> dict:
    extract_dir = _seed_extract(tmp_path)
    kwargs = _application_kwargs(tmp_path, extract_dir)
    kwargs["classifications"] = classifications
    return kwargs


@pytest.mark.parametrize(
    ("classifications", "match"),
    [
        (
            {"classifications": [], "ignored_paragraphs": []},
            "missing coverage",
        ),
        (
            {
                "classifications": [
                    {"paragraph_index": 0, "csi_role": "PARAGRAPH"},
                    {"paragraph_index": 0, "csi_role": "PARAGRAPH"},
                ],
                "ignored_paragraphs": [],
            },
            "[Dd]uplicate",
        ),
        (
            {
                "classifications": [{"paragraph_index": 0, "csi_role": "PARAGRAPH"}],
                "ignored_paragraphs": [{"paragraph_index": 0, "reason": "editorial"}],
            },
            "both|overlap|[Dd]uplicate",
        ),
        (
            {
                "classifications": [
                    {"paragraph_index": 0, "csi_role": "PARAGRAPH"},
                    {"paragraph_index": 7, "csi_role": "PARAGRAPH"},
                ],
                "ignored_paragraphs": [],
            },
            "not classifiable",
        ),
    ],
)
def test_shared_application_path_reverifies_disposition_coverage(
    monkeypatch: pytest.MonkeyPatch,
    tmp_path: Path,
    classifications: dict,
    match: str,
) -> None:
    kwargs = _seed_verification_kwargs(tmp_path, classifications)
    monkeypatch.setattr(
        batch_runner,
        "apply_environment_to_target",
        lambda **_kwargs: pytest.fail("application must not start"),
    )

    with pytest.raises(ApplicationStageError, match=match) as raised:
        batch_runner._apply_classified_target(**kwargs)

    assert raised.value.stage == "disposition_verification"
    assert not (tmp_path / "output").exists()


def test_shared_application_path_rejects_deterministic_override(
    monkeypatch: pytest.MonkeyPatch,
    tmp_path: Path,
) -> None:
    extract_dir = _seed_extract(tmp_path)
    kwargs = _application_kwargs(tmp_path, extract_dir)
    kwargs["bundle"] = {
        "paragraphs": [],
        "deterministic_classifications": [
            {"paragraph_index": 0, "csi_role": "PARAGRAPH"}
        ],
        "deterministic_ignored_paragraphs": [],
        "filter_report": {"paragraphs_out_of_scope": []},
    }
    kwargs["classifications"] = {
        "classifications": [{"paragraph_index": 0, "csi_role": "ARTICLE"}],
        "ignored_paragraphs": [],
    }
    kwargs["arch_registry"] = {"PARAGRAPH": "Body", "ARTICLE": "Body"}
    monkeypatch.setattr(
        batch_runner,
        "apply_environment_to_target",
        lambda **_kwargs: pytest.fail("application must not start"),
    )

    with pytest.raises(ApplicationStageError, match="override") as raised:
        batch_runner._apply_classified_target(**kwargs)

    assert raised.value.stage == "disposition_verification"


def test_classification_audit_does_not_clamp_an_overfull_payload() -> None:
    bundle, classifications = _bundle_and_classifications()
    classifications["classifications"].append(
        {"paragraph_index": 0, "csi_role": "PARAGRAPH"}
    )

    summary, _audit = batch_runner._classification_audit(bundle, classifications)

    assert summary["unresolved"] == -1


def test_classifier_usage_becomes_classify_phase_diagnostics_not_payload(
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
            "deterministic_ignored_paragraphs": [],
        },
    )
    monkeypatch.setattr(
        batch_runner,
        "classify_target_document",
        lambda **_kwargs: {
            "classifications": [{"paragraph_index": 0, "csi_role": "PARAGRAPH"}],
            "ignored_paragraphs": [],
            "notes": [],
            "usage": {
                "requests": 2,
                "input_tokens": 1200,
                "cache_read_input_tokens": 700,
                "cache_creation_input_tokens": 400,
                "bogus": "text that must never reach diagnostics",
            },
        },
    )
    seen_payloads = []

    def fake_apply(**kwargs):
        seen_payloads.append(kwargs["classifications"])
        return (
            tmp_path / "out.docx",
            None,
            {"styled": 1, "ignored": 0, "out_of_scope": 0, "unresolved": 0},
            {},
            {},
        )

    monkeypatch.setattr(batch_runner, "_apply_classified_target", fake_apply)

    result = batch_runner.process_single_file(
        docx_path=source,
        arch_registry={"PARAGRAPH": "Body"},
        env_registry={},
        arch_styles_xml="<w:styles/>",
        available_roles=["PARAGRAPH"],
        api_key="key",
        output_dir=tmp_path / "output",
    )

    assert result.success is True
    assert "usage" not in seen_payloads[0]
    classify_events = [e for e in result.diagnostics if e.get("event") == "classify"]
    assert classify_events, result.diagnostics
    fields = classify_events[-1]["fields"]
    assert fields["requests"] == 2
    assert fields["cache_read_input_tokens"] == 700
    assert fields["cache_creation_input_tokens"] == 400
    assert "bogus" not in fields

    # The phase event is per-phase detail; the field is what cost is read
    # from, because events are level-filtered and totals must not be. This
    # assertion was missing, so a success path that dropped the field
    # entirely still passed.
    assert result.usage["requests"] == 2
    assert result.usage["input_tokens"] == 1200
    assert result.usage["cache_read_input_tokens"] == 700



def test_successful_target_reports_its_usage_on_the_result(
    monkeypatch: pytest.MonkeyPatch,
    tmp_path: Path,
) -> None:
    """A deterministic-only target must report zero, not nothing.

    "We sent no requests" and "we could not tell you what we sent" are
    different answers and have to look different. The classifier returns an
    explicit zero snapshot for a target it resolved locally; if the success
    path drops it, ``record_usage`` sees an empty dict, ignores it, and
    ``run.json`` carries no target scope at all -- indistinguishable from a
    run whose counters never arrived.
    """

    from spec_formatter.llm_usage import UsageCollector

    extract_dir = _seed_extract(tmp_path)
    source = tmp_path / "source.docx"
    source.write_bytes(b"source package")
    zero = UsageCollector().snapshot()

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
            "paragraphs": [],
            "deterministic_classifications": [
                {"paragraph_index": 0, "csi_role": "PARAGRAPH"}
            ],
            "deterministic_ignored_paragraphs": [],
        },
    )
    monkeypatch.setattr(
        batch_runner,
        "classify_target_document",
        lambda **_kwargs: {
            "classifications": [{"paragraph_index": 0, "csi_role": "PARAGRAPH"}],
            "ignored_paragraphs": [],
            "notes": [],
            "usage": dict(zero),
        },
    )
    monkeypatch.setattr(
        batch_runner,
        "_apply_classified_target",
        lambda **_kwargs: (
            tmp_path / "out.docx",
            None,
            {"styled": 1, "ignored": 0, "out_of_scope": 0, "unresolved": 0},
            {},
            {},
        ),
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

    assert result.success is True
    assert result.usage == zero
    assert result.usage["requests_attempted"] == 0
    assert result.usage["usage_complete"] is True


def test_failed_classification_still_reports_what_it_spent(
    monkeypatch: pytest.MonkeyPatch,
    tmp_path: Path,
) -> None:
    """A target that fails mid-classification is not a free target.

    The counts are observed inside the classifier and would leave with the
    exception. Without the handoff the run reports a refused or exhausted
    target as costing nothing, which is worse than reporting nothing at all.
    """
    from spec_formatter.llm_usage import UsageCollector, attach_usage

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
            "deterministic_ignored_paragraphs": [],
        },
    )

    collector = UsageCollector()
    collector.record_attempt()
    collector.record_response(
        SimpleNamespace(
            stop_reason="refusal",
            usage=SimpleNamespace(
                input_tokens=1500,
                output_tokens=30,
                cache_read_input_tokens=0,
                cache_creation_input_tokens=0,
            ),
        )
    )
    collector.record_attempt()  # a second request whose usage never arrived

    def refuse(**_kwargs):
        raise attach_usage(RuntimeError(SECRET_TEXT), collector)

    monkeypatch.setattr(batch_runner, "classify_target_document", refuse)

    result = batch_runner.process_single_file(
        docx_path=source,
        arch_registry={"PARAGRAPH": "Body"},
        env_registry={},
        arch_styles_xml="<w:styles/>",
        available_roles=["PARAGRAPH"],
        api_key="key",
        output_dir=tmp_path / "output",
    )

    assert result.success is False
    classify_events = [e for e in result.diagnostics if e.get("event") == "classify"]
    assert classify_events, result.diagnostics
    fields = classify_events[-1]["fields"]
    assert fields["input_tokens"] == 1500
    assert fields["output_tokens"] == 30
    assert fields["requests_attempted"] == 2
    # The second request never reported counters, so the total is a lower
    # bound and says so rather than looking complete.
    assert fields["requests_with_unknown_usage"] == 1
    assert fields["usage_complete"] is False
    # The failure text is document-derived and must not ride along.
    assert SECRET_TEXT not in json.dumps(result.diagnostics)
