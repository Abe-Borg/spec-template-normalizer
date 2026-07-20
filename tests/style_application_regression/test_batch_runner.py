from inspect import signature
from pathlib import Path

from spec_formatter.style_application.batch_runner import (
    BatchResult,
    _build_and_patch_output,
    _patch_header_footer_tokens_if_imported,
    process_single_file,
    run_batch_api,
    run_batch_concurrent,
)
from spec_formatter.style_application.core.csi_to_canadian import CSI_TO_CANADIAN


def test_new_role_specs_parameter_does_not_break_existing_positional_callers():
    assert list(signature(process_single_file).parameters)[-3:] == [
        "model", "role_specs", "conversion_mode"
    ]
    assert list(signature(run_batch_concurrent).parameters)[-4:] == [
        "max_workers", "on_file_complete", "role_specs", "conversion_mode"
    ]
    assert list(signature(run_batch_api).parameters)[-7:] == [
        "max_workers", "poll_interval", "on_file_complete", "on_batch_poll", "model",
        "role_specs", "conversion_mode"
    ]


def test_run_batch_concurrent_sorts_results_and_calls_callback(monkeypatch):
    completed = []

    def fake_process_single_file(
        docx_path,
        arch_registry,
        env_registry,
        arch_styles_xml,
        available_roles,
        api_key,
        output_dir,
        model="claude-sonnet-5",
    ):
        return BatchResult(
            filename=docx_path.name,
            success=True,
            output_path=output_dir / f"{docx_path.stem}_PHASE2_FORMATTED.docx",
            log=[f"Processed {docx_path.name}"],
            error=None,
            duration_seconds=0.1,
        )

    monkeypatch.setattr("spec_formatter.style_application.batch_runner.process_single_file", fake_process_single_file)

    docx_paths = [Path("b.docx"), Path("a.docx"), Path("c.docx")]

    results = run_batch_concurrent(
        docx_paths=docx_paths,
        arch_registry={},
        env_registry={},
        arch_styles_xml="",
        available_roles=[],
        api_key="k",
        output_dir=Path("out"),
        max_workers=2,
        on_file_complete=lambda result: completed.append(result.filename),
    )

    assert sorted(completed) == ["a.docx", "b.docx", "c.docx"]
    assert [item.filename for item in results] == ["a.docx", "b.docx", "c.docx"]


def test_direct_batch_canadian_output_name_does_not_collide(monkeypatch, tmp_path):
    source = tmp_path / "source.docx"
    source.write_bytes(b"source")
    extract = tmp_path / "extract"
    (extract / "word").mkdir(parents=True)
    (extract / "word" / "document.xml").write_bytes(b"document")
    (extract / "word" / "styles.xml").write_bytes(b"styles")

    def fake_patch_docx(**kwargs):
        Path(kwargs["out_docx"]).write_bytes(b"output")

    monkeypatch.setattr(
        "spec_formatter.style_application.batch_runner.patch_docx",
        fake_patch_docx,
    )
    monkeypatch.setattr(
        "spec_formatter.style_application.batch_runner.validate_docx_package",
        lambda *_args, **_kwargs: None,
    )
    monkeypatch.setattr(
        "spec_formatter.style_application.batch_runner.verify_phase2_invariants",
        lambda *_args, **_kwargs: None,
    )

    output = _build_and_patch_output(
        source,
        extract,
        {},
        tmp_path / "out",
        conversion_mode=CSI_TO_CANADIAN,
    )

    assert output.name == "source_CANADIAN_FORMATTED.docx"


def test_file_key_yields_batch_api_safe_custom_ids():
    # Batch API custom_ids must match [a-zA-Z0-9_-]{1,64}; CSI spec filenames
    # are long and dotted, so the stem must be sanitized and bounded.
    import re as _re

    from spec_formatter.style_application.batch_runner import _build_file_key

    long_dotted = Path("/specs/23 05 13 Common Motor Requirements for HVAC Equipment v2.1 FINAL.docx")
    custom_id = f"{_build_file_key(long_dotted)}__chunk12"
    assert _re.fullmatch(r"[A-Za-z0-9_-]{1,64}", custom_id)

    # Truncation must not collapse distinct paths with identical long stems.
    same_stem = "x" * 80 + ".docx"
    assert _build_file_key(Path("/a") / same_stem) != _build_file_key(Path("/b") / same_stem)

def test_target_header_tokens_are_not_patched_without_imported_architect_parts(
    monkeypatch, tmp_path
):
    calls = []
    monkeypatch.setattr(
        "spec_formatter.style_application.batch_runner.patch_header_footer_tokens",
        lambda *args: calls.append(args),
    )
    log = []

    changed = _patch_header_footer_tokens_if_imported(
        tmp_path,
        {"header_footer_import": {"part_names": set()}},
        {"SectionID": "SECTION 01 00 00"},
        {"SectionID": "SECTION 23 00 00"},
        log,
    )

    assert changed is False
    assert calls == []
    assert any("preserved target tokens unchanged" in line for line in log)


def test_imported_architect_header_tokens_are_patched(monkeypatch, tmp_path):
    calls = []
    monkeypatch.setattr(
        "spec_formatter.style_application.batch_runner.patch_header_footer_tokens",
        lambda *args, **kwargs: calls.append((args, kwargs)),
    )

    changed = _patch_header_footer_tokens_if_imported(
        tmp_path,
        {"header_footer_import": {"part_names": {"word/header1.xml"}}},
        {"SectionID": "SECTION 01 00 00"},
        {"SectionID": "SECTION 23 00 00"},
        [],
    )

    assert changed is True
    assert len(calls) == 1
    assert calls[0][1]["part_names"] == ["word/header1.xml"]
