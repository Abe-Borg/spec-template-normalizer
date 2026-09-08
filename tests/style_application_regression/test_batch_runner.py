import pytest
from inspect import signature
from pathlib import Path

from spec_formatter.style_application.batch_runner import (
    _build_and_patch_output,
    _patch_header_footer_tokens_if_imported,
    process_single_file,
)
from spec_formatter.style_application.core.csi_to_canadian import CSI_TO_CANADIAN


def test_new_role_specs_parameter_does_not_break_existing_positional_callers():
    assert list(signature(process_single_file).parameters)[-3:] == [
        "model", "role_specs", "conversion_mode"
    ]


def test_unreachable_batch_entry_points_are_gone():
    from spec_formatter.style_application import batch_runner

    assert not hasattr(batch_runner, "run_batch_concurrent")
    assert not hasattr(batch_runner, "run_batch_api")
    with pytest.raises(ModuleNotFoundError):
        import importlib

        importlib.import_module("spec_formatter.style_application.core.batch_classifier")


def test_direct_canadian_output_name_does_not_collide(monkeypatch, tmp_path):
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


def test_direct_format_only_output_name_matches_the_pipeline_suffix(monkeypatch, tmp_path):
    # The engine used to stage format_only output as _PHASE2_FORMATTED.docx
    # while the pipeline planned _FORMATTED.docx; both now read the suffix
    # from the one ApplicationPolicy.
    from spec_formatter.pipeline import _plan_output_paths
    from spec_formatter.style_application.core.application_policy import (
        application_policy_for_mode,
    )

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

    output = _build_and_patch_output(source, extract, {}, tmp_path / "out")

    assert output.name == "source_FORMATTED.docx"
    assert output.name.endswith(application_policy_for_mode("format_only").output_suffix)
    assert _plan_output_paths([source], tmp_path / "planned")[source].name == output.name


def test_file_key_is_short_safe_and_path_unique():
    # PreparedFile keys must match [a-zA-Z0-9_-]{1,64}; CSI spec filenames are
    # long and dotted, so the stem must be sanitized and bounded.
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


def test_imported_architect_tokens_can_be_inferred_when_source_map_is_empty(
    monkeypatch,
    tmp_path,
):
    calls = []
    monkeypatch.setattr(
        "spec_formatter.style_application.batch_runner.patch_header_footer_tokens",
        lambda *args, **kwargs: calls.append((args, kwargs)),
    )

    changed = _patch_header_footer_tokens_if_imported(
        tmp_path,
        {
            "header_footer_import": {
                "part_names": {"word/header1.xml", "word/footer1.xml"}
            }
        },
        {},
        {
            "SectionID": "SECTION 012900",
            "SectionTitle": "PAYMENT PROCEDURES",
        },
        [],
    )

    assert changed is True
    assert len(calls) == 1
    assert calls[0][0][1] == {}
    assert calls[0][1]["part_names"] == [
        "word/footer1.xml",
        "word/header1.xml",
    ]


def test_runner_wrapper_fails_closed_when_target_cannot_fill_an_imported_slot(tmp_path):
    word_dir = tmp_path / "word"
    word_dir.mkdir(parents=True)
    footer = word_dir / "footer1.xml"
    footer.write_text(
        '<w:ftr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
        "<w:p><w:r><w:t>SECTION 01 00 00</w:t></w:r></w:p></w:ftr>",
        encoding="utf-8",
    )
    original = footer.read_bytes()
    log: list[str] = []

    with pytest.raises(ValueError, match="requires a recognisable target SectionID"):
        _patch_header_footer_tokens_if_imported(
            tmp_path,
            {"header_footer_import": {"part_names": {"word/footer1.xml"}}},
            {"SectionID": "SECTION 01 00 00"},
            {},
            log,
        )

    assert footer.read_bytes() == original
    assert not any("preserved target tokens unchanged" in line for line in log)

