import pytest
from inspect import signature
from pathlib import Path

from spec_formatter.style_application.batch_runner import (
    _build_and_patch_output,
    _patch_header_footer_tokens_if_imported,
    _remap_imported_header_footer_style_ids,
    process_single_file,
)
from spec_formatter.style_application.core.csi_to_canadian import CSI_TO_CANADIAN

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"

# Where each header/footer style reference lives.
_REFERENCE_CONTEXTS = {
    "pStyle": ("<w:p><w:pPr>", "</w:pPr></w:p>"),
    "rStyle": ("<w:p><w:r><w:rPr>", "</w:rPr><w:t>x</w:t></w:r></w:p>"),
    "tblStyle": ("<w:tbl><w:tblPr>", "</w:tblPr></w:tbl>"),
}


def _write_header(tmp_path, body):
    part = tmp_path / "word" / "header1.xml"
    part.parent.mkdir(parents=True, exist_ok=True)
    part.write_text(f'<w:hdr xmlns:w="{W_NS}">{body}</w:hdr>', encoding="utf-8")
    return part


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


def test_imported_header_double_quoted_references_are_remapped_as_before(tmp_path):
    # The form Word writes. A reference in tracked history names an
    # architect style too, so it is pointed at the same clone.
    body = (
        '<w:p><w:pPr><w:pStyle w:val="Header"/>'
        '<w:pPrChange w:id="1" w:author="A"><w:pPr><w:pStyle w:val="Header"/>'
        '</w:pPr></w:pPrChange></w:pPr>'
        '<w:r><w:rPr><w:rStyle w:val="HeaderChar"/></w:rPr><w:t>Header</w:t></w:r></w:p>'
        '<w:p><w:pPr><w:pStyle w:val="Unmapped"/></w:pPr></w:p>'
    )
    part = _write_header(tmp_path, body)
    log = []

    _remap_imported_header_footer_style_ids(
        tmp_path,
        ["word/header1.xml"],
        {"Header": "SF_header", "HeaderChar": "SF_header_char", "Normal": "Normal"},
        log,
    )

    assert part.read_text(encoding="utf-8") == (
        f'<w:hdr xmlns:w="{W_NS}">'
        '<w:p><w:pPr><w:pStyle w:val="SF_header"/>'
        '<w:pPrChange w:id="1" w:author="A"><w:pPr><w:pStyle w:val="SF_header"/>'
        '</w:pPr></w:pPrChange></w:pPr>'
        '<w:r><w:rPr><w:rStyle w:val="SF_header_char"/></w:rPr><w:t>Header</w:t></w:r></w:p>'
        '<w:p><w:pPr><w:pStyle w:val="Unmapped"/></w:pPr></w:p>'
        "</w:hdr>"
    )
    assert log == ["Remapped collision-safe style IDs in 1 imported header/footer parts"]


@pytest.mark.parametrize("tag", sorted(_REFERENCE_CONTEXTS))
@pytest.mark.parametrize(
    "reference",
    ["w:val='Header'", 'w:val = "Header"', "w:val = 'Header'"],
    ids=["single_quoted", "spaced_equals", "single_quoted_spaced"],
)
def test_imported_header_reference_reaches_its_clone_whatever_its_quoting(
    tmp_path, tag, reference
):
    # Legal XML that Word never writes. It used to be skipped, so the header
    # kept resolving to the target's own Header rather than the clone.
    opening, closing = _REFERENCE_CONTEXTS[tag]
    part = _write_header(tmp_path, f"{opening}<w:{tag} {reference}/>{closing}")
    log = []

    _remap_imported_header_footer_style_ids(
        tmp_path, ["word/header1.xml"], {"Header": "SF_header"}, log
    )

    # Written back in the double-quoted form every later reader matches.
    assert part.read_text(encoding="utf-8") == (
        f'<w:hdr xmlns:w="{W_NS}">{opening}<w:{tag} w:val="SF_header"/>{closing}</w:hdr>'
    )
    assert log == ["Remapped collision-safe style IDs in 1 imported header/footer parts"]


def test_imported_header_references_with_nothing_to_remap_are_left_as_written(tmp_path):
    # Empty, unmapped, and identity references keep their own form; none of
    # them stops the single-quoted reference after them from being remapped.
    kept = (
        "<w:p><w:pPr><w:pStyle w:val=''/></w:pPr></w:p>"
        '<w:p><w:pPr><w:pStyle w:val=""/></w:pPr></w:p>'
        "<w:p><w:pPr><w:pStyle w:val = 'TargetOwned' /></w:pPr></w:p>"
        "<w:p><w:pPr><w:pStyle w:val='Normal'/></w:pPr></w:p>"
    )
    part = _write_header(tmp_path, kept + "<w:p><w:pPr><w:pStyle w:val='Header'/></w:pPr></w:p>")

    _remap_imported_header_footer_style_ids(
        tmp_path,
        ["word/header1.xml"],
        {"": "SF_nothing", "Header": "SF_header", "Normal": "Normal"},
        [],
    )

    assert part.read_text(encoding="utf-8") == (
        f'<w:hdr xmlns:w="{W_NS}">{kept}'
        '<w:p><w:pPr><w:pStyle w:val="SF_header"/></w:pPr></w:p></w:hdr>'
    )


def test_imported_header_text_shaped_like_a_reference_is_left_alone(tmp_path):
    # CDATA is visible header text, and a comment or processing instruction
    # is no markup at all: rewriting any of them would change the document,
    # not point a reference at a clone.
    kept = (
        "<!-- <w:pStyle w:val='Header'/> -->"
        '<?pi <w:rStyle w:val="Header"/> ?>'
        "<w:p><w:r><w:t><![CDATA[<w:pStyle w:val='Header'/>]]></w:t></w:r></w:p>"
    )
    part = _write_header(tmp_path, kept + "<w:p><w:pPr><w:pStyle w:val='Header'/></w:pPr></w:p>")

    _remap_imported_header_footer_style_ids(
        tmp_path, ["word/header1.xml"], {"Header": "SF_header"}, []
    )

    assert part.read_text(encoding="utf-8") == (
        f'<w:hdr xmlns:w="{W_NS}">{kept}'
        '<w:p><w:pPr><w:pStyle w:val="SF_header"/></w:pPr></w:p></w:hdr>'
    )
