from __future__ import annotations

import zipfile
from pathlib import Path

from spec_formatter.pipeline import FORMAT_ONLY, format_specifications
from spec_formatter.style_application.core.classification import (
    _build_numbering_catalog,
    _effective_numbering_semantics,
)
from spec_formatter.style_application.core import classification as classification_module
from spec_formatter.style_application.core.ooxml_text import decode_xml_bytes
from spec_formatter.style_application.core.xml_helpers import (
    iter_paragraph_xml_blocks,
    paragraph_text_from_block,
)
from spec_formatter.style_application.phase2_invariants import (
    _numbering_definition_signatures,
    validate_docx_package,
)
from tests.fixtures.sanitized_format_only_corpus import (
    CRITICAL_MARKER_INDICES,
    NUMBERED_LEVEL_BY_INDEX,
    ROLE_BY_LEVEL,
    sanitized_template_classifier,
    write_sanitized_format_only_pair,
)


def _read_word_parts(path: Path) -> tuple[str, str, str]:
    with zipfile.ZipFile(path) as package:
        document = decode_xml_bytes(
            package.read("word/document.xml"),
            part_name="word/document.xml",
        )
        styles = decode_xml_bytes(
            package.read("word/styles.xml"),
            part_name="word/styles.xml",
        )
        numbering = decode_xml_bytes(
            package.read("word/numbering.xml"),
            part_name="word/numbering.xml",
        )
    return document, styles, numbering


def _paragraphs(document_xml: str) -> list[str]:
    return [block for _start, _end, block in iter_paragraph_xml_blocks(document_xml)]


def _numbering_semantics(
    paragraphs: list[str],
    styles_xml: str,
    numbering_xml: str,
) -> list[dict | None]:
    catalog = _build_numbering_catalog(numbering_xml)
    return [
        _effective_numbering_semantics(paragraph, styles_xml, catalog)
        for paragraph in paragraphs
    ]


def test_sanitized_154_paragraph_format_only_regression(tmp_path: Path) -> None:
    architect = tmp_path / "sanitized-architect.docx"
    target = tmp_path / "sanitized-payment-procedures.docx"
    write_sanitized_format_only_pair(architect, target)

    source_document, source_styles, source_numbering = _read_word_parts(target)
    source_paragraphs = _paragraphs(source_document)
    source_text = [paragraph_text_from_block(item) for item in source_paragraphs]
    source_semantics = _numbering_semantics(
        source_paragraphs,
        source_styles,
        source_numbering,
    )
    source_ignored_indices = set(range(154)) - set(NUMBERED_LEVEL_BY_INDEX) - {0}

    assert len(source_paragraphs) == 154
    assert sum(item is not None for item in source_semantics) == 121
    assert source_text[1] == "GENERAL"

    run = format_specifications(
        architect_template=architect,
        target_specs=[target],
        output_dir=tmp_path / "formatted",
        cache_dir=tmp_path / "template-cache",
        api_key="",
        max_workers=1,
        conversion_mode=FORMAT_ONLY,
        template_model="sanitized-format-only-corpus",
        template_classifier=sanitized_template_classifier,
    )

    assert run.success, "\n".join(run.targets[0].log)
    result = run.targets[0]
    assert result.output_path is not None
    assert result.audit_summary == {
        "styled": 122,
        "ignored": 32,
        "out_of_scope": 0,
        "unresolved": 0,
    }
    assert result.numbering_checks == {
        "policy": FORMAT_ONLY,
        "paragraphs_checked": 154,
        "automatic_numbered_before": 121,
        "effective_numbering_preserved": True,
        "body_text_preserved": True,
    }
    assert any(
        "All paragraphs classified deterministically" in line
        for line in result.log
    )

    classifications = {
        item["paragraph_index"]: item["csi_role"]
        for item in result.audit["classifications"]
    }
    ignored_indices = {
        item["paragraph_index"] for item in result.audit["ignored_paragraphs"]
    }
    assert ignored_indices == source_ignored_indices
    assert classifications[0] == "SectionID"
    assert classifications[1] == "PART"  # GENERAL follows automatic numbering.
    assert {
        index: classifications[index] for index in CRITICAL_MARKER_INDICES
    } == {
        index: ROLE_BY_LEVEL[NUMBERED_LEVEL_BY_INDEX[index]]
        for index in CRITICAL_MARKER_INDICES
    }

    validate_docx_package(result.output_path)
    output_document, output_styles, output_numbering = _read_word_parts(
        result.output_path
    )
    output_paragraphs = _paragraphs(output_document)
    output_semantics = _numbering_semantics(
        output_paragraphs,
        output_styles,
        output_numbering,
    )

    assert len(output_paragraphs) == 154
    assert [paragraph_text_from_block(item) for item in output_paragraphs] == source_text
    assert output_semantics == source_semantics
    assert _numbering_definition_signatures(output_numbering) == (
        _numbering_definition_signatures(source_numbering)
    )
    assert all(
        source_paragraphs[index] == output_paragraphs[index]
        for index in source_ignored_indices
    )

    # The architect PARAGRAPH style supplies bold. The formatter may therefore
    # remove only the conflicting direct bold override from styled level-4
    # paragraphs; their unrelated direct color must survive.
    style_replaced_indices = {
        index for index, level in NUMBERED_LEVEL_BY_INDEX.items() if level == 4
    }
    assert all(
        '<w:b w:val="0"/>' in source_paragraphs[index]
        for index in style_replaced_indices
    )
    assert all(
        '<w:b w:val="0"/>' not in output_paragraphs[index]
        and '<w:color w:val="556677"/>' in output_paragraphs[index]
        for index in style_replaced_indices
    )
    assert "<w:b" in output_styles

    # The target list stays authoritative while the selected architect shell
    # replaces the target page size and header/footer relationships.
    assert '<w:pgSz w:w="10000" w:h="15000"/>' in output_document
    assert '<w:pgSz w:w="12000" w:h="16000"/>' not in output_document
    with zipfile.ZipFile(result.output_path) as package:
        names = set(package.namelist())
    assert "word/header9.xml" not in names
    assert {"word/header1.xml", "word/header2.xml", "word/header3.xml"} <= names


def test_final_verifier_rejects_uncontracted_run_property_loss(
    tmp_path: Path,
    monkeypatch,
) -> None:
    architect = tmp_path / "sanitized-architect.docx"
    target = tmp_path / "sanitized-target.docx"
    write_sanitized_format_only_pair(architect, target)

    real_strip = classification_module.strip_direct_run_properties

    def overbroad_strip(paragraph_xml, properties):
        # Fault injection: make both application and its immediate normalized
        # contract believe color is replaceable. Final verification receives
        # the real style-derived manifest and must still reject the extra loss.
        return real_strip(paragraph_xml, set(properties) | {"color"})

    monkeypatch.setattr(
        classification_module,
        "strip_direct_run_properties",
        overbroad_strip,
    )

    run = format_specifications(
        architect_template=architect,
        target_specs=[target],
        output_dir=tmp_path / "formatted",
        cache_dir=tmp_path / "template-cache",
        api_key="",
        max_workers=1,
        conversion_mode=FORMAT_ONLY,
        template_model="sanitized-format-only-corpus",
        template_classifier=sanitized_template_classifier,
    )

    assert not run.success
    assert run.failed == 1
    result = run.targets[0]
    assert result.output_path is None
    assert result.error is not None
    assert "non-font run formatting was lost" in result.error
    assert "paragraph 3" in result.error
