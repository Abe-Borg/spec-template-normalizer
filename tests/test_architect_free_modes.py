"""End-to-end coverage for the two conversion modes that take no architect.

These runs are the ones a user reaches for when a spec has to go to Canadian
CSC PageFormat, or come back from it, and no architect template is in the
picture. What they must prove is narrow and specific: the numbering changes,
and nothing else does.
"""

from __future__ import annotations

import json
import zipfile
from pathlib import Path

import pytest

from spec_formatter import builtin_scheme
from spec_formatter.pipeline import (
    CANADIAN_TO_CSI,
    CSI_TO_CANADIAN_STANDALONE,
    format_specifications,
)
from spec_formatter.style_application.phase2_invariants import validate_docx_package

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"

#: A short CSI section with typed markers, written as (role, marker, text).
#: The marker is spelled out rather than split off the line, because a marker
#: must occupy whole ``w:t`` nodes for the converter to map it back to runs --
#: which is exactly how Word stores a typed marker followed by a list tab.
#: Every line is shaped so the deterministic classifier resolves it locally,
#: which is what lets these tests run with no API key at all.
CSI_LINES: tuple[tuple[str, str, str], ...] = (
    ("SectionID", "", "SECTION 21 13 13"),
    ("SectionTitle", "", "WET-PIPE SPRINKLER SYSTEMS"),
    ("PART", "PART 1", "- GENERAL"),
    ("ARTICLE", "1.1", "SUMMARY"),
    ("PARAGRAPH", "A.", "Section includes wet-pipe sprinkler systems."),
    ("PARAGRAPH", "B.", "Related requirements are specified elsewhere."),
    ("SUBPARAGRAPH", "1.", "Hydraulic calculations per NFPA 13."),
    ("ARTICLE", "1.2", "REFERENCES"),
    ("PARAGRAPH", "A.", "NFPA 13 governs sprinkler system installation."),
    ("PART", "PART 2", "- PRODUCTS"),
    ("ARTICLE", "2.1", "SPRINKLERS"),
    ("PARAGRAPH", "A.", "Provide listed quick-response sprinklers."),
    ("END_OF_SECTION", "", "END OF SECTION 21 13 13"),
)

#: The target's own page geometry, deliberately unlike anything the built-in
#: scheme could supply, so a shell that leaked in would be obvious.
TARGET_PAGE_SIZE = '<w:pgSz w:w="12240" w:h="15840"/>'


def _paragraph(marker: str, text: str) -> str:
    if not marker:
        return f'<w:p><w:pPr/><w:r><w:t xml:space="preserve">{text}</w:t></w:r></w:p>'
    return (
        "<w:p><w:pPr/>"
        f'<w:r><w:t xml:space="preserve">{marker}</w:t></w:r>'
        "<w:r><w:tab/></w:r>"
        f'<w:r><w:t xml:space="preserve">{text}</w:t></w:r></w:p>'
    )


def _document_xml() -> str:
    body = "".join(_paragraph(marker, text) for _role, marker, text in CSI_LINES)
    return (
        '<?xml version="1.0" encoding="UTF-8"?>'
        f'<w:document xmlns:w="{W_NS}" xmlns:r="{R_NS}"><w:body>{body}'
        f"<w:sectPr>{TARGET_PAGE_SIZE}"
        '<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/>'
        "</w:sectPr></w:body></w:document>"
    )


def _styles_xml() -> str:
    return (
        f'<w:styles xmlns:w="{W_NS}">'
        "<w:docDefaults><w:rPrDefault><w:rPr/></w:rPrDefault>"
        "<w:pPrDefault><w:pPr/></w:pPrDefault></w:docDefaults>"
        '<w:style w:type="paragraph" w:default="1" w:styleId="Normal">'
        '<w:name w:val="Normal"/><w:qFormat/></w:style></w:styles>'
    )


def _write_docx(path: Path, *, document: str | None = None) -> Path:
    content_types = (
        '<?xml version="1.0" encoding="UTF-8"?>'
        '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
        '<Default Extension="xml" ContentType="application/xml"/>'
        '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
        '<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>'
        "</Types>"
    )
    root_rels = (
        '<?xml version="1.0" encoding="UTF-8"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>'
        "</Relationships>"
    )
    doc_rels = (
        '<?xml version="1.0" encoding="UTF-8"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rIdStyles" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>'
        "</Relationships>"
    )
    path.parent.mkdir(parents=True, exist_ok=True)
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as package:
        package.writestr("[Content_Types].xml", content_types)
        package.writestr("_rels/.rels", root_rels)
        package.writestr("word/_rels/document.xml.rels", doc_rels)
        package.writestr("word/document.xml", document or _document_xml())
        package.writestr("word/styles.xml", _styles_xml())
    return path


def _text_lines(docx_path: Path) -> list[str]:
    import re

    with zipfile.ZipFile(docx_path) as package:
        document = package.read("word/document.xml").decode("utf-8")
    lines = []
    for paragraph in re.findall(r"<w:p[ >][\s\S]*?</w:p>", document):
        # A structural tab renders as separation, so treat it as whitespace.
        spaced = re.sub(r"<w:tab[^>]*/>", "<w:t> </w:t>", paragraph)
        pieces = re.findall(r"<w:t[^>]*>([\s\S]*?)</w:t>", spaced)
        if any(piece.strip() for piece in pieces):
            lines.append(" ".join("".join(pieces).split()))
    return lines


def _run(tmp_path: Path, source: Path, mode: str, name: str):
    return format_specifications(
        architect_template=None,
        target_specs=[source],
        output_dir=tmp_path / name,
        api_key="",
        max_workers=1,
        conversion_mode=mode,
    )


def test_standalone_canadian_run_needs_no_architect(tmp_path: Path) -> None:
    source = _write_docx(tmp_path / "21 13 13 Sprinklers.docx")
    before = source.read_bytes()

    run = _run(tmp_path, source, CSI_TO_CANADIAN_STANDALONE, "out")

    assert run.success, "\n".join(run.targets[0].log)
    result = run.targets[0]
    assert result.output_path is not None
    assert result.output_path.name == "21 13 13 Sprinklers_CANADIAN.docx"
    # The source is never touched, and no template was analyzed.
    assert source.read_bytes() == before
    assert run.template_profile is None
    validate_docx_package(result.output_path)

    # The typed CSI markers are gone from the text; the built-in list renders
    # the numbers now.
    lines = _text_lines(result.output_path)
    assert "GENERAL" in lines
    assert not any(line.startswith("PART 1") for line in lines)
    assert not any(line.startswith("A. Section includes") for line in lines)

    with zipfile.ZipFile(result.output_path) as package:
        document = package.read("word/document.xml").decode("utf-8")
        numbering = package.read("word/numbering.xml").decode("utf-8")
    # The built-in CSC levels arrived...
    assert '<w:lvlText w:val="PART %1"/>' in numbering
    assert '<w:lvlText w:val="%1.%2"/>' in numbering
    # ...and the target's own page geometry did not move.
    assert TARGET_PAGE_SIZE in document


def test_standalone_canadian_run_records_the_builtin_scheme(tmp_path: Path) -> None:
    source = _write_docx(tmp_path / "spec.docx")
    run = _run(tmp_path, source, CSI_TO_CANADIAN_STANDALONE, "out")
    assert run.success, "\n".join(run.targets[0].log)

    manifest = json.loads(Path(run.manifest_path).read_text(encoding="utf-8"))
    assert manifest["numbering_scheme"] == "builtin_csc"
    # Present and explicitly null, so an absent template can never be mistaken
    # for one that simply was not recorded.
    assert manifest["architect_template"] == {"path": None, "sha256": None}
    assert manifest["builtin_scheme"] == {
        "version": builtin_scheme.BUILTIN_SCHEME_VERSION,
        "digest": builtin_scheme.scheme_digest(),
    }


def test_reverse_run_does_not_claim_a_scheme_it_never_used(tmp_path: Path) -> None:
    """Typed markers come from the document's own counters, not the CSC list."""

    source = _write_docx(tmp_path / "spec.docx")
    canadian = _run(tmp_path, source, CSI_TO_CANADIAN_STANDALONE, "canadian")
    assert canadian.success, "\n".join(canadian.targets[0].log)
    intermediate = canadian.targets[0].output_path
    assert intermediate is not None

    back = _run(tmp_path, intermediate, CANADIAN_TO_CSI, "csi")
    assert back.success, "\n".join(back.targets[0].log)
    manifest = json.loads(Path(back.manifest_path).read_text(encoding="utf-8"))
    assert manifest["numbering_scheme"] == "typed_csi"
    assert manifest["builtin_scheme"] is None
    assert manifest["architect_template"] == {"path": None, "sha256": None}


def test_canadian_to_csi_writes_typed_markers(tmp_path: Path) -> None:
    source = _write_docx(tmp_path / "spec.docx")
    canadian = _run(tmp_path, source, CSI_TO_CANADIAN_STANDALONE, "canadian")
    assert canadian.success, "\n".join(canadian.targets[0].log)
    intermediate = canadian.targets[0].output_path
    assert intermediate is not None

    back = _run(tmp_path, intermediate, CANADIAN_TO_CSI, "csi")
    assert back.success, "\n".join(back.targets[0].log)
    result = back.targets[0]
    assert result.output_path is not None
    assert result.output_path.name.endswith("_CSI.docx")
    validate_docx_package(result.output_path)

    # Every line the document started with is back, marker and text together.
    #
    # One documented exception: the forward converter treats a dash or colon
    # after ``PART n`` as part of the typed marker and removes it with the
    # marker, because the architect's ``PART %1`` numbering supplies its own
    # separator. That is existing behaviour of ``csi_to_canadian``, not a loss
    # introduced coming back, so the round trip returns "PART 1 GENERAL" from
    # "PART 1 - GENERAL". The requirement text itself is untouched.
    lines = _text_lines(result.output_path)
    def _round_tripped(marker: str, text: str) -> str:
        line = f"{marker} {text}" if marker else text
        if marker.startswith("PART"):
            line = line.replace(" - ", " ")
        return " ".join(line.split())

    expected = [_round_tripped(marker, text) for _role, marker, text in CSI_LINES]
    assert lines == expected


def test_supplying_a_template_to_a_builtin_mode_is_rejected(tmp_path: Path) -> None:
    source = _write_docx(tmp_path / "spec.docx")
    template = _write_docx(tmp_path / "architect.docx")

    with pytest.raises(ValueError) as raised:
        format_specifications(
            architect_template=template,
            target_specs=[source],
            output_dir=tmp_path / "out",
            api_key="",
            max_workers=1,
            conversion_mode=CSI_TO_CANADIAN_STANDALONE,
        )
    assert getattr(raised.value, "safe_error_code", None) == "input_architect_not_accepted"
    # It must also round-trip through the public diagnostic surface, or
    # ``run.json`` would record the refusal as an untrusted error instead of a
    # code with a remediation the user can act on.
    from spec_formatter import pipeline

    diagnostic = pipeline.safe_error_diagnostic(raised.value)
    assert diagnostic is not None
    assert diagnostic.code == "input_architect_not_accepted"
    assert "does not take an architect template" in diagnostic.message
