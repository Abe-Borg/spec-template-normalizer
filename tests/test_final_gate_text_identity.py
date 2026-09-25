"""The final gate proves text identity in every mode, except the predicted diff.

Each conversion mode verifies its own edits in memory, at the conversion
stage. Environment application, numbering and style import, classification
application and repackaging all run after that, so before WI-03 anything they
did to a paragraph's text published. These tests stand in for such a stage:
they damage ``word/document.xml`` in the extraction directory after every
edit has been made and before the output is packaged, which only the final
gate can see.

Every paragraph the conversion did not predict to change must come through
with its exact run content. Every paragraph it did predict must read, run by
run, exactly as predicted. The damage below is chosen to be invisible to a
whitespace-normalized text check where it can be -- a doubled space, a
non-breaking space, a dropped tab or soft hyphen -- because that is the class
of loss the normalized check was blind to.
"""

from __future__ import annotations

import json
import re
import zipfile
from pathlib import Path
from typing import Callable

import pytest

from spec_formatter.pipeline import (
    CANADIAN_TO_CSI,
    CSI_TO_CANADIAN,
    CSI_TO_CANADIAN_STANDALONE,
    format_specifications,
)
from spec_formatter.style_application import batch_runner
from spec_formatter.style_application.core.xml_helpers import iter_paragraph_xml_blocks
from spec_formatter.style_application.phase2_invariants import validate_docx_package
from tests import test_architect_free_modes as free
from tests.test_unified_roundtrip import (
    _deterministic_classifier,
    _diagnostics_event,
    _format_run_content_pair,
    _write_canadian_pair,
    _write_run_content_pair,
)

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
PREDICTION_MISMATCH = "conversion_prediction_mismatch"


# --- Driving each mode without an API key --------------------------------


def _architect_canadian_run(tmp_path: Path):
    """Run csi_to_canadian; returns the run and the target it converted."""

    architect, target = _write_canadian_pair(tmp_path)
    run = format_specifications(
        architect_template=architect,
        target_specs=[target],
        output_dir=tmp_path / "formatted",
        cache_dir=tmp_path / "template-cache",
        api_key="",
        max_workers=1,
        conversion_mode=CSI_TO_CANADIAN,
        template_model="wi03-canadian-fixture",
        template_classifier=_deterministic_classifier,
    )
    return run, target


def _standalone_run(tmp_path: Path):
    source = free._write_docx(tmp_path / "source" / "spec.docx")
    return free._run(tmp_path, source, CSI_TO_CANADIAN_STANDALONE, "canadian")


def _canadian_source(tmp_path: Path) -> Path:
    """The standalone Canadian output, used as the reverse mode's input.

    Built before any corruption is installed, so the forward run it comes
    from is clean.
    """

    forward = _standalone_run(tmp_path)
    assert forward.success, "\n".join(forward.targets[0].log)
    output = forward.targets[0].output_path
    assert output is not None
    return output


def _reverse_run(tmp_path: Path, source: Path):
    return free._run(tmp_path, source, CANADIAN_TO_CSI, "csi")


_SETTINGS_TRACKING = (
    f'<w:settings xmlns:w="{W_NS}"><w:trackRevisions/></w:settings>'
)


def _tracked_bold_canadian_source(tmp_path: Path) -> Path:
    """A Canadian spec under review whose text runs carry direct formatting.

    Track Changes on means every marker is written as this application's own
    tracked insertion; the bold runs are the shape that used to put that
    insertion inside the run.
    """

    canadian = _canadian_source(tmp_path)
    with zipfile.ZipFile(canadian) as package:
        parts = {name: package.read(name) for name in package.namelist()}
    document = parts["word/document.xml"].decode("utf-8")
    parts["word/document.xml"] = document.replace(
        "<w:r><w:t", "<w:r><w:rPr><w:b/></w:rPr><w:t"
    ).encode("utf-8")
    parts["word/settings.xml"] = _SETTINGS_TRACKING.encode("utf-8")
    content_types = parts["[Content_Types].xml"].decode("utf-8")
    assert "settings.xml" not in content_types
    parts["[Content_Types].xml"] = content_types.replace(
        "</Types>",
        '<Override PartName="/word/settings.xml" ContentType="application/'
        'vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/></Types>',
    ).encode("utf-8")
    relationships = parts["word/_rels/document.xml.rels"].decode("utf-8")
    parts["word/_rels/document.xml.rels"] = relationships.replace(
        "</Relationships>",
        '<Relationship Id="rIdTrackedSettings" Type="http://schemas.openxmlformats.org/'
        'officeDocument/2006/relationships/settings" Target="settings.xml"/>'
        "</Relationships>",
    ).encode("utf-8")
    tracked = tmp_path / "tracked" / canadian.name
    tracked.parent.mkdir(parents=True)
    with zipfile.ZipFile(tracked, "w", zipfile.ZIP_DEFLATED) as package:
        for name, payload in parts.items():
            package.writestr(name, payload)
    return tracked


# --- Damage after every edit, before packaging ----------------------------


def _damage_before_publication(
    monkeypatch: pytest.MonkeyPatch,
    damage: Callable[[bytes], bytes],
) -> None:
    real_build = batch_runner._build_and_patch_output

    def build_after_a_late_corruption(docx_path, extract_dir, *args, **kwargs):
        # Stands in for any step after the conversion that damages text:
        # environment application, numbering or style import, classification
        # application, repackaging. The final gate is the only check that
        # runs after all of them. Bytes, not text: Path.write_text translates
        # newlines on Windows.
        document = Path(extract_dir) / "word" / "document.xml"
        payload = document.read_bytes()
        damaged = damage(payload)
        assert damaged != payload, "fixture drifted: the damage did not apply"
        document.write_bytes(damaged)
        return real_build(docx_path, extract_dir, *args, **kwargs)

    monkeypatch.setattr(
        batch_runner, "_build_and_patch_output", build_after_a_late_corruption
    )


def _replace_once(old: bytes, new: bytes) -> Callable[[bytes], bytes]:
    def damage(payload: bytes) -> bytes:
        assert payload.count(old) >= 1, f"fixture drifted: {old!r} not found"
        return payload.replace(old, new, 1)

    return damage


def _assert_withheld_by_the_gate(run, paragraph_index: int, *, never_published: str):
    assert not run.success
    result = run.targets[0]
    assert result.output_path is None
    assert result.stage == "output_publication"
    assert result.error_code == PREDICTION_MISMATCH
    # The failure says where, as a validated placement...
    assert result.error_location is not None
    assert result.error_location["paragraph_index"] == paragraph_index
    assert not list(run.run_dir.glob("*.docx"))
    # ...the gate says it ran on the failure path too...
    fields = _diagnostics_event(run, "build_output")["fields"]
    assert fields["failed"] is True
    assert fields["body_signature_paragraphs_compared"] >= 1
    assert fields["body_paragraphs_expected_changed"] >= 1
    # ...and no artifact carries the document's text.
    for artifact in (run.manifest_path, run.diagnostics_path, run.run_dir / "run.log"):
        assert never_published not in Path(artifact).read_text(encoding="utf-8")
    return result


def _paragraphs(docx: Path) -> list[str]:
    with zipfile.ZipFile(docx) as package:
        document = package.read("word/document.xml").decode("utf-8")
    return [block for _start, _end, block in iter_paragraph_xml_blocks(document)]


def _assert_counters(run, source: Path, predicted: set[int]) -> None:
    """The gate's counters match what it had to compare, on a real run.

    Every predicted paragraph is compared against its prediction; any other
    paragraph is compared only when its XML changed, because identical XML
    has identical run content.
    """

    output = run.targets[0].output_path
    assert output is not None
    changed_xml = {
        index
        for index, (before, after) in enumerate(zip(_paragraphs(source), _paragraphs(output)))
        if before != after
    }
    fields = _diagnostics_event(run, "build_output")["fields"]
    assert fields["body_paragraphs_expected_changed"] == len(predicted)
    assert fields["body_signature_paragraphs_compared"] == len(predicted | changed_xml)


# --- csi_to_canadian (architect template) ---------------------------------
#
# Paragraphs: 0 "A. Work Included" and 2 "1. Pumps" lose their typed markers;
# 1 carries a section break; 3 and 4 are table cells; 5 hosts a text box.


def test_architect_canadian_happy_path_records_the_prediction(tmp_path: Path) -> None:
    run, target = _architect_canadian_run(tmp_path)

    assert run.success, "\n".join(run.targets[0].log)
    validate_docx_package(run.targets[0].output_path)
    report = run.targets[0].conversion_report
    assert report is not None and report.literal_markers_removed == 2
    _assert_counters(run, target, predicted={0, 2})


def test_architect_canadian_gate_rejects_a_change_outside_the_prediction(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    # A doubled space in a table cell: the normalized text is identical.
    _damage_before_publication(monkeypatch, _replace_once(b">Outer cell<", b">Outer  cell<"))

    run, _target = _architect_canadian_run(tmp_path)

    _assert_withheld_by_the_gate(run, 3, never_published="Outer")


def test_architect_canadian_gate_rejects_a_converted_paragraph_changed_beyond_its_marker(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    # The marker came off as predicted; then the kept text lost its exact
    # spacing, which reads the same once whitespace is collapsed.
    _damage_before_publication(
        monkeypatch, _replace_once(b">Work Included<", b">Work  Included<")
    )

    run, _target = _architect_canadian_run(tmp_path)

    _assert_withheld_by_the_gate(run, 0, never_published="Included")


def test_architect_canadian_gate_rejects_a_converted_paragraph_that_reads_differently(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    _damage_before_publication(monkeypatch, _replace_once(b">Pumps<", b">Pumps and valves<"))

    run, _target = _architect_canadian_run(tmp_path)

    _assert_withheld_by_the_gate(run, 2, never_published="valves")


# --- csi_to_canadian_standalone -------------------------------------------
#
# Paragraphs follow tests/test_architect_free_modes.CSI_LINES: 0 SECTION,
# 1 title, 2-11 typed CSI markers, 12 END OF SECTION.


def test_standalone_canadian_happy_path_records_the_prediction(tmp_path: Path) -> None:
    run = _standalone_run(tmp_path)

    assert run.success, "\n".join(run.targets[0].log)
    _assert_counters(run, tmp_path / "source" / "spec.docx", predicted=set(range(2, 12)))


def test_standalone_canadian_gate_rejects_a_change_outside_the_prediction(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    _damage_before_publication(
        monkeypatch,
        _replace_once(b"WET-PIPE SPRINKLER SYSTEMS", b"WET-PIPE  SPRINKLER SYSTEMS"),
    )

    run = _standalone_run(tmp_path)

    result = _assert_withheld_by_the_gate(run, 1, never_published="WET-PIPE")
    # The placement names the SECTION the paragraph belongs to, which is what
    # lets a user find it in Word.
    assert result.error_location["section_number"] == "21 13 13"


def test_standalone_canadian_gate_rejects_text_moved_out_of_sight(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    # Wrapping the title's run in a move-from keeps every run's content -- the
    # run-content signature does not record containers -- while Word stops
    # showing the words. The visible-text reading catches it.
    _damage_before_publication(
        monkeypatch,
        _replace_once(
            b'<w:r><w:t xml:space="preserve">WET-PIPE SPRINKLER SYSTEMS</w:t></w:r>',
            b'<w:moveFrom w:id="7" w:author="Someone" w:date="2026-01-01T00:00:00Z">'
            b'<w:r><w:t xml:space="preserve">WET-PIPE SPRINKLER SYSTEMS</w:t></w:r>'
            b"</w:moveFrom>",
        ),
    )

    run = _standalone_run(tmp_path)

    _assert_withheld_by_the_gate(run, 1, never_published="WET-PIPE")


def test_standalone_canadian_gate_rejects_a_non_breaking_space_in_a_converted_paragraph(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    _damage_before_publication(
        monkeypatch, _replace_once(b"per NFPA 13.", "per NFPA 13.".encode("utf-8"))
    )

    run = _standalone_run(tmp_path)

    _assert_withheld_by_the_gate(run, 6, never_published="Hydraulic")


def test_standalone_canadian_gate_rejects_a_dropped_tab_in_a_converted_paragraph(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    # The PART heading keeps a structural tab after its removed marker. Its
    # visible text is "GENERAL" with or without it.
    _damage_before_publication(monkeypatch, _replace_once(b"<w:r><w:tab/></w:r>", b""))

    run = _standalone_run(tmp_path)

    _assert_withheld_by_the_gate(run, 2, never_published="GENERAL")


# --- canadian_to_csi -------------------------------------------------------


def test_reverse_happy_path_records_the_prediction(tmp_path: Path) -> None:
    source = _canadian_source(tmp_path)

    run = _reverse_run(tmp_path, source)

    assert run.success, "\n".join(run.targets[0].log)
    _assert_counters(run, source, predicted=set(range(2, 12)))


def test_reverse_gate_rejects_a_change_outside_the_prediction(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    source = _canadian_source(tmp_path)
    # A soft hyphen is invisible to the normalized text check.
    _damage_before_publication(
        monkeypatch,
        _replace_once(
            b"END OF SECTION",
            b'END OF SEC</w:t><w:softHyphen/><w:t xml:space="preserve">TION',
        ),
    )

    run = _reverse_run(tmp_path, source)

    _assert_withheld_by_the_gate(run, 12, never_published="END OF")


def test_reverse_gate_rejects_a_marked_paragraph_changed_beyond_its_marker(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    source = _canadian_source(tmp_path)
    _damage_before_publication(
        monkeypatch, _replace_once(b"Provide listed", b"Provide  listed")
    )

    run = _reverse_run(tmp_path, source)

    _assert_withheld_by_the_gate(run, 11, never_published="Provide")


def test_reverse_gate_rejects_a_marker_that_lost_its_tab(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    source = _canadian_source(tmp_path)
    _damage_before_publication(
        monkeypatch,
        _replace_once(
            b'<w:t xml:space="preserve">A.</w:t><w:tab/>',
            b'<w:t xml:space="preserve">A.</w:t>',
        ),
    )

    run = _reverse_run(tmp_path, source)

    _assert_withheld_by_the_gate(run, 4, never_published="Section includes")


def test_tracked_reverse_happy_path_publishes_valid_revisions(tmp_path: Path) -> None:
    source = _tracked_bold_canadian_source(tmp_path)

    run = _reverse_run(tmp_path, source)

    assert run.success, "\n".join(run.targets[0].log)
    report = run.targets[0].conversion_report
    assert report is not None and report.markers_tracked is True
    output = run.targets[0].output_path
    validate_docx_package(output)
    with zipfile.ZipFile(output) as package:
        document = package.read("word/document.xml")
    assert b"<w:r><w:ins" not in document
    _assert_counters(run, source, predicted=set(range(2, 12)))


def test_tracked_reverse_gate_rejects_a_change_beyond_the_marker(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    source = _tracked_bold_canadian_source(tmp_path)
    _damage_before_publication(
        monkeypatch, _replace_once(b"Provide listed", b"Provide  listed")
    )

    run = _reverse_run(tmp_path, source)

    _assert_withheld_by_the_gate(run, 11, never_published="Provide")


def test_tracked_reverse_gate_rejects_a_marker_taken_out_of_its_revision(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    # Unwrapping this application's own insertion leaves every run's content
    # exactly as it was -- the w:ins wrapper is not run content -- but turns a
    # reviewable marker into permanent text. The report would still say the
    # markers were tracked.
    own_insertion = re.compile(
        rb'<w:ins\b[^>]*w:author="Specification Formatter"[^>]*>(<w:r>.*?</w:r>)</w:ins>',
        re.S,
    )

    def unwrap_first_marker(payload: bytes) -> bytes:
        return own_insertion.sub(rb"\1", payload, count=1)

    source = _tracked_bold_canadian_source(tmp_path)
    _damage_before_publication(monkeypatch, unwrap_first_marker)

    run = _reverse_run(tmp_path, source)

    _assert_withheld_by_the_gate(run, 2, never_published="GENERAL")


# --- format_only -----------------------------------------------------------


def test_format_only_gate_records_that_it_predicts_no_change(tmp_path: Path) -> None:
    architect, target = _write_run_content_pair(tmp_path)

    run = _format_run_content_pair(tmp_path, architect, target)

    assert run.success, "\n".join(run.targets[0].log)
    fields = _diagnostics_event(run, "build_output")["fields"]
    # Format-only passes an empty prediction: every paragraph must keep its
    # exact run content, as WI-02 made it.
    assert fields["body_paragraphs_expected_changed"] == 0
    assert fields["body_signature_paragraphs_compared"] >= 2
    manifest = json.loads(run.manifest_path.read_text(encoding="utf-8"))
    assert manifest["status"] == "succeeded"
