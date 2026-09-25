"""Tracked revisions are counted, reported and never collided with (WI-06).

Three related gaps, each closed here and held by these tests:

* **The census.** Nothing at the final gate counted tracked revisions. The
  run-property check catches a lost revision *subtree* indirectly, through run
  paths, and the body check sees runs, not the revision a run sits in -- so a
  reviewer's ``w:del`` turned into a ``w:ins``, a ``w:pPrChange`` dropped, or a
  reviewer's revision re-signed with this application's name all published.
  The gate now counts every revision element in ``word/document.xml`` by
  author and kind, before and after, and holds the difference to exactly what
  the conversion predicted: nothing in every mode, except one ``w:ins`` and one
  ``w:pPrChange`` per paragraph a tracked ``canadian_to_csi`` run converts,
  authored by this application.
* **Header and footer revisions.** Replacing the target's header set deletes
  its parts, and any pending tracked change inside them went with them
  silently; the architect's parts arrive with whatever pending changes *they*
  carry. Both are now counted, logged and recorded.
* **Revision ids.** The reverse converter numbered its revisions from a fixed
  900000, "without needing to scan", so a source whose own annotation ids
  already reached that range got colliding ids. It now allocates above every
  annotation id in the package.
"""

from __future__ import annotations

import json
import re
import xml.etree.ElementTree as ET
import zipfile
from pathlib import Path
from typing import Callable, Optional

import pytest

from spec_formatter.pipeline import (
    CANADIAN_TO_CSI,
    CSI_TO_CANADIAN,
    CSI_TO_CANADIAN_STANDALONE,
    format_specifications,
)
from spec_formatter.style_application.core.canadian_to_csi import (
    MARKER_REVISION_AUTHOR,
    apply_canadian_to_csi,
    plan_canadian_to_csi,
)
from spec_formatter.style_application.core.expected_changes import (
    ExpectedParagraphChange,
    prediction_mismatch,
)
from spec_formatter.style_application.core.revisions import (
    REVISION_KINDS,
    count_revisions,
    max_annotation_id,
    revision_census,
)
from spec_formatter.style_application.core.xml_helpers import (
    paragraph_run_content_signature,
    paragraph_text_from_block,
)
from spec_formatter.style_application.phase2_invariants import validate_docx_package
from tests import test_architect_free_modes as free
from tests import test_canadian_to_csi as reverse
from tests import test_unified_roundtrip as roundtrip
from tests.test_final_gate_text_identity import (
    _canadian_source,
    _damage_before_publication,
    _replace_once,
    _reverse_run,
    _tracked_bold_canadian_source,
)
from tests.test_package_change_whitelist import _with_parts

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
OFFICE_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
DATE = 'w:date="2026-01-01T00:00:00Z"'
CENSUS_FAIL = "INVARIANT FAIL: tracked revisions in word/document.xml changed"
PREDICTION_MISMATCH = "conversion_prediction_mismatch"

#: Words that appear only inside revisions in these fixtures. None of them may
#: reach a structured artifact or the run log.
REVISION_WORDS = ("wiregrass", "quillfeather", "marrowstone", "Reviewer")


# --- Fixtures ----------------------------------------------------------------


def _reviewer_deletion(revision_id: int, text: str) -> str:
    return (
        f'<w:del w:id="{revision_id}" w:author="Reviewer" {DATE}>'
        f'<w:r><w:delText xml:space="preserve"> {text}</w:delText></w:r></w:del>'
    )


def _rewrite(path: Path, edits: dict[str, Callable[[bytes], bytes]]) -> Path:
    with zipfile.ZipFile(path) as package:
        parts = {name: package.read(name) for name in package.namelist()}
    for name, edit in edits.items():
        edited = edit(parts[name])
        assert edited != parts[name], f"fixture drifted: {name} unchanged"
        parts[name] = edited
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as package:
        for name, payload in parts.items():
            package.writestr(name, payload)
    return path


def _replace(old: str, new: str) -> Callable[[bytes], bytes]:
    def edit(payload: bytes) -> bytes:
        text = payload.decode("utf-8")
        assert text.count(old) == 1, f"fixture drifted: {old!r}"
        return text.replace(old, new).encode("utf-8")

    return edit


def _format_only_pair(tmp_path: Path) -> tuple[Path, Path]:
    """The WI-02 run-content pair, plus a reviewer's property revision.

    The target already carries a reviewer's ``w:del``; its second paragraph
    gains a ``w:pPrChange``, a revision that holds no run content at all.
    """

    architect, target = roundtrip._write_run_content_pair(tmp_path)
    _rewrite(
        target,
        {
            "word/document.xml": _replace(
                '<w:pPr><w:pStyle w:val="TargetLevel2"/></w:pPr>',
                '<w:pPr><w:pStyle w:val="TargetLevel2"/>'
                f'<w:pPrChange w:id="92" w:author="Reviewer" {DATE}>'
                "<w:pPr/></w:pPrChange></w:pPr>",
            )
        },
    )
    return architect, target


def _format_only_run(tmp_path: Path):
    architect, target = _format_only_pair(tmp_path)
    return roundtrip._format_run_content_pair(tmp_path, architect, target), target


def _architect_canadian_run(tmp_path: Path):
    architect, target = roundtrip._write_canadian_pair(tmp_path)
    _rewrite(
        target,
        {
            "word/document.xml": _replace(
                "<w:r><w:t>Outer cell</w:t></w:r>",
                "<w:r><w:t>Outer cell</w:t></w:r>" + _reviewer_deletion(93, "wiregrass"),
            )
        },
    )
    run = format_specifications(
        architect_template=architect,
        target_specs=[target],
        output_dir=tmp_path / "formatted",
        cache_dir=tmp_path / "template-cache",
        api_key="",
        max_workers=1,
        conversion_mode=CSI_TO_CANADIAN,
        template_model="wi06-canadian-fixture",
        template_classifier=roundtrip._deterministic_classifier,
    )
    return run, target


def _standalone_source(tmp_path: Path) -> Path:
    """The architect-free CSI source with a reviewer's deletion in its title."""

    title = '<w:r><w:t xml:space="preserve">WET-PIPE SPRINKLER SYSTEMS</w:t></w:r>'
    document = free._document_xml()
    assert document.count(title) == 1
    document = document.replace(title, title + _reviewer_deletion(94, "quillfeather"))
    return free._write_docx(tmp_path / "source" / "spec.docx", document=document)


def _standalone_run(tmp_path: Path):
    source = _standalone_source(tmp_path)
    return free._run(tmp_path, source, CSI_TO_CANADIAN_STANDALONE, "canadian"), source


def _reverse_source(tmp_path: Path) -> Path:
    forward, _source = _standalone_run(tmp_path)
    assert forward.success, "\n".join(forward.targets[0].log)
    assert forward.targets[0].output_path is not None
    return forward.targets[0].output_path


def _reverse_mode_run(tmp_path: Path):
    source = _reverse_source(tmp_path)
    return _reverse_run(tmp_path, source), source


def _tracked_reverse_source(tmp_path: Path) -> Path:
    """A Canadian spec under review, with a reviewer's own pending deletion."""

    source = _tracked_bold_canadian_source(tmp_path)
    end = '<w:t xml:space="preserve">END OF SECTION 21 13 13</w:t></w:r>'
    return _rewrite(
        source,
        {"word/document.xml": _replace(end, end + _reviewer_deletion(95, "marrowstone"))},
    )


def _tracked_reverse_run(tmp_path: Path):
    source = _tracked_reverse_source(tmp_path)
    return _reverse_run(tmp_path, source), source


def _diagnostics_events(run, name: str) -> list[dict]:
    return [
        event
        for event in (
            json.loads(line)
            for line in run.diagnostics_path.read_text(encoding="utf-8").splitlines()
            if line.strip()
        )
        if event.get("event") == name
    ]


def _fields(run, name: str) -> dict:
    events = _diagnostics_events(run, name)
    assert len(events) == 1, name
    return events[0]["fields"]


def _document_census(docx: Path) -> dict[tuple[str, str], int]:
    """Revision elements by (author, local name), read without the engine."""

    with zipfile.ZipFile(docx) as package:
        root = ET.fromstring(package.read("word/document.xml"))
    counts: dict[tuple[str, str], int] = {}
    for element in root.iter():
        if not element.tag.startswith(f"{{{W_NS}}}"):
            continue
        local = element.tag.split("}", 1)[1]
        if local in ("ins", "del", "pPrChange", "rPrChange", "moveFrom", "moveTo"):
            key = (element.get(f"{{{W_NS}}}author", ""), local)
            counts[key] = counts.get(key, 0) + 1
    return counts


# --- The census: happy paths -------------------------------------------------

_UNTRACKED_RUNNERS: dict[str, Callable] = {
    "format_only": _format_only_run,
    CSI_TO_CANADIAN: _architect_canadian_run,
    CSI_TO_CANADIAN_STANDALONE: _standalone_run,
    CANADIAN_TO_CSI: _reverse_mode_run,
}
# The reviewer's revisions each mode's source carries into the run.
_SOURCE_REVISIONS = {
    "format_only": 2,
    CSI_TO_CANADIAN: 1,
    CSI_TO_CANADIAN_STANDALONE: 1,
    CANADIAN_TO_CSI: 1,
}


@pytest.mark.parametrize("mode", sorted(_UNTRACKED_RUNNERS))
def test_every_untracked_mode_keeps_every_revision_and_adds_none(
    tmp_path: Path, mode: str
) -> None:
    run, source = _UNTRACKED_RUNNERS[mode](tmp_path)

    assert run.success, "\n".join(run.targets[0].log)
    output = run.targets[0].output_path
    assert output is not None
    validate_docx_package(output)
    assert _document_census(output) == _document_census(source)

    fields = _fields(run, "build_output")
    assert fields["revisions_before"] == _SOURCE_REVISIONS[mode]
    assert fields["revisions_after"] == _SOURCE_REVISIONS[mode]
    assert fields["revisions_added_by_application"] == 0


def test_tracked_reverse_conversion_adds_exactly_two_revisions_per_paragraph(
    tmp_path: Path,
) -> None:
    run, source = _tracked_reverse_run(tmp_path)

    assert run.success, "\n".join(run.targets[0].log)
    report = run.targets[0].conversion_report
    assert report is not None and report.markers_tracked is True
    converted = report.paragraphs_converted
    assert converted == 10
    output = run.targets[0].output_path
    assert output is not None
    validate_docx_package(output)

    before, after = _document_census(source), _document_census(output)
    # The reviewer's deletion is still there, exactly once...
    assert before == {("Reviewer", "del"): 1}
    # ...and the application added one insertion and one property change per
    # converted paragraph, under its own name, and nothing else.
    assert after == {
        ("Reviewer", "del"): 1,
        (MARKER_REVISION_AUTHOR, "ins"): converted,
        (MARKER_REVISION_AUTHOR, "pPrChange"): converted,
    }
    fields = _fields(run, "build_output")
    assert fields["revisions_before"] == 1
    assert fields["revisions_after"] == 1 + 2 * converted
    assert fields["revisions_added_by_application"] == 2 * converted


# --- The census: damage only the census can see -------------------------------


def _assert_withheld(run, message: Optional[str]) -> dict:
    assert not run.success
    result = run.targets[0]
    assert result.output_path is None
    assert result.stage == "output_publication"
    if message is not None:
        assert message in str(result.error)
    assert not list(run.run_dir.glob("*.docx"))
    fields = _fields(run, "build_output")
    assert fields["failed"] is True
    for artifact in (run.manifest_path, run.diagnostics_path, run.run_dir / "run.log"):
        text = Path(artifact).read_text(encoding="utf-8")
        for word in REVISION_WORDS:
            assert word not in text, (artifact, word)
    return fields


def _regex_once(pattern: bytes, replacement: bytes) -> Callable[[bytes], bytes]:
    def damage(payload: bytes) -> bytes:
        damaged, count = re.subn(pattern, replacement, payload, count=1, flags=re.S)
        assert count == 1, f"fixture drifted: {pattern!r} not found"
        return damaged

    return damage


def test_a_removed_deletion_fails_and_the_census_says_so(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    # The whole revision gone, content and all: the body check reports it in
    # its own terms first, and the census has still counted the loss.
    _damage_before_publication(
        monkeypatch, _regex_once(rb'<w:del w:id="91".*?</w:del>', b"")
    )

    run, _target = _format_only_run(tmp_path)

    fields = _assert_withheld(run, "FORMAT_ONLY INVARIANT FAIL")
    assert fields["revisions_before"] == 2
    assert fields["revisions_after"] == 1
    assert fields["revisions_added_by_application"] == 0


def test_a_deletion_turned_into_an_insertion_is_refused(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    # Same run, same content, a different revision: visible text and run
    # content cannot see it. Accepting it would now keep what the reviewer
    # deleted.
    _damage_before_publication(
        monkeypatch,
        _regex_once(rb'<w:del (w:id="91".*?)</w:del>', rb"<w:ins \1</w:ins>"),
    )

    run, _target = _format_only_run(tmp_path)

    fields = _assert_withheld(run, CENSUS_FAIL)
    assert "del by another author" in str(run.targets[0].error)
    assert fields["revisions_before"] == 2
    assert fields["revisions_after"] == 2


def test_a_dropped_property_revision_is_refused(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    _damage_before_publication(
        monkeypatch,
        _regex_once(rb'<w:pPrChange w:id="92".*?</w:pPrChange>', b""),
    )

    run, _target = _format_only_run(tmp_path)

    _assert_withheld(run, CENSUS_FAIL)
    assert "pPrChange by another author" in str(run.targets[0].error)


def test_a_reviewers_revision_signed_with_the_applications_name_is_refused(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    # The projection of this application's own revisions is scoped by author,
    # so a reviewer's revision re-signed with this application's name would
    # vanish from every comparison made through it. The census is by author.
    _damage_before_publication(
        monkeypatch,
        _replace_once(
            b'w:id="94" w:author="Reviewer"',
            f'w:id="94" w:author="{MARKER_REVISION_AUTHOR}"'.encode("utf-8"),
        ),
    )

    run, _source = _standalone_run(tmp_path)

    _assert_withheld(run, CENSUS_FAIL)


def test_a_dropped_marker_property_revision_is_refused_where_it_belongs(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    # Without its w:pPrChange, rejecting the marker's insertion leaves the
    # paragraph with its numbering suppressed and no number at all. The source
    # is built first: its own forward run must not be damaged.
    source = _tracked_reverse_source(tmp_path)
    _damage_before_publication(
        monkeypatch,
        _regex_once(
            rb'<w:pPrChange [^>]*w:author="Specification Formatter".*?</w:pPrChange>',
            b"",
        ),
    )

    run = _reverse_run(tmp_path, source)

    assert not run.success
    result = run.targets[0]
    assert result.error_code == PREDICTION_MISMATCH
    assert result.error_location is not None
    assert result.error_location["paragraph_index"] == 2
    _assert_withheld(run, None)


def test_an_unpredicted_revision_by_the_application_is_refused(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    # A property revision in this application's name on the section title,
    # which the conversion never touched: no run content changes.
    title = b'<w:t xml:space="preserve">WET-PIPE SPRINKLER SYSTEMS</w:t>'

    def damage(payload: bytes) -> bytes:
        index = payload.index(title)
        paragraph = payload.rindex(b"<w:p>", 0, index)
        ppr = payload.index(b"<w:pPr/>", paragraph)
        assert ppr < index, "fixture drifted: the title has no empty w:pPr"
        return (
            payload[:ppr]
            + b'<w:pPr><w:pPrChange w:id="999999" w:author="Specification Formatter" '
            + DATE.encode("utf-8")
            + b"><w:pPr/></w:pPrChange></w:pPr>"
            + payload[ppr + len(b"<w:pPr/>"):]
        )

    source = _tracked_reverse_source(tmp_path)
    _damage_before_publication(monkeypatch, damage)

    run = _reverse_run(tmp_path, source)

    fields = _assert_withheld(run, CENSUS_FAIL)
    assert "pPrChange by this application" in str(run.targets[0].error)
    assert fields["revisions_added_by_application"] == 21


# --- Header and footer revisions ---------------------------------------------

_TARGET_HEADER = (
    f'<w:hdr xmlns:w="{W_NS}"><w:p><w:r><w:t>Old target header</w:t></w:r>'
    f'<w:ins w:id="11" w:author="Reviewer" {DATE}><w:r><w:t> wiregrass</w:t></w:r></w:ins>'
    f'<w:del w:id="12" w:author="Reviewer" {DATE}><w:r><w:delText> quillfeather</w:delText>'
    "</w:r></w:del></w:p></w:hdr>"
)
_TARGET_FOOTER = (
    f'<w:ftr xmlns:w="{W_NS}"><w:p><w:r><w:t>Target footer</w:t></w:r>'
    f'<w:ins w:id="13" w:author="Reviewer" {DATE}><w:r><w:t> marrowstone</w:t></w:r></w:ins>'
    "</w:p></w:ftr>"
)
_ARCHITECT_FIRST_HEADER = (
    f'<w:hdr xmlns:w="{W_NS}"><w:p><w:r><w:t>First header</w:t></w:r>'
    f'<w:ins w:id="1" w:author="Reviewer" {DATE}><w:r><w:t> wiregrass</w:t></w:r></w:ins>'
    "</w:p></w:hdr>"
)


def _header_revision_pair(tmp_path: Path) -> tuple[Path, Path]:
    architect = tmp_path / "architect.docx"
    target = tmp_path / "target.docx"
    roundtrip._write_docx(architect, architect=True)
    roundtrip._write_docx(target, architect=False)
    roundtrip._rewrite_docx_parts(architect, {"word/header2.xml": _ARCHITECT_FIRST_HEADER})
    roundtrip._rewrite_docx_parts(target, {"word/header9.xml": _TARGET_HEADER})
    _with_parts(
        target,
        {
            "word/footer9.xml": (
                _TARGET_FOOTER.encode("utf-8"),
                "application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml",
                f"{OFFICE_REL}/footer",
            )
        },
    )
    return architect, target


def _format(tmp_path: Path, architect: Path, target: Path):
    return format_specifications(
        architect_template=architect,
        target_specs=[target],
        output_dir=tmp_path / "formatted",
        cache_dir=tmp_path / "template-cache",
        api_key="",
        max_workers=1,
        template_model="wi06-header-revisions-fixture",
        template_classifier=roundtrip._deterministic_classifier,
    )


def test_header_and_footer_revisions_are_counted_logged_and_recorded(
    tmp_path: Path,
) -> None:
    architect, target = _header_revision_pair(tmp_path)

    run = _format(tmp_path, architect, target)

    assert run.success, "\n".join(run.targets[0].log)
    # The target's header set was replaced, taking three pending revisions
    # with it; the architect's first-page header brought one.
    fields = _fields(run, "apply_environment")
    assert fields["header_footer_revisions_discarded"] == 3
    assert fields["header_footer_revisions_imported"] == 1

    # One warning per part, naming the part and the count, verbatim in run.log.
    run_log = (run.run_dir / "run.log").read_text(encoding="utf-8")
    assert (
        "WARNING: Discarded tracked revisions in replaced target part "
        "word/header9.xml: 2"
    ) in run_log
    assert (
        "WARNING: Discarded tracked revisions in replaced target part "
        "word/footer9.xml: 1"
    ) in run_log
    assert (
        "WARNING: Imported tracked revisions in architect part word/header2.xml: 1"
    ) in run_log

    # A warning-level event, so the counts survive any diagnostics verbosity.
    (warning,) = _diagnostics_events(run, "header_footer_revisions")
    assert warning["level"] == "WARNING"
    assert warning["fields"]["discarded"] == 3
    assert warning["fields"]["imported"] == 1
    manifest = json.loads(run.manifest_path.read_text(encoding="utf-8"))
    assert manifest["diagnostics"]["warnings"] >= 1

    # The audit carries the same counts.
    audit = json.loads(Path(run.targets[0].audit_path).read_text(encoding="utf-8"))
    (environment,) = [
        event for event in audit["diagnostics"] if event["event"] == "apply_environment"
    ]
    assert environment["fields"]["header_footer_revisions_discarded"] == 3
    assert environment["fields"]["header_footer_revisions_imported"] == 1

    # Counts and part names only: not a word of the revisions themselves.
    for artifact in (
        run.manifest_path,
        run.diagnostics_path,
        run.run_dir / "run.log",
        Path(run.targets[0].audit_path),
    ):
        text = artifact.read_text(encoding="utf-8")
        for word in REVISION_WORDS:
            assert word not in text, (artifact, word)


def test_a_clean_header_set_records_zero_rather_than_nothing(tmp_path: Path) -> None:
    architect = tmp_path / "architect.docx"
    target = tmp_path / "target.docx"
    roundtrip._write_docx(architect, architect=True)
    roundtrip._write_docx(target, architect=False)

    run = _format(tmp_path, architect, target)

    assert run.success, "\n".join(run.targets[0].log)
    fields = _fields(run, "apply_environment")
    assert fields["header_footer_revisions_discarded"] == 0
    assert fields["header_footer_revisions_imported"] == 0
    assert _diagnostics_events(run, "header_footer_revisions") == []
    run_log = (run.run_dir / "run.log").read_text(encoding="utf-8")
    assert "tracked revisions in" not in run_log


# --- Revision ids -------------------------------------------------------------

_HIGH_ID_ROWS = [("PART", "GENERAL"), ("ARTICLE", "SUMMARY"), ("PARAGRAPH", "Scope.")]


def _high_id_paragraphs() -> list[str]:
    """Paragraphs whose own annotation ids sit where the old base allocated."""

    paragraphs = []
    for index, (role, text) in enumerate(_HIGH_ID_ROWS):
        style = reverse._GEOMETRY_ROLE_STYLE[role]
        run = f'<w:r><w:t xml:space="preserve">{text}</w:t></w:r>'
        if index == 0:
            run = (
                '<w:bookmarkStart w:id="900000" w:name="_Toc1"/>'
                + run
                + '<w:bookmarkEnd w:id="900000"/>'
            )
        if index == 1:
            run += (
                f'<w:ins w:id="900001" w:author="Reviewer" {DATE}>'
                "<w:r><w:t> under review</w:t></w:r></w:ins>"
            )
        paragraphs.append(f'<w:p><w:pPr><w:pStyle w:val="{style}"/></w:pPr>{run}</w:p>')
    return paragraphs


def _own_revision_ids(document_xml: str) -> list[int]:
    return [
        int(match)
        for match in re.findall(
            rf'<w:(?:ins|pPrChange)\b[^>]*w:id="(-?\d+)"[^>]*w:author="{MARKER_REVISION_AUTHOR}"',
            document_xml,
        )
    ]


def test_marker_ids_are_allocated_above_every_id_in_the_document() -> None:
    document = reverse._document(_high_id_paragraphs())

    plan = plan_canadian_to_csi(
        document,
        reverse._geometry_styles_xml(),
        reverse._classifications([role for role, _text in _HIGH_ID_ROWS]),
        numbering_xml=reverse._geometry_numbering_xml(),
        settings_xml=reverse._TRACKING_ON,
        revision_date="2026-01-01T00:00:00Z",
    )

    allocated = _own_revision_ids(plan.document_xml)
    assert len(allocated) == 2 * len(_HIGH_ID_ROWS)
    assert len(set(allocated)) == len(allocated)
    # Above the highest id already there, in the order they were written.
    assert min(allocated) > 900001
    assert allocated == sorted(allocated)


def _extracted_high_id_target(root: Path, *, note_id: int) -> Path:
    """An extraction directory whose highest id lives outside the body."""

    word = root / "word"
    word.mkdir(parents=True)
    (word / "document.xml").write_text(
        reverse._document(
            [
                f'<w:p><w:pPr><w:pStyle w:val="{reverse._GEOMETRY_ROLE_STYLE[role]}"/>'
                f'</w:pPr><w:r><w:t xml:space="preserve">{text}</w:t></w:r></w:p>'
                for role, text in _HIGH_ID_ROWS
            ]
        ),
        encoding="utf-8",
    )
    (word / "styles.xml").write_text(reverse._geometry_styles_xml(), encoding="utf-8")
    (word / "numbering.xml").write_text(reverse._geometry_numbering_xml(), encoding="utf-8")
    (word / "settings.xml").write_text(reverse._TRACKING_ON, encoding="utf-8")
    (word / "footnotes.xml").write_text(
        f'<w:footnotes xmlns:w="{W_NS}"><w:footnote w:id="1"><w:p>'
        f'<w:ins w:id="{note_id}" w:author="Reviewer" {DATE}><w:r><w:t>note</w:t>'
        "</w:r></w:ins></w:p></w:footnote></w:footnotes>",
        encoding="utf-8",
    )
    return root


def test_marker_ids_are_allocated_above_ids_in_the_other_parts(tmp_path: Path) -> None:
    # The old base would have given the second paragraph's revisions
    # 900002 and 900003: the second is the footnote's insertion.
    target = _extracted_high_id_target(tmp_path / "target", note_id=900003)

    apply_canadian_to_csi(
        target,
        reverse._classifications([role for role, _text in _HIGH_ID_ROWS]),
        [],
    )

    document = (target / "word" / "document.xml").read_text(encoding="utf-8")
    allocated = _own_revision_ids(document)
    assert len(allocated) == 2 * len(_HIGH_ID_ROWS)
    assert len(set(allocated)) == len(allocated)
    assert min(allocated) > 900003


# --- The building blocks ------------------------------------------------------


def test_the_census_counts_every_revision_kind_by_author_and_nothing_else() -> None:
    xml = (
        f'<w:document xmlns:w="{W_NS}"><w:body><w:p><w:pPr><w:rPr>'
        f'<w:ins w:id="1" w:author="A" {DATE}/></w:rPr>'
        f'<w:pPrChange w:id="2" w:author="B" {DATE}><w:pPr/></w:pPrChange></w:pPr>'
        f'<w:ins w:id="3" w:author="A" {DATE}><w:r><w:rPr>'
        f'<w:rPrChange w:id="4" w:author="A" {DATE}><w:rPr/></w:rPrChange></w:rPr>'
        "<w:t>x</w:t></w:r></w:ins>"
        '<w:bookmarkStart w:id="5" w:name="b"/><w:bookmarkEnd w:id="5"/>'
        "<!-- <w:del w:id=\"6\" w:author=\"A\"/> -->"
        "</w:p></w:body></w:document>"
    )

    census = revision_census(xml, "word/document.xml")

    assert census == {("A", "ins"): 2, ("B", "pPrChange"): 1, ("A", "rPrChange"): 1}
    assert count_revisions(xml, "word/document.xml") == 4
    assert max_annotation_id(xml, "word/document.xml") == 5
    assert {"ins", "del", "moveFrom", "moveTo", "pPrChange", "rPrChange",
            "sectPrChange", "tblPrChange", "trPrChange", "tcPrChange",
            "numberingChange"} <= set(REVISION_KINDS)


def test_the_census_reads_a_utf16_part() -> None:
    xml = (
        '<?xml version="1.0" encoding="UTF-16"?>'
        f'<w:ftr xmlns:w="{W_NS}"><w:p><w:del w:id="7" w:author="R" {DATE}/></w:p></w:ftr>'
    )
    payload = b"\xff\xfe" + xml.encode("utf-16-le")

    assert count_revisions(payload, "word/footer1.xml") == 1
    assert max_annotation_id(payload, "word/footer1.xml") == 7


def test_a_prediction_holds_the_revisions_a_paragraph_gains() -> None:
    source = '<w:p><w:r><w:t xml:space="preserve">GENERAL</w:t></w:r></w:p>'
    ins = f'<w:ins w:id="1" w:author="{MARKER_REVISION_AUTHOR}" {DATE}>'
    marked = (
        "<w:p><w:pPr>"
        f'<w:pPrChange w:id="2" w:author="{MARKER_REVISION_AUTHOR}" {DATE}>'
        "<w:pPr/></w:pPrChange></w:pPr>"
        f'{ins}<w:r><w:t xml:space="preserve">PART 1</w:t><w:tab/></w:r></w:ins>'
        '<w:r><w:t xml:space="preserve">GENERAL</w:t></w:r></w:p>'
    )
    change = ExpectedParagraphChange(
        visible_text=paragraph_text_from_block(marked),
        run_content=paragraph_run_content_signature(marked),
        run_content_outside_own_revisions=paragraph_run_content_signature(source),
        own_revisions=("ins", "pPrChange"),
    )

    assert prediction_mismatch(source, marked, change) is None
    without_property_revision = re.sub(r"<w:pPr>.*?</w:pPr>", "", marked, count=1)
    mismatch = prediction_mismatch(source, without_property_revision, change)
    assert mismatch is not None
    assert mismatch.check == "own_revision_kinds"
    assert mismatch.kind is None

    with pytest.raises(ValueError):
        ExpectedParagraphChange(
            visible_text="",
            run_content=(),
            run_content_outside_own_revisions=(),
            own_revisions=("bookmarkStart",),
        )
