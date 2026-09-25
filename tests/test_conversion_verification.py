"""Independent verification of a Canadian -> CSI conversion.

**This module shares no helper code with the engine on purpose.** Everything
below reads the before and after packages with the standard library only. The
engine's own invariants and this file could otherwise agree because they make
the same mistake -- which is precisely what happened to the document that
motivated these checks: it passed every text-level check the engine had, and
XSD, with its outline visibly destroyed.

The checks are the ones a formatting change has to survive before anyone should
believe it: what changed in the package, what changed in the text, what
happened to structural children that render as characters but never appear in
``w:t``, what happened to tracked revisions, and whether the numbers still read
correctly under both accept-all and reject-all.
"""

from __future__ import annotations

import hashlib
import re
import zipfile
import xml.etree.ElementTree as ET
from pathlib import Path

import pytest

from spec_formatter.style_application.core.canadian_to_csi import (
    apply_canadian_to_csi,
    plan_canadian_to_csi,
)

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"


def _q(local: str) -> str:
    return f"{{{W}}}{local}"


# --- The document under test ---------------------------------------------
#
# Modelled on the real template the defect was found in: list styles that
# carry w:numPr and no w:ind, so every indent in the document comes from the
# numbering level and disappears with it.

_LEVELS = {
    0: (720, 720, "decimal", "%1"),
    1: (720, 720, "decimalZero", "%1.%2"),
    2: (1259, 539, "decimal", ".%3"),
    3: (1797, 538, "decimal", ".%4"),
}
_ROLE_STYLE = {"PART": "L0", "ARTICLE": "L1", "PARAGRAPH": "L2", "SUBPARAGRAPH": "L3"}

_BODY = [
    ("PART", "GENERAL"),
    ("ARTICLE", "SUMMARY"),
    ("PARAGRAPH", "Section includes standpipe piping and valves."),
    ("SUBPARAGRAPH", "Related requirement one."),
    ("SUBPARAGRAPH", "Related requirement two."),
    ("ARTICLE", "REFERENCE STANDARDS"),
    ("PARAGRAPH", "NFPA 14, Standpipe and Hose Systems."),
    ("PART", "PRODUCTS"),
    ("ARTICLE", "DESIGN CRITERIA"),
    ("PARAGRAPH", "Design the standpipe system to NFPA 14."),
]


def _numbering_xml() -> str:
    levels = "".join(
        f'<w:lvl w:ilvl="{ilvl}"><w:start w:val="1"/><w:numFmt w:val="{fmt}"/>'
        f'<w:lvlText w:val="{text}"/><w:pPr>'
        f'<w:tabs><w:tab w:val="num" w:pos="{left}"/></w:tabs>'
        f'<w:ind w:left="{left}" w:hanging="{hang}"/></w:pPr></w:lvl>'
        for ilvl, (left, hang, fmt, text) in sorted(_LEVELS.items())
    )
    return (
        f'<w:numbering xmlns:w="{W}"><w:abstractNum w:abstractNumId="1">{levels}'
        f'</w:abstractNum><w:num w:numId="4"><w:abstractNumId w:val="1"/></w:num>'
        f"</w:numbering>"
    )


def _styles_xml() -> str:
    styles = "".join(
        f'<w:style w:type="paragraph" w:styleId="L{ilvl}"><w:name w:val="L {ilvl}"/>'
        f'<w:basedOn w:val="Normal"/><w:pPr><w:numPr><w:ilvl w:val="{ilvl}"/>'
        f'<w:numId w:val="4"/></w:numPr></w:pPr></w:style>'
        for ilvl in sorted(_LEVELS)
    )
    return (
        f'<w:styles xmlns:w="{W}"><w:style w:type="paragraph" w:styleId="Normal">'
        f'<w:name w:val="Normal"/></w:style>{styles}</w:styles>'
    )


def _document_xml(*, tracked_insertion_at: int | None = None) -> str:
    paragraphs = []
    for index, (role, text) in enumerate(_BODY):
        run = f'<w:r><w:t xml:space="preserve">{text}</w:t></w:r>'
        if index == tracked_insertion_at:
            run += (
                '<w:ins w:id="900" w:author="Reviewer" w:date="2026-01-01T00:00:00Z">'
                '<w:r><w:t xml:space="preserve"> Added under review.</w:t></w:r></w:ins>'
            )
        paragraphs.append(
            f'<w:p><w:pPr><w:pStyle w:val="{_ROLE_STYLE[role]}"/></w:pPr>{run}</w:p>'
        )
    return f'<w:document xmlns:w="{W}"><w:body>{"".join(paragraphs)}</w:body></w:document>'


_CONTENT_TYPES = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
    '<Default Extension="xml" ContentType="application/xml"/>'
    '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
    '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
    "</Types>"
)
_ROOT_RELS = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
    '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>'
    "</Relationships>"
)


_TRACKING_ON = f'<w:settings xmlns:w="{W}"><w:trackRevisions/></w:settings>'


def _write_docx(path: Path, document_xml: str, settings_xml: str = "") -> None:
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", _CONTENT_TYPES)
        z.writestr("_rels/.rels", _ROOT_RELS)
        z.writestr("word/document.xml", document_xml)
        z.writestr("word/styles.xml", _styles_xml())
        z.writestr("word/numbering.xml", _numbering_xml())
        if settings_xml:
            z.writestr("word/settings.xml", settings_xml)


@pytest.fixture
def converted(tmp_path: Path) -> tuple[Path, Path]:
    """Run the real converter and package its output. Returns (before, after)."""

    before = tmp_path / "before.docx"
    _write_docx(before, _document_xml(tracked_insertion_at=2))

    plan = plan_canadian_to_csi(
        _document_xml(tracked_insertion_at=2),
        _styles_xml(),
        {
            "classifications": [
                {"paragraph_index": index, "csi_role": role}
                for index, (role, _text) in enumerate(_BODY)
            ]
        },
        numbering_xml=_numbering_xml(),
    )
    after = tmp_path / "after.docx"
    _write_docx(after, plan.document_xml)
    return before, after


# --- Independent extractors (stdlib only) ---------------------------------

_MARKER = re.compile(
    r"^(PART \d+|\d+\.\d+|[A-Z]\.|\d+\.|[a-z]\.|\d+\)|[a-z]\)|\(\d+\)|\([a-z]\))\t"
)


def _members(path: Path) -> dict[str, str]:
    with zipfile.ZipFile(path) as z:
        return {n: hashlib.md5(z.read(n)).hexdigest() for n in z.namelist()}


def _paragraphs(path: Path) -> list[ET.Element]:
    with zipfile.ZipFile(path) as z:
        return list(ET.fromstring(z.read("word/document.xml")).iter(_q("p")))


def _walk_text(element: ET.Element, *, in_run: bool, skip_ins: bool, skip_del: bool) -> list[str]:
    """Collect rendered characters, tracking which element we are inside.

    ``w:tab`` is the reason this is a recursive walk rather than ``.iter()``:
    ``w:pPr/w:tabs/w:tab`` shares the element name but defines a tab *stop*
    instead of rendering a character. A flat scan counts both, and this
    conversion legitimately adds one of each per paragraph -- so a flat scan
    reads the two as cancelling out, or as one change when there were two.
    """

    out: list[str] = []
    for child in element:
        if child.tag == _q("ins") and skip_ins:
            continue
        if child.tag == _q("del") and skip_del:
            continue
        if child.tag == _q("t"):
            out.append(child.text or "")
        elif child.tag == _q("delText"):
            out.append(child.text or "")
        elif child.tag == _q("tab"):
            if in_run:
                out.append("\t")
        else:
            out.extend(
                _walk_text(
                    child,
                    in_run=in_run or child.tag == _q("r"),
                    skip_ins=skip_ins,
                    skip_del=skip_del,
                )
            )
    return out


def _text(paragraph: ET.Element) -> str:
    """Every rendered character, both revision views collapsed together."""

    return "".join(_walk_text(paragraph, in_run=False, skip_ins=False, skip_del=False))


def _run_tabs(paragraph: ET.Element) -> int:
    """Count only ``w:tab`` whose parent is ``w:r``."""

    return sum(1 for run in paragraph.findall(_q("r")) for _t in run.findall(_q("tab")))


def _revision_counts(path: Path) -> dict[str, tuple[int, int]]:
    counts = {}
    with zipfile.ZipFile(path) as z:
        for name in z.namelist():
            if not name.startswith("word/") or not name.endswith(".xml"):
                continue
            blob = z.read(name)
            pair = (blob.count(b"<w:ins "), blob.count(b"<w:del "))
            if pair != (0, 0):
                counts[name] = pair
    return counts


def _view(path: Path, mode: str) -> list[str]:
    """Per-paragraph rendered text under accept-all or reject-all."""

    result = []
    for paragraph in _paragraphs(path):
        ppr = paragraph.find(_q("pPr"))
        mark_inserted = (
            ppr is not None and ppr.find(f"{_q('rPr')}/{_q('ins')}") is not None
        )
        if mode == "reject" and mark_inserted:
            # A rejected inserted paragraph mark merges the paragraph away.
            continue
        result.append(
            "".join(
                _walk_text(
                    paragraph,
                    in_run=False,
                    skip_ins=mode == "reject",
                    skip_del=mode == "accept",
                )
            )
        )
    return result


# --- The checks -----------------------------------------------------------


def test_package_hash_whitelist(converted) -> None:
    """Only the parts we meant to touch changed; nothing added or removed."""

    before, after = converted
    before_members, after_members = _members(before), _members(after)
    assert set(before_members) == set(after_members)
    changed = {n for n in before_members if before_members[n] != after_members[n]}
    assert changed == {"word/document.xml"}


def test_text_changes_are_exactly_the_enumerated_markers(converted) -> None:
    """Every text diff is a prepended marker and tab, and nothing else.

    Asserted as an exact enumerated list rather than a count, because "roughly
    the right number of changes" is how a wrong one gets through.
    """

    before, after = converted
    before_paragraphs, after_paragraphs = _paragraphs(before), _paragraphs(after)
    assert len(before_paragraphs) == len(after_paragraphs)

    diffs = []
    for index, (b, a) in enumerate(zip(before_paragraphs, after_paragraphs)):
        before_text, after_text = _text(b), _text(a)
        if before_text == after_text:
            continue
        match = _MARKER.match(after_text)
        assert match, f"paragraph {index} changed beyond a marker: {after_text!r}"
        assert after_text[match.end():] == before_text, index
        diffs.append((index, match.group(1)))

    assert diffs == [
        (0, "PART 1"),
        (1, "1.1"),
        (2, "A."),
        (3, "1."),
        (4, "2."),
        (5, "1.2"),
        (6, "A."),
        (7, "PART 2"),
        (8, "2.1"),
        (9, "A."),
    ]


def test_structural_children_change_only_by_the_marker_tabs(converted) -> None:
    """Characters that never appear in ``w:t`` are accounted for too."""

    before, after = converted
    before_paragraphs, after_paragraphs = _paragraphs(before), _paragraphs(after)

    added_tabs = sum(_run_tabs(a) - _run_tabs(b) for b, a in zip(before_paragraphs, after_paragraphs))
    assert added_tabs == len(_BODY)

    for local in ("br", "noBreakHyphen"):
        delta = sum(
            len(a.findall(f".//{_q(local)}")) - len(b.findall(f".//{_q(local)}"))
            for b, a in zip(before_paragraphs, after_paragraphs)
        )
        assert delta == 0, local


def test_run_structure_is_unchanged(converted) -> None:
    """The marker joins an existing run rather than adding one."""

    before, after = converted
    for b, a in zip(_paragraphs(before), _paragraphs(after)):
        assert len(b.findall(f".//{_q('r')}")) == len(a.findall(f".//{_q('r')}"))


def test_tracked_revisions_are_untouched(converted) -> None:
    """A reviewer's pending edits survive the conversion exactly."""

    before, after = converted
    assert _revision_counts(before) == _revision_counts(after)
    assert _revision_counts(before), "fixture must actually carry a revision"


@pytest.mark.parametrize("mode", ["accept", "reject"])
def test_marker_sequence_holds_in_both_revision_views(converted, mode) -> None:
    """The numbers still read correctly however the revisions are resolved.

    A literal marker cannot renumber itself, so a sequence that is only correct
    in one view is a document that silently goes wrong later.
    """

    _before, after = converted
    markers = [_MARKER.match(t).group(1) for t in _view(after, mode) if _MARKER.match(t)]
    assert markers, "check must not pass vacuously"

    parts = [m for m in markers if m.startswith("PART ")]
    assert parts == [f"PART {i}" for i in range(1, len(parts) + 1)]

    articles: dict[str, list[int]] = {}
    for marker in markers:
        if re.fullmatch(r"\d+\.\d+", marker):
            part, number = marker.split(".")
            articles.setdefault(part, []).append(int(number))
    assert articles
    for numbers in articles.values():
        assert numbers == list(range(1, len(numbers) + 1))


def test_numbering_geometry_survives_in_the_output(converted) -> None:
    """The indent each level supplied is still on the paragraph.

    The check the engine did not have, stated independently: this fixture's
    styles carry no ``w:ind`` at all, so if the level's geometry is not written
    onto the paragraph it is gone from the document entirely.
    """

    _before, after = converted
    for index, (role, _text) in enumerate(_BODY):
        ilvl = int(_ROLE_STYLE[role][1])
        left, hanging, _fmt, _lvl_text = _LEVELS[ilvl]
        ppr = _paragraphs(after)[index].find(_q("pPr"))
        ind = ppr.find(_q("ind"))
        assert ind is not None, f"paragraph {index} lost its indentation"
        assert ind.get(_q("left")) == str(left)
        assert ind.get(_q("hanging")) == str(hanging)


# --- The same document, converted with revision tracking on ---------------


@pytest.fixture
def converted_tracked(tmp_path: Path) -> tuple[Path, Path]:
    before = tmp_path / "before.docx"
    _write_docx(before, _document_xml(tracked_insertion_at=2), _TRACKING_ON)

    plan = plan_canadian_to_csi(
        _document_xml(tracked_insertion_at=2),
        _styles_xml(),
        {
            "classifications": [
                {"paragraph_index": index, "csi_role": role}
                for index, (role, _text) in enumerate(_BODY)
            ]
        },
        numbering_xml=_numbering_xml(),
        settings_xml=_TRACKING_ON,
        revision_date="2026-01-01T00:00:00Z",
    )
    after = tmp_path / "after.docx"
    _write_docx(after, plan.document_xml, _TRACKING_ON)
    return before, after


def _insertions_by_author(path: Path, author: str) -> int:
    with zipfile.ZipFile(path) as z:
        blob = z.read("word/document.xml").decode("utf-8")
    return len(re.findall(rf'<w:ins\b[^>]*w:author="{re.escape(author)}"', blob))


def test_tracked_markers_add_exactly_one_revision_each(converted_tracked) -> None:
    before, after = converted_tracked
    assert _insertions_by_author(before, "Specification Formatter") == 0
    assert _insertions_by_author(after, "Specification Formatter") == len(_BODY)
    # The reviewer's own pending edit is neither removed nor duplicated.
    assert _insertions_by_author(before, "Reviewer") == 1
    assert _insertions_by_author(after, "Reviewer") == 1


def test_rejecting_every_revision_restores_the_source_exactly(converted_tracked) -> None:
    """The strongest statement available: the conversion is undoable in Word.

    A marker that survives rejection is a permanent edit wearing a revision's
    clothing, and would be discovered only by a reader who trusted it.
    """

    before, after = converted_tracked
    assert _view(after, "reject") == _view(before, "reject")


def test_accepting_every_revision_yields_the_numbered_document(converted_tracked) -> None:
    _before, after = converted_tracked
    markers = [
        _MARKER.match(text).group(1)
        for text in _view(after, "accept")
        if _MARKER.match(text)
    ]
    assert markers == [
        "PART 1", "1.1", "A.", "1.", "2.", "1.2", "A.", "PART 2", "2.1", "A.",
    ]


def test_tracked_markers_do_not_disturb_the_package(converted_tracked) -> None:
    before, after = converted_tracked
    before_members, after_members = _members(before), _members(after)
    assert set(before_members) == set(after_members)
    changed = {n for n in before_members if before_members[n] != after_members[n]}
    assert changed == {"word/document.xml"}


# --- Revision ids in a document whose own ids are already high ------------
#
# The converter used to number its revisions from a fixed 900000 "without
# needing to scan". This source already uses that range: a bookmark, a comment
# and a reviewer's insertion in the body, and a reviewer's insertion inside a
# footnote, which only a scan of the other parts can see.

_HIGH_ID_AUTHOR = "Specification Formatter"
_HIGH_ID_DATE = 'w:date="2026-01-01T00:00:00Z"'


def _high_id_document_xml() -> str:
    paragraphs = []
    for index, (role, text) in enumerate(_BODY):
        run = f'<w:r><w:t xml:space="preserve">{text}</w:t></w:r>'
        if index == 0:
            run = (
                '<w:bookmarkStart w:id="900000" w:name="_Toc1"/>'
                f'{run}<w:bookmarkEnd w:id="900000"/>'
            )
        if index == 3:
            run = (
                '<w:commentRangeStart w:id="900001"/>'
                f"{run}<w:commentRangeEnd w:id=\"900001\"/>"
                '<w:r><w:commentReference w:id="900001"/></w:r>'
            )
        if index == 5:
            run += (
                f'<w:ins w:id="900002" w:author="Reviewer" {_HIGH_ID_DATE}>'
                '<w:r><w:t xml:space="preserve"> Added under review.</w:t></w:r></w:ins>'
            )
        paragraphs.append(
            f'<w:p><w:pPr><w:pStyle w:val="{_ROLE_STYLE[role]}"/></w:pPr>{run}</w:p>'
        )
    return f'<w:document xmlns:w="{W}"><w:body>{"".join(paragraphs)}</w:body></w:document>'


_HIGH_ID_COMMENTS = (
    f'<w:comments xmlns:w="{W}"><w:comment w:id="900001" w:author="Reviewer" '
    f'{_HIGH_ID_DATE}><w:p><w:r><w:t>Check this.</w:t></w:r></w:p></w:comment>'
    "</w:comments>"
)
# 900005 is where the old base put the third paragraph's property revision.
_HIGH_ID_FOOTNOTES = (
    f'<w:footnotes xmlns:w="{W}"><w:footnote w:id="1"><w:p>'
    f'<w:ins w:id="900005" w:author="Reviewer" {_HIGH_ID_DATE}>'
    "<w:r><w:t>A note under review.</w:t></w:r></w:ins></w:p></w:footnote>"
    "</w:footnotes>"
)


def _high_id_parts(document_xml: str) -> dict[str, str]:
    return {
        "[Content_Types].xml": _CONTENT_TYPES,
        "_rels/.rels": _ROOT_RELS,
        "word/document.xml": document_xml,
        "word/styles.xml": _styles_xml(),
        "word/numbering.xml": _numbering_xml(),
        "word/settings.xml": _TRACKING_ON,
        "word/comments.xml": _HIGH_ID_COMMENTS,
        "word/footnotes.xml": _HIGH_ID_FOOTNOTES,
    }


@pytest.fixture
def converted_high_ids(tmp_path: Path) -> tuple[Path, Path]:
    """Convert through the directory entry point, which sees every part."""

    parts = _high_id_parts(_high_id_document_xml())
    before = tmp_path / "before.docx"
    with zipfile.ZipFile(before, "w", zipfile.ZIP_DEFLATED) as z:
        for name, payload in parts.items():
            z.writestr(name, payload)

    extracted = tmp_path / "extracted"
    for name, payload in parts.items():
        path = extracted / name
        path.parent.mkdir(parents=True, exist_ok=True)
        path.write_bytes(payload.encode("utf-8"))
    apply_canadian_to_csi(
        extracted,
        {
            "classifications": [
                {"paragraph_index": index, "csi_role": role}
                for index, (role, _text) in enumerate(_BODY)
            ]
        },
        [],
    )
    after = tmp_path / "after.docx"
    with zipfile.ZipFile(after, "w", zipfile.ZIP_DEFLATED) as z:
        for name in parts:
            z.writestr(name, (extracted / name).read_bytes())
    return before, after


# Every element whose w:id is an annotation id, and the subset that are
# revisions. Written out here rather than borrowed, so the two lists can
# disagree with the engine's.
_PAIRED_ANNOTATIONS = {
    "bookmarkStart", "bookmarkEnd",
    "commentRangeStart", "commentRangeEnd", "commentReference", "comment",
    "moveFromRangeStart", "moveFromRangeEnd", "moveToRangeStart", "moveToRangeEnd",
    "customXmlInsRangeStart", "customXmlInsRangeEnd",
    "customXmlDelRangeStart", "customXmlDelRangeEnd",
    "customXmlMoveFromRangeStart", "customXmlMoveFromRangeEnd",
    "customXmlMoveToRangeStart", "customXmlMoveToRangeEnd",
}
_REVISION_ELEMENTS = {
    "ins", "del", "moveFrom", "moveTo", "cellIns", "cellDel", "cellMerge",
    "rPrChange", "pPrChange", "sectPrChange", "tblPrChange", "tblPrExChange",
    "tblGridChange", "trPrChange", "tcPrChange", "numberingChange",
}


def _annotations(path: Path) -> list[tuple[str, str, int, str | None]]:
    """(part, local name, id, author) for every annotation in every word part."""

    found = []
    with zipfile.ZipFile(path) as z:
        for name in z.namelist():
            if not name.startswith("word/") or not name.endswith(".xml"):
                continue
            for element in ET.fromstring(z.read(name)).iter():
                if not element.tag.startswith(f"{{{W}}}"):
                    continue
                local = element.tag.split("}", 1)[1]
                if local not in _PAIRED_ANNOTATIONS | _REVISION_ELEMENTS:
                    continue
                raw = element.get(_q("id"))
                if raw is None or not re.fullmatch(r"-?\d+", raw):
                    continue
                found.append((name, local, int(raw), element.get(_q("author"))))
    return found


def _allocated_ids(path: Path) -> list[int]:
    return [
        revision_id
        for _part, local, revision_id, author in _annotations(path)
        if local in _REVISION_ELEMENTS and author == _HIGH_ID_AUTHOR
    ]


def test_allocated_ids_collide_with_no_id_in_the_source(converted_high_ids) -> None:
    before, after = converted_high_ids
    source_ids = {revision_id for _p, _l, revision_id, _a in _annotations(before)}
    assert {900000, 900001, 900002, 900005} <= source_ids, "fixture must reach the old base"

    allocated = _allocated_ids(after)
    # One insertion and one property revision per converted paragraph.
    assert len(allocated) == 2 * len(_BODY)
    assert not set(allocated) & source_ids


def test_allocated_ids_are_distinct_from_each_other(converted_high_ids) -> None:
    _before, after = converted_high_ids
    allocated = _allocated_ids(after)
    assert allocated, "check must not pass vacuously"
    assert len(set(allocated)) == len(allocated)


def test_revision_ids_are_unique_among_revisions_in_the_output(converted_high_ids) -> None:
    """Unique among revision elements, and only among them.

    Valid OOXML repeats an id across a paired annotation -- a bookmark's start
    and end, a comment's range and reference and the comment itself -- so
    requiring every id in the package to differ would reject ordinary
    documents. Those pairs must still be intact.
    """

    before, after = converted_high_ids
    revision_ids = [
        revision_id
        for _part, local, revision_id, _author in _annotations(after)
        if local in _REVISION_ELEMENTS
    ]
    assert len(revision_ids) == 2 * len(_BODY) + 2
    assert len(set(revision_ids)) == len(revision_ids)

    def paired(path: Path) -> list[tuple[str, str, int]]:
        return sorted(
            (part, local, revision_id)
            for part, local, revision_id, _author in _annotations(path)
            if local in _PAIRED_ANNOTATIONS
        )

    assert paired(after) == paired(before)
    assert [local for _p, local, rid in paired(after) if rid == 900001].count("comment") == 1
