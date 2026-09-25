"""Tracked revisions: what the final gate counts, and the ids they share.

Three things read revisions, and they read them the same way from here:

* **The census.** The final gate counts every revision element in
  ``word/document.xml`` by author and kind, before and after, and holds the
  difference to exactly what the conversion predicted. The body check sees
  runs, not the revision a run sits in, and the run-property check sees a
  lost revision only where it changes a run path -- so without a census a
  reviewer's ``w:del`` turned into a ``w:ins``, a ``w:pPrChange`` dropped, or
  a reviewer's revision re-signed with this application's name all published.
* **Header and footer revisions.** Replacing the target's header set deletes
  its parts, and the architect's parts arrive with whatever pending changes
  they carry. Both are counted here, so neither happens silently.
* **The id space.** A revision's ``w:id`` shares one space with every other
  annotation -- bookmarks, comments, permission ranges, other revisions -- so
  the reverse converter allocates its own ids above the highest one the
  package already uses.

Counts are by element, namespace-aware, from a parsed part: a move is its
``w:moveFrom`` and ``w:moveTo`` *and* the range markers around them, a
paragraph-mark insertion is one ``w:ins`` in ``w:pPr/w:rPr``. What is counted
is never content: an author is a name the census compares but never reports,
and a kind is one of :data:`REVISION_KINDS`.
"""

from __future__ import annotations

import re
from collections import Counter
from pathlib import Path
from typing import Mapping, Optional, Tuple, Union

from .ooxml_namespaces import CT_NS
from .untrusted_xml import parse_untrusted_xml
from .xml_helpers import W_NS, iter_start_tags

#: Every WordprocessingML revision element (ECMA-376 Part 1, 17.13.5), by
#: local name: content revisions, the paired range markers around a move or a
#: custom-XML revision, table cell revisions, and the property-change family.
#: A closed set of identifiers, so a kind can go into a failure message.
REVISION_KINDS: Tuple[str, ...] = (
    "ins",
    "del",
    "moveFrom",
    "moveTo",
    "moveFromRangeStart",
    "moveFromRangeEnd",
    "moveToRangeStart",
    "moveToRangeEnd",
    "customXmlInsRangeStart",
    "customXmlInsRangeEnd",
    "customXmlDelRangeStart",
    "customXmlDelRangeEnd",
    "customXmlMoveFromRangeStart",
    "customXmlMoveFromRangeEnd",
    "customXmlMoveToRangeStart",
    "customXmlMoveToRangeEnd",
    "cellIns",
    "cellDel",
    "cellMerge",
    "rPrChange",
    "pPrChange",
    "sectPrChange",
    "tblPrChange",
    "tblPrExChange",
    "tblGridChange",
    "trPrChange",
    "tcPrChange",
    "numberingChange",
)

#: Elements whose ``w:id`` is an annotation id: the revisions above, plus
#: bookmarks, comment ranges, references and the comments themselves, and
#: permission ranges. Footnote and endnote ids are a separate space (a
#: ``w:footnoteReference`` names a note, not an annotation) and sdt ids are
#: not annotations at all, so neither is here.
ANNOTATION_ID_ELEMENTS: Tuple[str, ...] = REVISION_KINDS + (
    "bookmarkStart",
    "bookmarkEnd",
    "commentRangeStart",
    "commentRangeEnd",
    "commentReference",
    "comment",
    "permStart",
    "permEnd",
)

_REVISION_TAGS = {f"{{{W_NS}}}{kind}": kind for kind in REVISION_KINDS}
_ANNOTATION_TAGS = frozenset(f"{{{W_NS}}}{name}" for name in ANNOTATION_ID_ELEMENTS)
_REVISION_QUALIFIED_NAMES = {f"w:{kind}": kind for kind in REVISION_KINDS}
_AUTHOR = f"{{{W_NS}}}author"
_ID = f"{{{W_NS}}}id"
_INTEGER_RX = re.compile(r"[+-]?[0-9]+")

def revision_census(xml: Union[bytes, str], part_name: str) -> Counter:
    """Every revision element in one part, counted by ``(author, kind)``.

    ``author`` is ``None`` on an element that has none: a range end carries
    only its id.

    Parsed, so the count is namespace-aware and a comment, CDATA section or
    processing instruction that merely looks like a revision is not one. The
    part is untrusted: a DOCTYPE or a malformed payload fails with the part's
    name, and a revision nobody can count is never reported as none.
    """

    counts: Counter = Counter()
    for element in parse_untrusted_xml(xml, part_name).iter():
        kind = _REVISION_TAGS.get(element.tag) if isinstance(element.tag, str) else None
        if kind is not None:
            counts[(element.get(_AUTHOR), kind)] += 1
    return counts


def count_revisions(xml: Union[bytes, str], part_name: str) -> int:
    """How many revision elements one part carries."""

    return sum(revision_census(xml, part_name).values())


def max_annotation_id(xml: Union[bytes, str], part_name: str) -> Optional[int]:
    """The highest integer annotation id in one part, or ``None``.

    Only :data:`ANNOTATION_ID_ELEMENTS` count, and only an id that reads as an
    integer: a permission range may carry a string id, which occupies no
    number a revision could be given.
    """

    highest: Optional[int] = None
    for element in parse_untrusted_xml(xml, part_name).iter():
        if element.tag not in _ANNOTATION_TAGS:
            continue
        value = _integer(element.get(_ID))
        if value is not None and (highest is None or value > highest):
            highest = value
    return highest


def highest_annotation_id_in_package(extract_dir: Path) -> Optional[int]:
    """The highest annotation id in any WordprocessingML part of a package.

    Every XML part below ``word/`` is read -- the body, headers, footers,
    footnotes, endnotes and comments, and the styles, numbering and glossary
    parts, which can carry property revisions of their own -- rather than the
    parts a relationship happens to name, because an id is taken wherever it
    is written. A part is XML when its extension is ``.xml`` in any case (OPC
    part names are case-insensitive) or when ``[Content_Types].xml`` declares
    an XML content type for it, by an ``Override`` or by its extension's
    ``Default``. Relationship parts carry no annotations and are skipped.
    """

    root = Path(extract_dir)
    word = root / "word"
    if not word.is_dir():
        return None
    overrides, defaults = _declared_content_types(root)
    highest: Optional[int] = None
    for path in sorted(word.rglob("*")):
        relative = path.relative_to(root)
        if "_rels" in relative.parts or not path.is_file():
            continue
        name = relative.as_posix()
        extension = path.suffix[1:].casefold()
        content_type = overrides.get(name.casefold(), defaults.get(extension, ""))
        if extension != "xml" and not _is_xml_content_type(content_type):
            continue
        value = max_annotation_id(path.read_bytes(), name)
        if value is not None and (highest is None or value > highest):
            highest = value
    return highest


def _declared_content_types(root: Path) -> Tuple[dict, dict]:
    """``[Content_Types].xml`` as (override by part name, default by extension).

    Both keyed casefolded, as OPC compares them. An absent part declares
    nothing, which leaves the ``.xml`` extension rule on its own.
    """

    path = root / "[Content_Types].xml"
    if not path.is_file():
        return {}, {}
    types = parse_untrusted_xml(path.read_bytes(), "[Content_Types].xml")
    overrides = {
        node.get("PartName", "").lstrip("/").casefold(): node.get("ContentType", "")
        for node in types.findall(f"{{{CT_NS}}}Override")
    }
    defaults = {
        node.get("Extension", "").casefold(): node.get("ContentType", "")
        for node in types.findall(f"{{{CT_NS}}}Default")
    }
    return overrides, defaults


def _is_xml_content_type(content_type: str) -> bool:
    value = content_type.strip().casefold()
    return value.endswith("+xml") or value in ("application/xml", "text/xml")


def revision_kinds_by_author(xml_fragment: str, author: str) -> Counter:
    """The revision kinds one author's name carries in a fragment, lexically.

    For a paragraph block, which declares no namespaces of its own and so
    cannot be parsed alone. The ``w`` prefix is the engine's throughout, and
    it is the prefix this application writes its own revisions with.
    """

    counts: Counter = Counter()
    for name, attributes in iter_start_tags(xml_fragment):
        kind = _REVISION_QUALIFIED_NAMES.get(name)
        if kind is not None and attributes.get("w:author") == author:
            counts[kind] += 1
    return counts


def revision_census_differences(
    before: Mapping,
    after: Mapping,
    own_author: str,
    own_added: Mapping[str, int],
) -> Tuple[str, ...]:
    """How ``after`` departs from ``before`` plus the predicted additions.

    Every author other than ``own_author`` must keep exactly the revisions
    they had, kind by kind; ``own_author`` must have exactly what it had plus
    ``own_added``. Each difference is described by kind and count and by
    whether it is this application's, never by an author's name: a
    reviewer's name is personal data, and the failure is about the count.
    """

    authors = {author for author, _kind in before} | {author for author, _kind in after}
    differences = []
    for kind in REVISION_KINDS:
        own_expected = before.get((own_author, kind), 0) + own_added.get(kind, 0)
        own_found = after.get((own_author, kind), 0)
        if own_expected != own_found:
            differences.append(
                f"{kind} by this application: expected {own_expected}, found {own_found}"
            )
        for author in sorted(authors - {own_author}, key=lambda value: (value is None, value or "")):
            expected = before.get((author, kind), 0)
            found = after.get((author, kind), 0)
            if expected != found:
                differences.append(
                    f"{kind} by another author: expected {expected}, found {found}"
                )
    return tuple(differences)


def _integer(value: Optional[str]) -> Optional[int]:
    if value is None or _INTEGER_RX.fullmatch(value.strip()) is None:
        return None
    return int(value.strip())


__all__ = [
    "ANNOTATION_ID_ELEMENTS",
    "REVISION_KINDS",
    "count_revisions",
    "highest_annotation_id_in_package",
    "max_annotation_id",
    "revision_census",
    "revision_census_differences",
    "revision_kinds_by_author",
]
