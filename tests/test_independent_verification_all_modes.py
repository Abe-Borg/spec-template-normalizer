"""Independent verification of Format-only and both forward Canadian conversions.

**This module shares no code with the engine on purpose.** It imports exactly
one thing from ``spec_formatter``, the public entry point, and only to
*produce* the output under test; ``test_module_imports_nothing_from_the_engine``
at the bottom enforces that. Every package, relationship, paragraph and
revision below is read with the standard library, by code written here, so
these checks can disagree with the engine's own invariants -- which is the
point. A check that runs the engine's extraction on both sides of a transform
cannot see a loss that extraction cannot see.
``tests/test_conversion_verification.py`` does the same for
``canadian_to_csi``; this module covers the other three modes:

* ``format_only`` (an architect template, no numbering change);
* ``csi_to_canadian`` (an architect template with a Canadian list);
* ``csi_to_canadian_standalone`` (no template: the built-in CSC list).

**Every expectation is written out by hand, before the engine runs.** The text
each paragraph must read afterwards is spelled out in the tables below, with
its tabs, breaks, revisions, fields and references in the notation of
:func:`_render`. The package members each mode may touch, and the ones it is
predicted to touch for these fixtures, are written from the mode's name --
never from ``ApplicationPolicy``, and never read off the output. Nothing
here is computed from what the engine produced and then compared with itself.

What is checked, per mode:

1. **The package whitelist.** Every member outside the mode's remit is
   byte-identical, and the members the mode changes, adds and removes are
   exactly the predicted ones.
2. **Exact text.** Every paragraph -- including table cells and text boxes --
   reads exactly as predicted, deleted text, revision containers, tabs,
   breaks, fields, symbols, bookmarks and references included. Paragraphs the
   mode does not change are predicted to read as they did.
3. **The revision census.** Every revision element in the main document,
   counted by author and kind, is unchanged; none of these modes writes one.
4. **Revision ids.** Unique among revision elements across every output
   part, and every paired annotation keeps its id.
5. **Even/odd header parity.** With the architect's header set imported, the
   output's ``w:evenAndOddHeaders`` reads as the architect's; with the
   target's own set kept, the target's settings part is byte-identical. The
   settings part is found through the document's relationship, not by name.
6. **Namespaces.** Every XML part of the output -- by extension or by the
   content type ``[Content_Types].xml`` declares for it -- parses, and every
   prefix a markup-compatibility attribute names is declared where it is used.

And beyond those: a paragraph the mode does not edit is element-identical,
properties and attributes included, not merely the same text; the imported
headers and footers read as the architect's; both inputs are byte-identical
after the run. Finally, every check is run against the real output damaged
the way a defect would damage it, and must reject that -- a check that has
never failed proves nothing.
"""

from __future__ import annotations

import ast
import codecs
import hashlib
import io
import posixpath
import re
import sys
import zipfile
import xml.etree.ElementTree as ET
from collections import Counter
from pathlib import Path

import pytest

from spec_formatter import format_specifications

# --- Namespaces -------------------------------------------------------------

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
V = "urn:schemas-microsoft-com:vml"
A = "http://schemas.openxmlformats.org/drawingml/2006/main"
W14 = "http://schemas.microsoft.com/office/word/2010/wordml"
MC = "http://schemas.openxmlformats.org/markup-compatibility/2006"
PKG_REL = "http://schemas.openxmlformats.org/package/2006/relationships"
CT = "http://schemas.openxmlformats.org/package/2006/content-types"
XML_NS = "http://www.w3.org/XML/1998/namespace"
REL_TYPE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"

FORMAT_ONLY = "format_only"
CSI_TO_CANADIAN = "csi_to_canadian"
CSI_TO_CANADIAN_STANDALONE = "csi_to_canadian_standalone"
ALL_MODES = (FORMAT_ONLY, CSI_TO_CANADIAN, CSI_TO_CANADIAN_STANDALONE)


def _w(local: str) -> str:
    return f"{{{W}}}{local}"


# --- The fixtures -----------------------------------------------------------
#
# Written out as literal XML. The target bodies carry every kind of run
# content a whitespace-normalized text check cannot see, pending revisions by
# a reviewer, paired annotations, a table, a text box and a VML shape -- in
# paragraphs each mode leaves alone as well as in the ones it changes.

REVIEWER = 'w:author="Reviewer" w:date="2026-01-02T03:04:05Z"'

DOCUMENT_ROOT = (
    f'<w:document xmlns:w="{W}" xmlns:r="{R}" xmlns:v="{V}" '
    f'xmlns:w14="{W14}" xmlns:mc="{MC}" mc:Ignorable="w14">'
)

#: An editorial note: pStyle ``CMT`` is ignored in every mode, so this
#: paragraph must come through exactly. It carries a paragraph-mark insertion,
#: a paragraph-property revision, a reviewer's insertion and deletion, a
#: run-property revision, a field, a symbol, a positional tab, a carriage
#: return, and the characters a normalized reading loses.
NOTE = (
    '<w:p w14:paraId="1A2B3C4D" w14:textId="77777777"><w:pPr><w:pStyle w:val="CMT"/>'
    f'<w:rPr><w:ins w:id="17" {REVIEWER}/></w:rPr>'
    f'<w:pPrChange w:id="14" {REVIEWER}><w:pPr/></w:pPrChange></w:pPr>'
    '<w:r><w:t xml:space="preserve">Retain  this note </w:t><w:tab/><w:t>for</w:t>'
    '<w:br/><w:t xml:space="preserve"> the</w:t><w:t>\u00a0designer</w:t>'
    "<w:softHyphen/><w:t>of</w:t><w:noBreakHyphen/><w:t>record</w:t><w:cr/>"
    '<w:ptab w:relativeTo="margin" w:alignment="right" w:leader="none"/></w:r>'
    f'<w:ins w:id="11" {REVIEWER}><w:r><w:t xml:space="preserve"> now</w:t></w:r></w:ins>'
    f'<w:del w:id="12" {REVIEWER}><w:r><w:delText xml:space="preserve"> later</w:delText>'
    "</w:r></w:del>"
    f'<w:r><w:rPr><w:b/><w:rPrChange w:id="13" {REVIEWER}><w:rPr/></w:rPrChange></w:rPr>'
    '<w:t xml:space="preserve"> see </w:t></w:r>'
    '<w:r><w:fldChar w:fldCharType="begin"/></w:r>'
    '<w:r><w:instrText xml:space="preserve"> PAGE </w:instrText></w:r>'
    '<w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>7</w:t></w:r>'
    '<w:r><w:fldChar w:fldCharType="end"/></w:r>'
    '<w:r><w:sym w:font="Symbol" w:char="F0B0"/></w:r></w:p>'
)
NOTE_TEXT = (
    "Retain  this note \tfor\n the\u00a0designer\u00adof\u2011record\r[ptab]"
    "{+ now+}{- later-} see [fld begin][instr  PAGE ][fld separate]7[fld end]"
    "[sym Symbol F0B0]"
)

#: A table: out of scope in every mode. The second cell starts with a typed
#: marker that no mode may treat as a list item. A reviewer's tracked move
#: runs from the first cell to the second, each end inside its named range.
TABLE = (
    '<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr>'
    '<w:tblGrid><w:gridCol w:w="4000"/><w:gridCol w:w="4000"/></w:tblGrid><w:tr>'
    '<w:tc><w:tcPr><w:tcW w:w="4000" w:type="dxa"/></w:tcPr><w:p>'
    '<w:r><w:t xml:space="preserve">Pipe </w:t></w:r>'
    f'<w:ins w:id="15" {REVIEWER}><w:r><w:t>size</w:t></w:r></w:ins>'
    f'<w:del w:id="16" {REVIEWER}><w:r><w:delText>diameter</w:delText></w:r></w:del>'
    f'<w:moveFromRangeStart w:id="41" {REVIEWER} w:name="move1"/>'
    f'<w:moveFrom w:id="42" {REVIEWER}><w:r><w:t xml:space="preserve"> schedule</w:t></w:r>'
    '</w:moveFrom><w:moveFromRangeEnd w:id="41"/>'
    "</w:p></w:tc>"
    '<w:tc><w:tcPr><w:tcW w:w="4000" w:type="dxa"/></w:tcPr><w:p>'
    "<w:r><w:t>A.</w:t><w:tab/><w:t>Not a list item</w:t></w:r>"
    f'<w:moveToRangeStart w:id="43" {REVIEWER} w:name="move1"/>'
    f'<w:moveTo w:id="44" {REVIEWER}><w:r><w:t xml:space="preserve"> schedule</w:t></w:r>'
    '</w:moveTo><w:moveToRangeEnd w:id="43"/>'
    "</w:p></w:tc></w:tr></w:tbl>"
)
TABLE_TEXTS = (
    "Pipe {+size+}{-diameter-}[moveFromRangeStart 41]{< schedule<}[moveFromRangeEnd 41]",
    "A.\tNot a list item[moveToRangeStart 43]{> schedule>}[moveToRangeEnd 43]",
)

#: A paragraph holding a VML text box, and the text box's own paragraph.
TEXT_BOX = (
    '<w:p><w:r><w:t xml:space="preserve">Figure </w:t></w:r><w:r><w:pict>'
    '<v:shape id="tb1" style="width:100pt;height:40pt"><v:textbox><w:txbxContent>'
    "<w:p><w:r><w:t>Text box text</w:t></w:r></w:p>"
    "</w:txbxContent></v:textbox></v:shape></w:pict></w:r></w:p>"
)
TEXT_BOX_TEXTS = ("Figure [drawing]", "Text box text")

EMPTY = "<w:p/>"

#: The content no mode may change, in document order: the note, the table's
#: two cells, an empty paragraph, the text box host and the text box's own
#: paragraph.
UNTOUCHED_BLOCK = NOTE + TABLE + EMPTY + TEXT_BOX
UNTOUCHED_TEXTS = (NOTE_TEXT, *TABLE_TEXTS, "", *TEXT_BOX_TEXTS)

# The typed CSI list items the architect modes classify. ``A.`` sits in a run
# of its own with the list tab in the next run; ``1.`` shares one text node
# with its text; ``2.`` has the tab inside its own run; ``B.`` ends its node
# with the separating space. A paragraph a converter changes may hold no
# tracked change or field (the conversion refuses one), so those live in the
# note and the table; Format-only's ``B.`` carries one of its own.
ITEM_A = (
    '<w:p w14:paraId="2B3C4D5E" w14:textId="77777777"><w:pPr><w:jc w:val="left"/></w:pPr>'
    "<w:r><w:t>A.</w:t></w:r><w:r><w:tab/></w:r>"
    '<w:r><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial"/></w:rPr>'
    '<w:t xml:space="preserve">Provide  wet-pipe sprinklers </w:t><w:tab/>'
    "<w:t>throughout\u00a0the</w:t><w:softHyphen/><w:t>building.</w:t></w:r>"
    '<w:bookmarkStart w:id="31" w:name="_Ref1"/>'
    '<w:r><w:t xml:space="preserve"> Coordinate</w:t></w:r><w:bookmarkEnd w:id="31"/></w:p>'
)
ITEM_1 = (
    '<w:p><w:r><w:t xml:space="preserve">1. Hydraulic calculations per NFPA 13.</w:t>'
    "<w:br/><w:t>Submit before installation.</w:t></w:r></w:p>"
)
ITEM_2 = (
    '<w:p><w:commentRangeStart w:id="0"/><w:r><w:t>2.</w:t><w:tab/>'
    '<w:t xml:space="preserve">Sprinkler heads: listed</w:t><w:noBreakHyphen/>'
    "<w:t>quick</w:t><w:softHyphen/><w:t>response</w:t></w:r>"
    '<w:r><w:footnoteReference w:id="1"/></w:r><w:commentRangeEnd w:id="0"/>'
    '<w:r><w:commentReference w:id="0"/></w:r></w:p>'
)


def _item_b(*, tracked: bool) -> str:
    insertion = (
        f'<w:ins w:id="18" {REVIEWER}><w:r><w:t xml:space="preserve"> promptly</w:t>'
        "</w:r></w:ins>"
        if tracked
        else ""
    )
    return (
        '<w:p><w:r><w:t xml:space="preserve">B. </w:t></w:r>'
        '<w:hyperlink w:anchor="_Ref1"><w:r><w:t>Related</w:t></w:r></w:hyperlink>'
        f'<w:r><w:t xml:space="preserve"> requirements  apply.</w:t></w:r>{insertion}</w:p>'
    )


TARGET_SECTION = (
    '<w:sectPr><w:headerReference w:type="default" r:id="rIdTargetHeader"/>'
    '<w:pgSz w:w="12240" w:h="15840"/>'
    '<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" '
    'w:header="720" w:footer="720" w:gutter="0"/></w:sectPr>'
)


def _architect_mode_target_body(*, tracked_item: bool) -> str:
    return (
        "<w:p><w:r><w:t>SECTION 21 13 13</w:t></w:r></w:p>"
        "<w:p><w:r><w:t>WET-PIPE SPRINKLER SYSTEMS</w:t></w:r></w:p>"
        + NOTE
        + ITEM_A
        + ITEM_1
        + ITEM_2
        + _item_b(tracked=tracked_item)
        + TABLE
        + EMPTY
        + TEXT_BOX
        + "<w:p><w:r><w:t>END OF SECTION 21 13 13</w:t></w:r></w:p>"
    )


# The source text of every paragraph of the architect modes' target, and what
# each mode must leave. Format-only changes no text at all. The Canadian
# conversion removes each typed marker with the separator after it -- and,
# where the marker's text node ends at the marker, the first run-level tab
# before the next text node -- and nothing else.
ARCHITECT_MODE_SOURCE_TEXTS = (
    "SECTION 21 13 13",
    "WET-PIPE SPRINKLER SYSTEMS",
    NOTE_TEXT,
    "A.\tProvide  wet-pipe sprinklers \tthroughout\u00a0the\u00adbuilding."
    "[bm 31 _Ref1] Coordinate[/bm 31]",
    "1. Hydraulic calculations per NFPA 13.\nSubmit before installation.",
    "[cr 0]2.\tSprinkler heads: listed\u2011quick\u00adresponse[footnote 1][/cr 0]"
    "[comment 0]",
    "B. [link]Related[/link] requirements  apply.",
    *TABLE_TEXTS,
    "",
    *TEXT_BOX_TEXTS,
    "END OF SECTION 21 13 13",
)
FORMAT_ONLY_SOURCE_TEXTS = (
    *ARCHITECT_MODE_SOURCE_TEXTS[:6],
    "B. [link]Related[/link] requirements  apply.{+ promptly+}",
    *ARCHITECT_MODE_SOURCE_TEXTS[7:],
)
FORMAT_ONLY_EXPECTED_TEXTS = FORMAT_ONLY_SOURCE_TEXTS
CSI_TO_CANADIAN_EXPECTED_TEXTS = (
    "SECTION 21 13 13",
    "WET-PIPE SPRINKLER SYSTEMS",
    NOTE_TEXT,
    "Provide  wet-pipe sprinklers \tthroughout\u00a0the\u00adbuilding."
    "[bm 31 _Ref1] Coordinate[/bm 31]",
    "Hydraulic calculations per NFPA 13.\nSubmit before installation.",
    "[cr 0]Sprinkler heads: listed\u2011quick\u00adresponse[footnote 1][/cr 0]"
    "[comment 0]",
    "[link]Related[/link] requirements  apply.",
    *TABLE_TEXTS,
    "",
    *TEXT_BOX_TEXTS,
    "END OF SECTION 21 13 13",
)
#: The paragraphs the Canadian conversion is predicted to change, by index.
CSI_TO_CANADIAN_CHANGED = (3, 4, 5, 6)


def _standalone_target_body() -> str:
    return (
        "<w:p><w:r><w:t>SECTION 21 13 13</w:t></w:r></w:p>"
        "<w:p><w:r><w:t>WET-PIPE SPRINKLER SYSTEMS</w:t></w:r></w:p>"
        "<w:p><w:r><w:t>PART 1</w:t></w:r><w:r><w:tab/></w:r>"
        "<w:r><w:t>GENERAL</w:t></w:r></w:p>"
        "<w:p><w:r><w:t>1.01</w:t><w:tab/><w:t>SUMMARY</w:t></w:r></w:p>"
        + ITEM_A
        + NOTE
        + ITEM_1
        + ITEM_2
        + _item_b(tracked=False)
        + '<w:p><w:r><w:t xml:space="preserve">1.02 REFERENCES</w:t></w:r></w:p>'
        + "<w:p><w:r><w:t>A.</w:t><w:tab/>"
        '<w:t xml:space="preserve">NFPA 13 governs  installation.</w:t></w:r></w:p>'
        + TABLE
        + EMPTY
        + TEXT_BOX
        + '<w:p><w:r><w:t xml:space="preserve">PART 2 - PRODUCTS</w:t></w:r></w:p>'
        + "<w:p><w:r><w:t>2.01</w:t></w:r><w:r><w:tab/><w:t>SPRINKLERS</w:t></w:r></w:p>"
        + "<w:p><w:r><w:t>A.</w:t><w:tab/><w:t>Provide listed sprinklers.</w:t></w:r></w:p>"
        + "<w:p><w:r><w:t>END OF SECTION 21 13 13</w:t></w:r></w:p>"
    )


STANDALONE_SOURCE_TEXTS = (
    "SECTION 21 13 13",
    "WET-PIPE SPRINKLER SYSTEMS",
    "PART 1\tGENERAL",
    "1.01\tSUMMARY",
    ARCHITECT_MODE_SOURCE_TEXTS[3],
    NOTE_TEXT,
    ARCHITECT_MODE_SOURCE_TEXTS[4],
    ARCHITECT_MODE_SOURCE_TEXTS[5],
    ARCHITECT_MODE_SOURCE_TEXTS[6],
    "1.02 REFERENCES",
    "A.\tNFPA 13 governs  installation.",
    *TABLE_TEXTS,
    "",
    *TEXT_BOX_TEXTS,
    "PART 2 - PRODUCTS",
    "2.01\tSPRINKLERS",
    "A.\tProvide listed sprinklers.",
    "END OF SECTION 21 13 13",
)
STANDALONE_EXPECTED_TEXTS = (
    "SECTION 21 13 13",
    "WET-PIPE SPRINKLER SYSTEMS",
    "GENERAL",
    "SUMMARY",
    CSI_TO_CANADIAN_EXPECTED_TEXTS[3],
    NOTE_TEXT,
    CSI_TO_CANADIAN_EXPECTED_TEXTS[4],
    CSI_TO_CANADIAN_EXPECTED_TEXTS[5],
    CSI_TO_CANADIAN_EXPECTED_TEXTS[6],
    "REFERENCES",
    "NFPA 13 governs  installation.",
    *TABLE_TEXTS,
    "",
    *TEXT_BOX_TEXTS,
    "PRODUCTS",
    "SPRINKLERS",
    "Provide listed sprinklers.",
    "END OF SECTION 21 13 13",
)
STANDALONE_CHANGED = (2, 3, 4, 6, 7, 8, 9, 10, 16, 17, 18)

#: Every revision element in the targets' main documents, by (author, kind).
#: No mode covered here writes or removes one.
ARCHITECT_MODE_REVISIONS = {
    ("Reviewer", "ins"): 3,
    ("Reviewer", "del"): 2,
    ("Reviewer", "rPrChange"): 1,
    ("Reviewer", "pPrChange"): 1,
    ("Reviewer", "moveFrom"): 1,
    ("Reviewer", "moveTo"): 1,
    ("Reviewer", "moveFromRangeStart"): 1,
    ("Reviewer", "moveToRangeStart"): 1,
    # A range's end carries only the id it shares with its start.
    (None, "moveFromRangeEnd"): 1,
    (None, "moveToRangeEnd"): 1,
}
FORMAT_ONLY_REVISIONS = {**ARCHITECT_MODE_REVISIONS, ("Reviewer", "ins"): 4}

# --- Package parts ------------------------------------------------------------

TARGET_STYLES = (
    f'<w:styles xmlns:w="{W}">'
    "<w:docDefaults><w:rPrDefault><w:rPr/></w:rPrDefault>"
    "<w:pPrDefault><w:pPr/></w:pPrDefault></w:docDefaults>"
    '<w:style w:type="paragraph" w:default="1" w:styleId="Normal">'
    '<w:name w:val="Normal"/><w:qFormat/></w:style>'
    '<w:style w:type="paragraph" w:styleId="CMT"><w:name w:val="CMT"/>'
    '<w:basedOn w:val="Normal"/><w:rPr><w:color w:val="0000FF"/></w:rPr></w:style>'
    "</w:styles>"
)
#: A list the target defines and never uses: Format-only must leave it alone.
TARGET_NUMBERING = (
    f'<w:numbering xmlns:w="{W}"><w:abstractNum w:abstractNumId="42">'
    '<w:multiLevelType w:val="multilevel"/><w:lvl w:ilvl="0"><w:start w:val="1"/>'
    '<w:numFmt w:val="lowerRoman"/><w:lvlText w:val="(%1)"/></w:lvl></w:abstractNum>'
    '<w:num w:numId="17"><w:abstractNumId w:val="42"/></w:num></w:numbering>'
)
TARGET_HEADER = (
    f'<w:hdr xmlns:w="{W}"><w:p><w:r><w:t>Target header</w:t></w:r>'
    f'<w:ins w:id="19" {REVIEWER}><w:r><w:t xml:space="preserve"> draft</w:t></w:r></w:ins>'
    "</w:p></w:hdr>"
)
TARGET_FOOTNOTES = (
    f'<w:footnotes xmlns:w="{W}"><w:footnote w:id="1"><w:p>'
    f'<w:ins w:id="21" {REVIEWER}><w:r><w:t>A footnote under review.</w:t></w:r></w:ins>'
    "</w:p></w:footnote></w:footnotes>"
)
TARGET_COMMENTS = (
    f'<w:comments xmlns:w="{W}"><w:comment w:id="0" {REVIEWER} w:initials="R">'
    "<w:p><w:r><w:t>Confirm the listing.</w:t></w:r></w:p></w:comment></w:comments>"
)
TARGET_THEME = (
    f'<a:theme xmlns:a="{A}" name="Target Theme"><a:themeElements>'
    '<a:clrScheme name="Target"/><a:fontScheme name="Target"/><a:fmtScheme name="Target"/>'
    "</a:themeElements></a:theme>"
)
TARGET_FONTS = (
    f'<w:fonts xmlns:w="{W}"><w:font w:name="Arial"><w:family w:val="swiss"/></w:font>'
    "</w:fonts>"
)
#: An XML part whose extension says nothing: only ``[Content_Types].xml``
#: declares it XML.
CUSTOM_XML_DATA = b'<?xml version="1.0" encoding="UTF-8"?><data xmlns="urn:example:wi08"><item>1</item></data>'
TRASH_ITEM = b"\x00a discarded physical ZIP item\x00"


def _utf16(text: str, *, little_endian: bool) -> bytes:
    declared = f'<?xml version="1.0" encoding="UTF-16"?>{text}'
    if little_endian:
        return codecs.BOM_UTF16_LE + declared.encode("utf-16-le")
    return codecs.BOM_UTF16_BE + declared.encode("utf-16-be")


def _target_settings(*, even_and_odd: bool, utf16: bool) -> bytes:
    switch = "<w:evenAndOddHeaders/>" if even_and_odd else ""
    text = (
        f'<w:settings xmlns:w="{W}"><w:zoom w:percent="110"/>'
        f'<w:defaultTabStop w:val="720"/>{switch}</w:settings>'
    )
    if utf16:
        return _utf16(text, little_endian=False)
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' + text).encode("utf-8")


_MAIN = "application/vnd.openxmlformats-officedocument.wordprocessingml"


def _content_types(overrides: dict[str, str]) -> str:
    return (
        f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="{CT}">'
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
        '<Default Extension="xml" ContentType="application/xml"/>'
        '<Default Extension="png" ContentType="image/png"/>'
        + "".join(
            f'<Override PartName="/{name}" ContentType="{content_type}"/>'
            for name, content_type in overrides.items()
        )
        + "</Types>"
    )


def _relationships(relationships: list[tuple[str, str, str]]) -> str:
    return (
        f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="{PKG_REL}">'
        + "".join(
            f'<Relationship Id="{rid}" Type="{REL_TYPE}/{kind}" Target="{target}"/>'
            for rid, kind, target in relationships
        )
        + "</Relationships>"
    )


ROOT_RELS = _relationships([("rId1", "officeDocument", "word/document.xml")])


def _write(path: Path, parts: dict[str, bytes | str]) -> Path:
    path.parent.mkdir(parents=True, exist_ok=True)
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as package:
        for name, payload in parts.items():
            package.writestr(name, payload)
    return path


def _write_target(path: Path, body: str, *, numbering: bool, settings: bytes) -> Path:
    overrides = {
        "word/document.xml": f"{_MAIN}.document.main+xml",
        "word/styles.xml": f"{_MAIN}.styles+xml",
        "word/settings.xml": f"{_MAIN}.settings+xml",
        "word/header9.xml": f"{_MAIN}.header+xml",
        "word/footnotes.xml": f"{_MAIN}.footnotes+xml",
        "word/comments.xml": f"{_MAIN}.comments+xml",
        "word/theme/theme1.xml": "application/vnd.openxmlformats-officedocument.theme+xml",
        "word/fontTable.xml": f"{_MAIN}.fontTable+xml",
        "customXml/item2.data": "application/xml",
    }
    relationships = [
        ("rIdStyles", "styles", "styles.xml"),
        ("rIdSettings", "settings", "settings.xml"),
        ("rIdTargetHeader", "header", "header9.xml"),
        ("rIdFootnotes", "footnotes", "footnotes.xml"),
        ("rIdComments", "comments", "comments.xml"),
        ("rIdTheme", "theme", "theme/theme1.xml"),
        ("rIdFonts", "fontTable", "fontTable.xml"),
    ]
    parts: dict[str, bytes | str] = {
        "[Content_Types].xml": "",
        "_rels/.rels": ROOT_RELS,
        "word/document.xml": (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            f"{DOCUMENT_ROOT}<w:body>{body}{TARGET_SECTION}</w:body></w:document>"
        ),
        "word/_rels/document.xml.rels": "",
        "word/styles.xml": TARGET_STYLES,
        "word/settings.xml": settings,
        "word/header9.xml": TARGET_HEADER,
        "word/footnotes.xml": TARGET_FOOTNOTES,
        "word/comments.xml": TARGET_COMMENTS,
        "word/theme/theme1.xml": TARGET_THEME,
        "word/fontTable.xml": TARGET_FONTS,
        "customXml/item2.data": CUSTOM_XML_DATA,
        "[trash]/0000.dat": TRASH_ITEM,
    }
    if numbering:
        overrides["word/numbering.xml"] = f"{_MAIN}.numbering+xml"
        relationships.append(("rIdNumbering", "numbering", "numbering.xml"))
        parts["word/numbering.xml"] = TARGET_NUMBERING
    parts["[Content_Types].xml"] = _content_types(overrides)
    parts["word/_rels/document.xml.rels"] = _relationships(relationships)
    return _write(path, parts)


# The architect. Its two numbered paragraphs are the exemplars of the
# PARAGRAPH and SUBPARAGRAPH roles; its section references a default, a first
# and an even header and a default footer. Its stylesheet is shaped the way
# current Word writes one, with ligatures in the document defaults.

PNG_1X1 = bytes.fromhex(
    "89504e470d0a1a0a0000000d49484452000000010000000108060000001f15c489"
    "0000000d49444154789c6360f8cfc00000040101005f8d3a0000000049454e44ae426082"
)

ARCHITECT_STYLES = (
    f'<w:styles xmlns:w="{W}" xmlns:w14="{W14}" xmlns:mc="{MC}" mc:Ignorable="w14">'
    "<w:docDefaults><w:rPrDefault><w:rPr>"
    '<w:rFonts w:ascii="Aptos" w:hAnsi="Aptos"/><w:sz w:val="22"/>'
    '<w14:ligatures w14:val="standardContextual"/>'
    "</w:rPr></w:rPrDefault><w:pPrDefault><w:pPr/></w:pPrDefault></w:docDefaults>"
    '<w:style w:type="paragraph" w:default="1" w:styleId="Normal">'
    '<w:name w:val="Normal"/><w:qFormat/>'
    '<w:rPr><w14:ligatures w14:val="standard"/></w:rPr></w:style>'
    "</w:styles>"
)
ARCHITECT_THEME = (
    f'<a:theme xmlns:a="{A}" name="Architect Theme"><a:themeElements>'
    '<a:clrScheme name="Architect"/><a:fontScheme name="Architect"/>'
    '<a:fmtScheme name="Architect"/></a:themeElements></a:theme>'
)
ARCHITECT_FONTS = (
    f'<w:fonts xmlns:w="{W}"><w:font w:name="Aptos"><w:family w:val="swiss"/></w:font>'
    "</w:fonts>"
)
ARCHITECT_HEADERS = {
    "word/header1.xml": (
        f'<w:hdr xmlns:w="{W}" xmlns:r="{R}" xmlns:a="{A}">'
        '<w:p><w:r><w:drawing><a:blip r:embed="rIdImage"/></w:drawing></w:r>'
        '<w:hyperlink r:id="rIdLink"><w:r><w:t>Architect link</w:t></w:r></w:hyperlink>'
        "</w:p></w:hdr>"
    ),
    "word/_rels/header1.xml.rels": (
        f'<Relationships xmlns="{PKG_REL}">'
        f'<Relationship Id="rIdImage" Type="{REL_TYPE}/image" Target="media/logo.png"/>'
        f'<Relationship Id="rIdLink" Type="{REL_TYPE}/hyperlink" '
        'Target="https://example.com/spec" TargetMode="External"/></Relationships>'
    ),
    "word/header2.xml": f'<w:hdr xmlns:w="{W}"><w:p><w:r><w:t>First header</w:t></w:r></w:p></w:hdr>',
    "word/header3.xml": f'<w:hdr xmlns:w="{W}"><w:p><w:r><w:t>Even header</w:t></w:r></w:p></w:hdr>',
    "word/footer1.xml": f'<w:ftr xmlns:w="{W}"><w:p><w:r><w:t>Default footer</w:t></w:r></w:p></w:ftr>',
    "word/media/logo.png": PNG_1X1,
}
#: What each imported header or footer must read, by where the output's
#: section references it.
ARCHITECT_HEADER_TEXTS = {
    ("header", "default"): ("[drawing][link]Architect link[/link]",),
    ("header", "first"): ("First header",),
    ("header", "even"): ("Even header",),
    ("footer", "default"): ("Default footer",),
}


def _architect_numbering(*, canadian: bool) -> str:
    if canadian:
        levels = (
            '<w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/>'
            '<w:lvlText w:val="PART %1"/></w:lvl>'
            '<w:lvl w:ilvl="1"><w:start w:val="1"/><w:numFmt w:val="decimal"/>'
            '<w:lvlText w:val="%1.%2"/></w:lvl>'
            '<w:lvl w:ilvl="2"><w:start w:val="1"/><w:numFmt w:val="decimal"/>'
            '<w:lvlText w:val=".%3"/><w:pPr><w:ind w:left="720" w:hanging="360"/></w:pPr></w:lvl>'
            '<w:lvl w:ilvl="3"><w:start w:val="1"/><w:numFmt w:val="decimal"/>'
            '<w:lvlText w:val=".%4"/><w:pPr><w:ind w:left="1440" w:hanging="360"/></w:pPr></w:lvl>'
        )
    else:
        levels = (
            '<w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="upperLetter"/>'
            '<w:lvlText w:val="%1."/><w:pPr><w:ind w:left="720" w:hanging="360"/></w:pPr></w:lvl>'
            '<w:lvl w:ilvl="1"><w:start w:val="1"/><w:numFmt w:val="decimal"/>'
            '<w:lvlText w:val="%2."/><w:pPr><w:ind w:left="1440" w:hanging="360"/></w:pPr></w:lvl>'
        )
    return (
        f'<w:numbering xmlns:w="{W}"><w:abstractNum w:abstractNumId="5">'
        f'<w:multiLevelType w:val="multilevel"/>{levels}</w:abstractNum>'
        '<w:num w:numId="5"><w:abstractNumId w:val="5"/></w:num></w:numbering>'
    )


def _write_architect(path: Path, *, canadian: bool, even_and_odd: bool) -> Path:
    exemplars = (
        ((0, "GENERAL"), (1, "SUMMARY"),
         (2, "Architect paragraph one"), (3, "Architect paragraph two"))
        if canadian else
        ((0, "Architect paragraph one"), (1, "Architect paragraph two"))
    )
    body = "".join(
        f'<w:p><w:pPr><w:numPr><w:ilvl w:val="{ilvl}"/><w:numId w:val="5"/></w:numPr></w:pPr>'
        f"<w:r><w:t>{text}</w:t></w:r></w:p>"
        for ilvl, text in exemplars
    )
    section = (
        '<w:sectPr><w:headerReference w:type="default" r:id="rIdHdrDefault"/>'
        '<w:headerReference w:type="first" r:id="rIdHdrFirst"/>'
        '<w:headerReference w:type="even" r:id="rIdHdrEven"/>'
        '<w:footerReference w:type="default" r:id="rIdFtrDefault"/>'
        '<w:pgSz w:w="12240" w:h="15840"/>'
        '<w:pgMar w:top="1080" w:right="1080" w:bottom="1080" w:left="1080" '
        'w:header="540" w:footer="540" w:gutter="0"/><w:titlePg/></w:sectPr>'
    )
    switch = "<w:evenAndOddHeaders/>" if even_and_odd else ""
    settings = _utf16(
        f'<w:settings xmlns:w="{W}"><w:zoom w:percent="95"/>{switch}<w:compat>'
        '<w:compatSetting w:name="compatibilityMode" '
        'w:uri="http://schemas.microsoft.com/office/word" w:val="15"/>'
        "</w:compat></w:settings>",
        little_endian=True,
    )
    overrides = {
        "word/document.xml": f"{_MAIN}.document.main+xml",
        "word/styles.xml": f"{_MAIN}.styles+xml",
        "word/settings.xml": f"{_MAIN}.settings+xml",
        "word/numbering.xml": f"{_MAIN}.numbering+xml",
        "word/theme/theme1.xml": "application/vnd.openxmlformats-officedocument.theme+xml",
        "word/fontTable.xml": f"{_MAIN}.fontTable+xml",
        "word/header1.xml": f"{_MAIN}.header+xml",
        "word/header2.xml": f"{_MAIN}.header+xml",
        "word/header3.xml": f"{_MAIN}.header+xml",
        "word/footer1.xml": f"{_MAIN}.footer+xml",
    }
    relationships = [
        ("rIdStyles", "styles", "styles.xml"),
        ("rIdSettings", "settings", "settings.xml"),
        ("rIdNumbering", "numbering", "numbering.xml"),
        ("rIdTheme", "theme", "theme/theme1.xml"),
        ("rIdFonts", "fontTable", "fontTable.xml"),
        ("rIdHdrDefault", "header", "header1.xml"),
        ("rIdHdrFirst", "header", "header2.xml"),
        ("rIdHdrEven", "header", "header3.xml"),
        ("rIdFtrDefault", "footer", "footer1.xml"),
    ]
    return _write(
        path,
        {
            "[Content_Types].xml": _content_types(overrides),
            "_rels/.rels": ROOT_RELS,
            "word/document.xml": (
                f'<w:document xmlns:w="{W}" xmlns:r="{R}"><w:body>{body}{section}'
                "</w:body></w:document>"
            ),
            "word/_rels/document.xml.rels": _relationships(relationships),
            "word/styles.xml": ARCHITECT_STYLES,
            "word/settings.xml": settings,
            "word/numbering.xml": _architect_numbering(canadian=canadian),
            "word/theme/theme1.xml": ARCHITECT_THEME,
            "word/fontTable.xml": ARCHITECT_FONTS,
            **ARCHITECT_HEADERS,
        },
    )


def _template_classifier(**kwargs):
    """Classify the list exemplars, including PART/ARTICLE for Canadian mode.

    Written here rather than borrowed from another test module, so this file
    depends on nothing but the public entry point.
    """

    classifiable = [
        item
        for item in kwargs["slim_bundle"].get("paragraphs", [])
        if item.get("skip_reason") is None
    ]
    assert len(classifiable) in {2, 4}, "the architect fixture has two or four exemplars"
    roles = [
        ("PARAGRAPH", "CSI_Paragraph__ARCH", "CSI Paragraph"),
        ("SUBPARAGRAPH", "CSI_Subparagraph__ARCH", "CSI Subparagraph"),
    ]
    if len(classifiable) == 4:
        roles = [
            ("PART", "CSI_Part__ARCH", "CSI Part"),
            ("ARTICLE", "CSI_Article__ARCH", "CSI Article"),
        ] + roles
    rows = [(paragraph, *role) for paragraph, role in zip(classifiable, roles)]
    return {
        "create_styles": [
            {
                "styleId": style_id,
                "name": name,
                "type": "paragraph",
                "derive_from_paragraph_index": paragraph["paragraph_index"],
                "basedOn": "Normal",
                "role": role,
            }
            for paragraph, role, style_id, name in rows
        ],
        "apply_pStyle": [
            {"paragraph_index": paragraph["paragraph_index"], "styleId": style_id}
            for paragraph, _role, style_id, _name in rows
        ],
        "ignored_paragraphs": [],
        "roles": {
            role: {
                "styleId": style_id,
                "exemplar_paragraph_index": paragraph["paragraph_index"],
            }
            for paragraph, role, style_id, _name in rows
        },
        "notes": ["independent verification fixture"],
    }


# --- Each mode's hand-written fixture and predictions ---------------------------

#: Per mode: the target's source text and its predicted output text, the
#: architect's even/odd switch (None where the target keeps its own header
#: set), the revision census, and the package members predicted to change,
#: be added (besides the architect's header set) and be removed.
PREDICTIONS = {
    FORMAT_ONLY: {
        "source_texts": FORMAT_ONLY_SOURCE_TEXTS,
        "expected_texts": FORMAT_ONLY_EXPECTED_TEXTS,
        "changed_paragraphs": (),
        "architect_even_and_odd": True,
        "target_even_and_odd": False,
        "revisions": FORMAT_ONLY_REVISIONS,
        "changed": {
            "[Content_Types].xml",
            "word/_rels/document.xml.rels",
            "word/document.xml",
            "word/fontTable.xml",
            "word/settings.xml",
            "word/styles.xml",
            "word/theme/theme1.xml",
        },
        "added": set(),
        "removed": {"word/header9.xml"},
    },
    CSI_TO_CANADIAN: {
        "source_texts": ARCHITECT_MODE_SOURCE_TEXTS,
        "expected_texts": CSI_TO_CANADIAN_EXPECTED_TEXTS,
        "changed_paragraphs": CSI_TO_CANADIAN_CHANGED,
        # A dormant even header: the architect references one with the switch
        # off, so Word renders its default header on even pages. The target's
        # switch is on, and must be cleared with its headers replaced.
        "architect_even_and_odd": False,
        "target_even_and_odd": True,
        "revisions": ARCHITECT_MODE_REVISIONS,
        "changed": {
            "[Content_Types].xml",
            "word/_rels/document.xml.rels",
            "word/document.xml",
            "word/fontTable.xml",
            "word/numbering.xml",
            "word/settings.xml",
            "word/styles.xml",
            "word/theme/theme1.xml",
        },
        "added": set(),
        "removed": {"word/header9.xml"},
    },
    CSI_TO_CANADIAN_STANDALONE: {
        "source_texts": STANDALONE_SOURCE_TEXTS,
        "expected_texts": STANDALONE_EXPECTED_TEXTS,
        "changed_paragraphs": STANDALONE_CHANGED,
        "architect_even_and_odd": None,
        "target_even_and_odd": True,
        "revisions": ARCHITECT_MODE_REVISIONS,
        # No shell and no role styles: the built-in list joins the target's
        # own numbering part, which is already wired in.
        "changed": {"word/document.xml", "word/numbering.xml"},
        "added": set(),
        "removed": set(),
    },
}

#: The paragraphs whose XML each mode edits at all: restyled in Format-only
#: (text unchanged), converted in the Canadian modes. Every other paragraph,
#: nested ones included, must come through element for element.
EDITED_PARAGRAPHS = {
    FORMAT_ONLY: (3, 4, 5, 6),
    CSI_TO_CANADIAN: CSI_TO_CANADIAN_CHANGED,
    CSI_TO_CANADIAN_STANDALONE: STANDALONE_CHANGED,
}

#: What each mode may touch at all, written from the mode, by what the
#: member *is* (resolved through the packages' relationships) rather than by
#: name. ``shell``: the architect's document shell and header set.
REMIT = {
    FORMAT_ONLY: {"styles", "shell", "numbering"},
    CSI_TO_CANADIAN: {"styles", "shell", "numbering"},
    CSI_TO_CANADIAN_STANDALONE: {"numbering"},
}


def _produce(tmp_path: Path, mode: str) -> tuple[Path, object]:
    """Write the mode's fixture, run the public entry point, and prove the
    inputs came through it unchanged. Returns the target and the run."""

    expected = PREDICTIONS[mode]
    target_path = tmp_path / "source" / "21 13 13 Sprinklers.docx"
    if mode == CSI_TO_CANADIAN_STANDALONE:
        # The target brings a numbering part of its own. A target with its own
        # header and *no* numbering part is withheld today: wiring in a new
        # numbering part re-serializes the document relationships, and the
        # gate compares the target's header relationships as raw text
        # (recorded in handoff 09, not fixed by this item).
        target = _write_target(
            target_path,
            _standalone_target_body(),
            numbering=True,
            settings=_target_settings(even_and_odd=expected["target_even_and_odd"], utf16=True),
        )
        architect = None
        options: dict = {}
    else:
        target = _write_target(
            target_path,
            _architect_mode_target_body(tracked_item=mode == FORMAT_ONLY),
            numbering=True,
            settings=_target_settings(even_and_odd=expected["target_even_and_odd"], utf16=False),
        )
        architect = _write_architect(
            tmp_path / "template" / "architect.docx",
            canadian=mode == CSI_TO_CANADIAN,
            even_and_odd=expected["architect_even_and_odd"],
        )
        options = {
            "cache_dir": tmp_path / "cache",
            "template_model": f"wi08-{mode}-fixture",
            "template_classifier": _template_classifier,
        }
    inputs = {path: path.read_bytes() for path in (target, architect) if path is not None}
    run = format_specifications(
        architect_template=architect,
        target_specs=[target],
        output_dir=tmp_path / "out",
        api_key="",
        max_workers=1,
        conversion_mode=mode,
        **options,
    )
    # The inputs are immutable: the run works on private snapshots of them.
    for path, payload in inputs.items():
        assert path.read_bytes() == payload, path.name
    return target, run


@pytest.fixture(scope="module", params=ALL_MODES)
def produced(request, tmp_path_factory) -> tuple[str, Path, Path]:
    """Run each mode once; every check below reads the same two packages."""

    mode = request.param
    source, run = _produce(tmp_path_factory.mktemp(mode), mode)
    assert run.success, "\n".join(run.targets[0].log)
    output = run.targets[0].output_path
    assert output is not None and output.is_file()
    return mode, source, output


# --- Independent readers (standard library only) --------------------------------


def _members(path: Path) -> dict[str, bytes]:
    with zipfile.ZipFile(path) as package:
        return {name: package.read(name) for name in package.namelist()}


def _rels_part_for(part: str) -> str:
    directory, name = posixpath.split(part)
    return posixpath.join(directory, "_rels", f"{name}.rels")


def _relationships_of(members: dict[str, bytes], part: str) -> list[dict[str, str]]:
    """Every relationship of ``part``, its internal targets resolved to part names."""

    payload = members.get(_rels_part_for(part))
    if payload is None:
        return []
    found = []
    for element in ET.fromstring(payload).iter(f"{{{PKG_REL}}}Relationship"):
        target = element.get("Target", "")
        external = element.get("TargetMode") == "External"
        if not external:
            if target.startswith("/"):
                target = target[1:]
            else:
                target = posixpath.normpath(posixpath.join(posixpath.dirname(part), target))
        found.append(
            {
                "id": element.get("Id", ""),
                "type": element.get("Type", "").rsplit("/", 1)[-1],
                "target": target,
                "external": external,
            }
        )
    return found


def _main_document(members: dict[str, bytes]) -> str:
    documents = [
        rel["target"]
        for rel in _relationships_of(members, "")
        if rel["type"] == "officeDocument"
    ]
    assert len(documents) == 1, documents
    return documents[0]


def _related(members: dict[str, bytes], kind: str) -> set[str]:
    """The parts the main document relates with relationship type ``kind``."""

    document = _main_document(members)
    return {
        rel["target"]
        for rel in _relationships_of(members, document)
        if rel["type"] == kind and not rel["external"]
    }


def _content_type(members: dict[str, bytes], name: str) -> str | None:
    root = ET.fromstring(members["[Content_Types].xml"])
    for override in root.iter(f"{{{CT}}}Override"):
        if override.get("PartName", "").lstrip("/").lower() == name.lower():
            return override.get("ContentType")
    extension = posixpath.splitext(name)[1].lstrip(".").lower()
    for default in root.iter(f"{{{CT}}}Default"):
        if default.get("Extension", "").lower() == extension:
            return default.get("ContentType")
    return None


def _is_xml_part(members: dict[str, bytes], name: str) -> bool:
    if name.lower().endswith((".xml", ".rels")):
        return True
    content_type = (_content_type(members, name) or "").lower()
    return content_type.endswith("+xml") or content_type in {"application/xml", "text/xml"}


# The rendering every paragraph is predicted in. Plain text is itself; the
# rest is spelled out so no two different paragraphs render alike: a tab is
# \t, a line break \n, a carriage return \r, a soft hyphen U+00AD, a
# non-breaking hyphen U+2011, a tracked insertion {+...+}, a tracked deletion
# {-...-}, and everything else a bracketed token. A text node without
# xml:space="preserve" loses its edge whitespace, as Word reads it. A nested
# paragraph (a text box's) renders on its own, after the one holding it.

_XML_WHITESPACE = " \t\r\n"
_INERT_RUN_CHILDREN = {"rPr", "lastRenderedPageBreak"}
_DRAWINGS = {_w("drawing"), _w("pict"), _w("object"), f"{{{MC}}}AlternateContent"}
_WRAPPERS = {
    "ins": ("{+", "+}"),
    "del": ("{-", "-}"),
    "moveFrom": ("{<", "<}"),
    "moveTo": ("{>", ">}"),
    "hyperlink": ("[link]", "[/link]"),
}


def _local(element: ET.Element) -> str:
    return element.tag.split("}", 1)[1] if element.tag.startswith(f"{{{W}}}") else element.tag


def _text_of(element: ET.Element) -> str:
    text = element.text or ""
    if element.get(f"{{{XML_NS}}}space") != "preserve":
        text = text.strip(_XML_WHITESPACE)
    return text


#: Which text and field-instruction elements a run may hold, by the revision
#: it sits in. Deleted text is w:delText and only deleted text is, so a w:t
#: inside a deletion, or a w:delText outside one, renders as what it is not.
#: Moved-from text is both deleted and still in the document; either is read.
_TEXT_TAGS = {
    "plain": ({"t"}, {"instrText"}),
    "deleted": ({"delText"}, {"delInstrText"}),
    "moved_from": ({"t", "delText"}, {"instrText", "delInstrText"}),
}


def _render_run(run: ET.Element, *, context: str) -> str:
    text_tags, instruction_tags = _TEXT_TAGS[context]
    out = []
    for child in run:
        if child.tag in _DRAWINGS:
            out.append("[drawing]")
            continue
        local = _local(child)
        if local in _INERT_RUN_CHILDREN:
            continue
        if local in text_tags:
            out.append(_text_of(child))
        elif local in instruction_tags:
            out.append(f"[instr {child.text or ''}]")
        elif local == "tab":
            out.append("\t")
        elif local == "br":
            out.append({"page": "\f", "column": "\v"}.get(child.get(_w("type"), ""), "\n"))
        elif local == "cr":
            out.append("\r")
        elif local == "softHyphen":
            out.append("\u00ad")
        elif local == "noBreakHyphen":
            out.append("\u2011")
        elif local == "ptab":
            out.append("[ptab]")
        elif local == "sym":
            out.append(f"[sym {child.get(_w('font'))} {child.get(_w('char'))}]")
        elif local == "fldChar":
            out.append(f"[fld {child.get(_w('fldCharType'))}]")
        elif local == "footnoteReference":
            out.append(f"[footnote {child.get(_w('id'))}]")
        elif local == "endnoteReference":
            out.append(f"[endnote {child.get(_w('id'))}]")
        elif local == "commentReference":
            out.append(f"[comment {child.get(_w('id'))}]")
        else:
            out.append(f"[? {child.tag}]")
    return "".join(out)


def _render(element: ET.Element, *, context: str = "plain") -> str:
    out = []
    for child in element:
        local = _local(child)
        if local in ("p", "pPr"):
            continue
        if local == "r":
            out.append(_render_run(child, context=context))
        elif local in _WRAPPERS:
            opening, closing = _WRAPPERS[local]
            inner_context = context
            if local == "del":
                inner_context = "deleted"
            elif local == "moveFrom" and context == "plain":
                inner_context = "moved_from"
            out.append(opening + _render(child, context=inner_context) + closing)
        elif local in _REVISION_RANGES:
            out.append(f"[{local} {child.get(_w('id'))}]")
        elif local == "bookmarkStart":
            out.append(f"[bm {child.get(_w('id'))} {child.get(_w('name'))}]")
        elif local == "bookmarkEnd":
            out.append(f"[/bm {child.get(_w('id'))}]")
        elif local == "commentRangeStart":
            out.append(f"[cr {child.get(_w('id'))}]")
        elif local == "commentRangeEnd":
            out.append(f"[/cr {child.get(_w('id'))}]")
        else:
            out.append(f"[? {child.tag}]")
    return "".join(out)


def _paragraph_texts(document: bytes) -> list[str]:
    return [_render(paragraph) for paragraph in ET.fromstring(document).iter(_w("p"))]


_REVISION_ELEMENTS = {
    "ins", "del", "moveFrom", "moveTo", "cellIns", "cellDel", "cellMerge",
    "rPrChange", "pPrChange", "sectPrChange", "tblPrChange", "tblPrExChange",
    "tblGridChange", "trPrChange", "tcPrChange", "numberingChange",
}
_REVISION_RANGES = {
    "moveFromRangeStart", "moveFromRangeEnd", "moveToRangeStart", "moveToRangeEnd",
    "customXmlInsRangeStart", "customXmlInsRangeEnd",
    "customXmlDelRangeStart", "customXmlDelRangeEnd",
    "customXmlMoveFromRangeStart", "customXmlMoveFromRangeEnd",
    "customXmlMoveToRangeStart", "customXmlMoveToRangeEnd",
}
#: Annotations whose id is shared by design: a bookmark's two ends, a
#: comment's range, reference and body, and every revision range's two ends.
_PAIRED_ANNOTATIONS = {
    "bookmarkStart", "bookmarkEnd", "commentRangeStart", "commentRangeEnd",
    "commentReference", "comment",
} | _REVISION_RANGES


def _revision_census(document: bytes) -> dict[tuple[str | None, str], int]:
    counts: Counter = Counter()
    for element in ET.fromstring(document).iter():
        if not element.tag.startswith(f"{{{W}}}"):
            continue
        local = _local(element)
        if local in _REVISION_ELEMENTS or local in _REVISION_RANGES:
            counts[(element.get(_w("author")), local)] += 1
    return dict(counts)


def _annotation_id(raw: str) -> int | str:
    """An id as Word reads it: ``01`` and ``1`` are the same annotation."""

    return int(raw) if re.fullmatch(r"[+-]?\d+", raw.strip()) else raw


def _annotations(members: dict[str, bytes]) -> list[tuple[str, str, int | str]]:
    """(part, local name, id) for every revision and paired annotation."""

    found = []
    for name, payload in sorted(members.items()):
        if not name.startswith("word/") or not _is_xml_part(members, name):
            continue
        for element in ET.fromstring(payload).iter():
            if not element.tag.startswith(f"{{{W}}}"):
                continue
            local = _local(element)
            if local in _REVISION_ELEMENTS | _PAIRED_ANNOTATIONS:
                identifier = element.get(_w("id"))
                if identifier is not None:
                    found.append((name, local, _annotation_id(identifier)))
    return found


def _even_and_odd_headers(settings: bytes) -> bool:
    switches = ET.fromstring(settings).findall(_w("evenAndOddHeaders"))
    assert len(switches) <= 1, "one switch at most"
    if not switches:
        return False
    value = (switches[0].get(_w("val")) or "true").strip().lower()
    assert value in {"true", "1", "on", "false", "0", "off"}, value
    return value in {"true", "1", "on"}


def _section_references(document: bytes, members: dict[str, bytes]) -> dict[tuple[str, str], str]:
    """(header|footer, type) -> the part each section reference resolves to."""

    part = _main_document(members)
    by_id = {rel["id"]: rel["target"] for rel in _relationships_of(members, part)}
    found = {}
    for section in ET.fromstring(document).iter(_w("sectPr")):
        for kind in ("header", "footer"):
            for reference in section.findall(_w(f"{kind}Reference")):
                key = (kind, reference.get(_w("type")))
                target = by_id[reference.get(f"{{{R}}}id")]
                assert found.setdefault(key, target) == target, key
    return found


def _declared_namespaces_cover_markup_compatibility(name: str, payload: bytes) -> None:
    """Parse ``payload``, and require every prefix an ``mc:`` attribute names
    to be declared where it is used.

    An unbound element or attribute prefix already fails the parse. A prefix
    named only inside an attribute *value* does not, so it is checked here.
    """

    try:
        ET.fromstring(payload)
    except ET.ParseError as exc:
        raise AssertionError(f"{name} does not parse: {exc}") from exc
    scopes: list[dict[str, str]] = [{}]
    pending: dict[str, str] = {}
    for event, item in ET.iterparse(io.BytesIO(payload), events=("start-ns", "start", "end")):
        if event == "start-ns":
            prefix, uri = item
            pending[prefix] = uri
        elif event == "start":
            scopes.append({**scopes[-1], **pending})
            pending = {}
            named = []
            for attribute, value in item.attrib.items():
                if attribute in (f"{{{MC}}}Ignorable", f"{{{MC}}}MustUnderstand"):
                    named.extend(value.split())
                elif attribute in (
                    f"{{{MC}}}ProcessContent",
                    f"{{{MC}}}PreserveElements",
                    f"{{{MC}}}PreserveAttributes",
                ):
                    named.extend(token.split(":", 1)[0] for token in value.split())
                elif attribute == "Requires" and item.tag == f"{{{MC}}}Choice":
                    named.extend(value.split())
            undeclared = [prefix for prefix in named if prefix not in scopes[-1]]
            assert not undeclared, f"{name}: undeclared prefixes {undeclared} in {item.tag}"
        else:
            scopes.pop()


# --- The checks -------------------------------------------------------------------
#
# Each check is a function of the mode and both packages' members, so that the
# self-test at the end of the module can run it against damaged output too.

Members = dict[str, bytes]


def _check_source_reads_as_written(mode: str, before: Members, after: Members) -> None:
    """Guard on the fixtures: the reader agrees with the hand-written source text.

    Every text check compares the output with a hand-written prediction; this
    one proves the predictions start from the document as it really is.
    """

    texts = _paragraph_texts(before[_main_document(before)])
    assert tuple(texts) == PREDICTIONS[mode]["source_texts"]


def _check_paragraph_texts(mode: str, before: Members, after: Members) -> None:
    """Exact text, deleted text and structural children, paragraph by paragraph.

    A paragraph the mode is not predicted to change must read exactly as it
    did; a changed one must read exactly as written in the prediction.
    """

    texts = _paragraph_texts(after[_main_document(after)])
    expected = PREDICTIONS[mode]["expected_texts"]
    assert len(texts) == len(expected), "paragraph count"
    for index, (found, predicted) in enumerate(zip(texts, expected)):
        assert found == predicted, f"paragraph {index}: {found!r} != {predicted!r}"


def _shape(element: ET.Element) -> tuple:
    """An element's complete structure -- tag, attributes, text, and every
    child with its tail -- everything but its own tail."""

    return (
        element.tag,
        tuple(sorted(element.attrib.items())),
        element.text,
        tuple((_shape(child), child.tail) for child in element),
    )


def _check_untouched_structure(mode: str, before: Members, after: Members) -> None:
    """Element identity where the mode is not predicted to edit at all.

    Stronger than text: a property, attribute or container changed in a
    paragraph the mode leaves alone is a difference here even when every
    character still reads the same. Tables are out of scope in every mode; a
    mode with no shell leaves every section's properties alone too.
    """

    source_root = ET.fromstring(before[_main_document(before)])
    output_root = ET.fromstring(after[_main_document(after)])
    source = list(source_root.iter(_w("p")))
    output = list(output_root.iter(_w("p")))
    assert len(source) == len(output), "paragraph count"
    edited = EDITED_PARAGRAPHS[mode]
    for index, (b, a) in enumerate(zip(source, output)):
        if index in edited:
            assert _shape(b) != _shape(a), f"paragraph {index} was predicted to be edited"
        else:
            assert _shape(b) == _shape(a), f"paragraph {index}"
    tables = [
        (_shape(b), _shape(a))
        for b, a in zip(source_root.iter(_w("tbl")), output_root.iter(_w("tbl")))
    ]
    assert len(tables) == 1 and tables[0][0] == tables[0][1], "table"
    if "shell" not in REMIT[mode]:
        assert [_shape(s) for s in source_root.iter(_w("sectPr"))] == [
            _shape(s) for s in output_root.iter(_w("sectPr"))
        ], "section properties"


def _check_package_whitelist(mode: str, before: Members, after: Members) -> None:
    """The package whitelist, from the mode, by what each member is.

    Every member outside the remit is byte-identical; the members the mode
    changes, adds and removes are exactly the predicted ones.
    """

    digests_before = {name: hashlib.sha256(payload).hexdigest() for name, payload in before.items()}
    digests_after = {name: hashlib.sha256(payload).hexdigest() for name, payload in after.items()}
    changed = {n for n in set(before) & set(after) if digests_before[n] != digests_after[n]}
    added = set(after) - set(before)
    removed = set(before) - set(after)

    remit = REMIT[mode]
    document = _main_document(after)
    assert document == _main_document(before) == "word/document.xml"
    may_change = {document}
    if "styles" in remit:
        may_change |= _related(before, "styles") | _related(after, "styles")
    if "numbering" in remit or "shell" in remit:
        # Wiring a part in or out: its content type and its relationship.
        may_change |= {"[Content_Types].xml", _rels_part_for(document)}
    if "numbering" in remit:
        may_change |= _related(before, "numbering") | _related(after, "numbering")
    header_set_before: set[str] = set()
    header_set_after: set[str] = set()
    header_rels: set[str] = set()
    media: set[str] = set()
    if "shell" in remit:
        for kind in ("settings", "theme", "fontTable"):
            may_change |= _related(before, kind) | _related(after, kind)
        for kind in ("header", "footer"):
            header_set_before |= _related(before, kind)
            header_set_after |= _related(after, kind)
        header_rels = {
            _rels_part_for(part) for part in header_set_after if _rels_part_for(part) in after
        }
        media = {
            rel["target"]
            for part in header_set_after
            for rel in _relationships_of(after, part)
            if rel["type"] == "image"
        }
        may_change |= header_set_after | header_rels | media
    may_remove = header_set_before | {_rels_part_for(part) for part in header_set_before}

    assert changed | added <= may_change, sorted((changed | added) - may_change)
    assert removed <= may_remove, sorted(removed - may_remove)

    expected = PREDICTIONS[mode]
    assert changed == expected["changed"], sorted(changed ^ expected["changed"])
    assert removed == expected["removed"], sorted(removed ^ expected["removed"])
    if "shell" in remit:
        # The architect's header set, under whatever names the importer
        # chose: four parts, the image header's relationships, its one image.
        assert header_set_after.isdisjoint(before)
        assert len(header_set_after) == 4 and len(header_rels) == 1 and len(media) == 1
        assert after[next(iter(media))] == PNG_1X1
        imported = header_set_after | header_rels | media
        assert added == expected["added"] | imported, sorted(added)
    else:
        assert added == expected["added"], sorted(added)
    # The parts no mode has any business with, byte for byte.
    for name in (
        "word/footnotes.xml",
        "word/comments.xml",
        "customXml/item2.data",
        "[trash]/0000.dat",
    ):
        assert before[name] == after[name], name


def _check_header_set(mode: str, before: Members, after: Members) -> None:
    """Each section references the architect's parts, reading as the
    architect's -- or, with no architect, the target's own, byte-identical.

    Header and footer wording is not the target's to keep in the architect
    modes: it is replaced, so it is held to the architect's, not the target's.
    """

    references = _section_references(after[_main_document(after)], after)
    if PREDICTIONS[mode]["architect_even_and_odd"] is None:
        assert references == {("header", "default"): "word/header9.xml"}
        assert after["word/header9.xml"] == before["word/header9.xml"]
        return
    assert set(references) == set(ARCHITECT_HEADER_TEXTS)
    for key, part in references.items():
        assert tuple(_paragraph_texts(after[part])) == ARCHITECT_HEADER_TEXTS[key], key


def _check_revision_census(mode: str, before: Members, after: Members) -> None:
    """Every revision in the main document, by author and kind, before and after."""

    predicted = PREDICTIONS[mode]["revisions"]
    assert _revision_census(before[_main_document(before)]) == predicted
    assert _revision_census(after[_main_document(after)]) == predicted


def _check_revision_ids(mode: str, before: Members, after: Members) -> None:
    """Unique among revision elements across the output, and only among them.

    Valid OOXML repeats an id across a paired annotation -- a bookmark's start
    and end, a comment's range, reference and the comment itself -- so those
    are held to the source instead.
    """

    identifiers = [
        identifier
        for _part, local, identifier in _annotations(after)
        if local in _REVISION_ELEMENTS
    ]
    assert identifiers, "check must not pass vacuously"
    duplicated = sorted(value for value, count in Counter(identifiers).items() if count > 1)
    assert not duplicated, duplicated

    def paired(members: Members) -> list[tuple[str, str, int | str]]:
        return [item for item in _annotations(members) if item[1] in _PAIRED_ANNOTATIONS]

    assert paired(before), "check must not pass vacuously"
    assert paired(after) == paired(before)


def _check_header_parity(mode: str, before: Members, after: Members) -> None:
    """The even/odd switch follows the header set it governs.

    Imported, the architect's set brings its switch: on for the Format-only
    architect, off (a dormant even header) for the Canadian one, whatever the
    target's was. Kept, the target's own settings part is byte-identical.
    """

    (settings_before,) = _related(before, "settings")
    (settings_after,) = _related(after, "settings")
    expected = PREDICTIONS[mode]
    assert _even_and_odd_headers(before[settings_before]) is expected["target_even_and_odd"]
    if expected["architect_even_and_odd"] is None:
        assert settings_after == settings_before
        assert after[settings_after] == before[settings_before]
        return
    assert _even_and_odd_headers(after[settings_after]) is expected["architect_even_and_odd"]
    # The rest of the settings stay the target's: its zoom, not the architect's.
    zoom = ET.fromstring(after[settings_after]).find(_w("zoom"))
    assert zoom is not None and zoom.get(_w("percent")) == "110"


def _check_namespaces(mode: str, before: Members, after: Members) -> None:
    """Every XML part, by extension or by declared content type, parses; and
    every prefix a markup-compatibility attribute names is declared."""

    xml_parts = [name for name in sorted(after) if _is_xml_part(after, name)]
    assert "customXml/item2.data" in xml_parts, "the content-type-only XML part is checked"
    for name in xml_parts:
        _declared_namespaces_cover_markup_compatibility(name, after[name])
    # The markup-compatibility half has something to check.
    root_tag = re.search(rb"<w:document\b[^>]*>", after[_main_document(after)])
    assert root_tag is not None and b'mc:Ignorable="w14"' in root_tag.group(0)


@pytest.mark.parametrize("mode", ALL_MODES)
def test_the_predicted_changes_are_exactly_the_edited_paragraphs(mode: str) -> None:
    """The hand-written tables agree with each other: text differs exactly
    where each mode converts, and nowhere in Format-only."""

    expected = PREDICTIONS[mode]
    changed = tuple(
        index
        for index, (b, a) in enumerate(zip(expected["source_texts"], expected["expected_texts"]))
        if b != a
    )
    assert len(expected["source_texts"]) == len(expected["expected_texts"])
    assert changed == expected["changed_paragraphs"]
    if mode == FORMAT_ONLY:
        assert changed == ()
    else:
        assert changed == EDITED_PARAGRAPHS[mode]


def test_the_fixtures_read_as_written(produced) -> None:
    mode, source, output = produced
    _check_source_reads_as_written(mode, _members(source), _members(output))


def test_every_paragraph_reads_exactly_as_predicted(produced) -> None:
    mode, source, output = produced
    _check_paragraph_texts(mode, _members(source), _members(output))


def test_paragraphs_the_mode_does_not_edit_are_element_identical(produced) -> None:
    mode, source, output = produced
    _check_untouched_structure(mode, _members(source), _members(output))


def test_only_members_inside_the_remit_change(produced) -> None:
    mode, source, output = produced
    _check_package_whitelist(mode, _members(source), _members(output))


def test_the_header_set_is_the_architects_or_the_targets_own(produced) -> None:
    mode, source, output = produced
    _check_header_set(mode, _members(source), _members(output))


def test_the_revision_census_is_unchanged(produced) -> None:
    mode, source, output = produced
    _check_revision_census(mode, _members(source), _members(output))


def test_revision_ids_are_unique_and_paired_annotations_keep_theirs(produced) -> None:
    mode, source, output = produced
    _check_revision_ids(mode, _members(source), _members(output))


def test_even_and_odd_header_parity(produced) -> None:
    mode, source, output = produced
    _check_header_parity(mode, _members(source), _members(output))


def test_every_output_xml_part_parses_with_its_declared_namespaces(produced) -> None:
    mode, source, output = produced
    _check_namespaces(mode, _members(source), _members(output))


# --- The checks can fail ------------------------------------------------------------
#
# A check that has never rejected anything proves nothing. Each mutation below
# damages the real output of every mode the way a defect would, and the check
# that exists to find that damage must reject it. The damage is chosen so the
# other checks mostly cannot see it: an attribute on a table, a revision
# re-signed by another author, a prefix named only in an attribute value.


def _replace_once(payload: bytes, old: bytes, new: bytes) -> bytes:
    assert payload.count(old) == 1, f"fixture drifted: {old!r}"
    return payload.replace(old, new)


def _in_document(old: bytes, new: bytes):
    def mutate(members: Members) -> Members:
        document = _main_document(members)
        return {**members, document: _replace_once(members[document], old, new)}

    return mutate


def _in_part(name: str, old: bytes, new: bytes):
    def mutate(members: Members) -> Members:
        return {**members, name: _replace_once(members[name], old, new)}

    return mutate


def _with_member(name: str, payload: bytes | None):
    def mutate(members: Members) -> Members:
        changed = dict(members)
        if payload is None:
            del changed[name]
        else:
            changed[name] = payload
        return changed

    return mutate


def _toggle_even_and_odd_headers(members: Members) -> Members:
    (settings,) = _related(members, "settings")
    payload = members[settings]
    for bom, codec in (
        (codecs.BOM_UTF16_BE, "utf-16-be"),
        (codecs.BOM_UTF16_LE, "utf-16-le"),
        (codecs.BOM_UTF8, "utf-8"),
        (b"", "utf-8"),
    ):
        if payload.startswith(bom):
            break
    text = payload[len(bom):].decode(codec)
    if "<w:evenAndOddHeaders/>" in text:
        text = text.replace("<w:evenAndOddHeaders/>", "", 1)
    else:
        text = re.sub(r"(<w:settings\b[^>]*>)", r"\1<w:evenAndOddHeaders/>", text, count=1)
    return {**members, settings: bom + text.encode(codec)}


_REVIEWER_BYTES = REVIEWER.encode("utf-8")

MUTATIONS = {
    "a tab beside a word dropped": (
        _in_document(b"<w:tab/><w:t>for</w:t>", b"<w:t>for</w:t>"),
        _check_paragraph_texts,
    ),
    "an edge space no longer preserved": (
        _in_document(
            b'<w:t xml:space="preserve">Retain  this note </w:t>',
            b"<w:t>Retain  this note </w:t>",
        ),
        _check_paragraph_texts,
    ),
    "a tracked deletion made plain text": (
        _in_document(
            b'<w:delText xml:space="preserve"> later</w:delText>',
            b'<w:t xml:space="preserve"> later</w:t>',
        ),
        _check_paragraph_texts,
    ),
    "a table's properties changed": (
        _in_document(b'<w:tblW w:w="0" w:type="auto"/>', b'<w:tblW w:w="5000" w:type="pct"/>'),
        _check_untouched_structure,
    ),
    "a run property added to an ignored paragraph": (
        _in_document(b"<w:rPr><w:b/><w:rPrChange", b"<w:rPr><w:b/><w:i/><w:rPrChange"),
        _check_untouched_structure,
    ),
    "a run-property revision dropped": (
        _in_document(b'<w:rPrChange w:id="13" ' + _REVIEWER_BYTES + b"><w:rPr/></w:rPrChange>", b""),
        _check_revision_census,
    ),
    "a reviewer's deletion re-signed": (
        _in_document(
            b'<w:del w:id="12" w:author="Reviewer"',
            b'<w:del w:id="12" w:author="Specification Formatter"',
        ),
        _check_revision_census,
    ),
    "a revision id duplicated": (
        _in_document(b'<w:del w:id="16" ', b'<w:del w:id="15" '),
        _check_revision_ids,
    ),
    "a revision id duplicated under another spelling": (
        _in_document(b'<w:del w:id="16" ', b'<w:del w:id="015" '),
        _check_revision_ids,
    ),
    "a move range's end renumbered": (
        _in_document(b'<w:moveFromRangeEnd w:id="41"/>', b'<w:moveFromRangeEnd w:id="45"/>'),
        _check_revision_ids,
    ),
    "a bookmark's end renumbered": (
        _in_document(b'<w:bookmarkEnd w:id="31"/>', b'<w:bookmarkEnd w:id="32"/>'),
        _check_revision_ids,
    ),
    "the footnotes part changed": (
        _in_part("word/footnotes.xml", b"A footnote under review.", b"A footnote under review!"),
        _check_package_whitelist,
    ),
    "a stray member added": (
        _with_member("word/stray.xml", b"<stray/>"),
        _check_package_whitelist,
    ),
    "a trash item dropped": (
        _with_member("[trash]/0000.dat", None),
        _check_package_whitelist,
    ),
    "the even/odd switch flipped": (
        _toggle_even_and_odd_headers,
        _check_header_parity,
    ),
    "an ignorable prefix left undeclared": (
        _in_document(b'mc:Ignorable="w14"', b'mc:Ignorable="w14 w15"'),
        _check_namespaces,
    ),
    "an undeclared prefix in a preserve list": (
        _in_document(
            b'mc:Ignorable="w14"', b'mc:Ignorable="w14" mc:PreserveElements="w15:x"'
        ),
        _check_namespaces,
    ),
    "a content-type-only XML part malformed": (
        _with_member("customXml/item2.data", b"<data><item>1</item>"),
        _check_namespaces,
    ),
}


@pytest.mark.parametrize("mutation", sorted(MUTATIONS))
def test_each_check_rejects_the_damage_it_exists_to_find(produced, mutation: str) -> None:
    mode, source, output = produced
    before, after = _members(source), _members(output)
    mutate, check = MUTATIONS[mutation]
    check(mode, before, after)
    damaged = mutate(after)
    assert damaged != after
    with pytest.raises(AssertionError):
        check(mode, before, damaged)


# --- The verifier's own independence ----------------------------------------------

_ALLOWED_IMPORTS = {
    "__future__",
    "ast",
    "codecs",
    "collections",
    "hashlib",
    "io",
    "pathlib",
    "posixpath",
    "pytest",
    "re",
    "sys",
    "xml.etree.ElementTree",
    "zipfile",
}


def test_module_imports_nothing_from_the_engine() -> None:
    """The one engine import is the public entry point that produces the output."""

    tree = ast.parse(Path(__file__).read_text(encoding="utf-8"))
    engine_imports = []
    for node in ast.walk(tree):
        if isinstance(node, ast.Import):
            for alias in node.names:
                assert alias.name in _ALLOWED_IMPORTS, alias.name
        elif isinstance(node, ast.ImportFrom):
            if node.module == "spec_formatter":
                engine_imports.extend(alias.name for alias in node.names)
            else:
                assert node.level == 0 and node.module in _ALLOWED_IMPORTS, node.module
    assert engine_imports == ["format_specifications"]
    for module in _ALLOWED_IMPORTS - {"__future__", "pytest"}:
        top = module.split(".", 1)[0]
        assert top in sys.stdlib_module_names, module
