"""The exact run-content signature behind the Format-only body gate.

``paragraph_text_from_block`` removes deleted text, maps tabs and breaks to
spaces and collapses whitespace. That is the right reading for
classification and a blind one for a gate: the same function run on both
sides of a transform cannot see a tab or break dropped beside a space, a lost
``xml:space="preserve"``, a non-breaking space turned into a plain one, a
collapsed double space, a dropped soft hyphen, or tracked-deleted text
emptied. ``paragraph_run_content_signature`` records every run's content
children exactly, as appendix A of the DOCX Method Hardening plan specifies,
and ``run_content_difference`` names the kind of item that differs -- never
its text.
"""

from __future__ import annotations

import pytest

from spec_formatter.style_application.core.xml_helpers import (
    RUN_CONTENT_DIFFERENCE_KINDS,
    paragraph_run_content_signature,
    paragraph_text_from_block,
    run_content_difference,
)


def _signature(paragraph_content: str):
    return paragraph_run_content_signature(f"<w:p>{paragraph_content}</w:p>")


def _run(content: str) -> str:
    return f"<w:r>{content}</w:r>"


# ── text-bearing children ──────────────────────────────────────────────────


def test_text_is_decoded_but_never_trimmed_collapsed_or_mapped():
    signature = _signature(
        _run(
            '<w:t xml:space="preserve">  two  spaces nbsp\ttab '
            "&amp; &lt;&gt;&quot;&apos; &#x41;&#66; </w:t>"
        )
    )

    assert signature == (
        (("t", "  two  spaces nbsp\ttab & <>\"' AB ", True),),
    )


@pytest.mark.parametrize(
    ("start_tag", "preserve"),
    [
        ('<w:t xml:space="preserve">', True),
        ("<w:t xml:space='preserve'>", True),
        ('<w:t xml:space = "preserve" >', True),
        ('<w:t xml:space="default">', False),
        ("<w:t>", False),
    ],
)
def test_preserve_flag_reads_the_attribute_in_any_quoting(start_tag, preserve):
    assert _signature(_run(f"{start_tag}x</w:t>")) == ((("t", "x", preserve),),)


def test_decoding_is_xml_only_not_html():
    # html.unescape maps the C1 range through the Windows-1252 table, so
    # "&#x80;" would come back as a euro sign. XML means U+0080.
    assert _signature(_run("<w:t>&#x80;&#150;</w:t>")) == (
        (("t", "\u0080\u0096", False),),
    )
    # An escaped ampersand before an HTML entity name is literal text.
    assert _signature(_run("<w:t>&amp;nbsp;</w:t>")) == ((("t", "&nbsp;", False),),)


def test_cdata_is_literal_and_comments_are_not_character_content():
    signature = _signature(
        _run("<w:t><![CDATA[a<b &amp; c]]><!-- editorial note --><?pi x?>d</w:t>")
    )

    assert signature == ((("t", "a<b &amp; cd", False),),)


def test_line_ends_are_normalized_the_way_an_xml_parser_reports_them():
    # A literal CR cannot survive any XML parser (XML 1.0 section 2.11), so
    # Word never sees one; a character reference to CR is real content.
    signature = _signature(_run("<w:t>a\r\nb\rc&#13;</w:t>"))

    assert signature == ((("t", "a\nb\nc\r", False),),)


def test_each_text_bearing_child_is_its_own_item():
    signature = _signature(
        _run(
            '<w:t xml:space="preserve">kept </w:t>'
            '<w:delText xml:space="preserve">gone </w:delText>'
            '<w:instrText xml:space="preserve"> PAGE </w:instrText>'
            "<w:delInstrText> NUMPAGES </w:delInstrText>"
        )
    )

    assert signature == (
        (
            ("t", "kept ", True),
            ("delText", "gone ", True),
            ("instr", " PAGE "),
            ("delInstr", " NUMPAGES "),
        ),
    )


def test_empty_text_nodes_and_empty_runs_are_recorded():
    assert _signature(_run("<w:t/>") + _run("<w:t></w:t>") + "<w:r/>" + _run("")) == (
        (("t", "", False),),
        (("t", "", False),),
        (),
        (),
    )


# ── structural children ────────────────────────────────────────────────────


@pytest.mark.parametrize(
    "local_name",
    [
        "tab",
        "cr",
        "noBreakHyphen",
        "softHyphen",
        "separator",
        "continuationSeparator",
        "footnoteRef",
        "endnoteRef",
        "annotationRef",
        "dayShort",
        "dayLong",
        "monthShort",
        "monthLong",
        "yearShort",
        "yearLong",
        "pgNum",
    ],
)
def test_content_free_children_are_recorded_by_name(local_name):
    empty = _signature(_run(f"<w:{local_name}/>"))
    paired = _signature(_run(f"<w:{local_name}></w:{local_name}>"))

    assert empty == paired == (((local_name,),),)


def test_attribute_bearing_children_record_their_listed_attributes():
    signature = _signature(
        _run(
            '<w:ptab w:alignment="right" w:relativeTo="margin" w:leader="dot"/>'
            "<w:br/>"
            '<w:br w:type="page"/>'
            "<w:br w:type='textWrapping' w:clear='all'/>"
            '<w:sym w:font="Wingdings" w:char="F0FC"/>'
            '<w:fldChar w:fldCharType="begin" w:dirty="true"><w:ffData>'
            '<w:name w:val="Check1"/></w:ffData></w:fldChar>'
        )
    )

    assert signature == (
        (
            ("ptab", "right", "margin", "dot"),
            ("br", "", ""),
            ("br", "page", ""),
            ("br", "textWrapping", "all"),
            ("sym", "Wingdings", "F0FC"),
            ("fldChar", "begin"),
        ),
    )


def test_references_record_their_kind_and_id():
    signature = _signature(
        _run('<w:footnoteReference w:id="2"/>')
        + _run('<w:endnoteReference w:id="3"/>')
        + _run('<w:commentReference w:id="4"/>')
    )

    assert signature == (
        (("ref", "footnoteReference", "2"),),
        (("ref", "endnoteReference", "3"),),
        (("ref", "commentReference", "4"),),
    )


def test_run_properties_and_rendering_hints_are_not_content():
    signature = _signature(
        _run(
            '<w:rPr><w:b/><w:rFonts w:ascii="Arial"/></w:rPr>'
            "<w:lastRenderedPageBreak/><w:t>x</w:t>"
        )
    )

    assert signature == ((("t", "x", False),),)
    assert signature == _signature(_run("<w:t>x</w:t>"))


def test_an_unknown_child_is_a_difference_rather_than_silence():
    signature = _signature(
        _run('<w:contentPart r:id="rId9"/>')
        + _run('<w14:unknownThing w14:val="1"/>')
    )

    assert signature == (
        (("other", "w:contentPart"),),
        (("other", "w14:unknownThing"),),
    )


# ── which runs are walked ──────────────────────────────────────────────────


@pytest.mark.parametrize(
    ("opening", "closing"),
    [
        ('<w:ins w:id="1" w:author="A" w:date="2026-01-01T00:00:00Z">', "</w:ins>"),
        ('<w:del w:id="2" w:author="A" w:date="2026-01-01T00:00:00Z">', "</w:del>"),
        ('<w:moveTo w:id="3" w:author="A">', "</w:moveTo>"),
        ('<w:moveFrom w:id="4" w:author="A">', "</w:moveFrom>"),
        ('<w:hyperlink r:id="rId5">', "</w:hyperlink>"),
        ('<w:smartTag w:element="place"><w:smartTagPr/>', "</w:smartTag>"),
        (
            "<w:sdt><w:sdtPr><w:rPr><w:b/></w:rPr><w:alias w:val='Field'/></w:sdtPr>"
            "<w:sdtContent>",
            "</w:sdtContent></w:sdt>",
        ),
        ('<w:customXml w:element="item"><w:customXmlPr/>', "</w:customXml>"),
        ('<w:fldSimple w:instr=" PAGE ">', "</w:fldSimple>"),
        ('<w:dir w:val="rtl">', "</w:dir>"),
        ('<w:bdo w:val="ltr">', "</w:bdo>"),
        (
            '<w:ins w:id="6" w:author="A"><w:hyperlink w:anchor="x"><w:smartTag w:element="y">',
            "</w:smartTag></w:hyperlink></w:ins>",
        ),
    ],
)
def test_runs_are_walked_inside_every_run_container(opening, closing):
    signature = _signature(
        _run("<w:t>before</w:t>")
        + opening
        + _run("<w:t>inside</w:t><w:tab/>")
        + closing
        + _run("<w:t>after</w:t>")
    )

    assert signature == (
        (("t", "before", False),),
        (("t", "inside", False), ("tab",)),
        (("t", "after", False),),
    )


def test_runs_nested_in_ruby_follow_the_run_that_holds_them():
    signature = _signature(
        _run(
            "<w:ruby><w:rubyPr><w:rubyAlign w:val='center'/></w:rubyPr>"
            "<w:rt><w:r><w:t>guide</w:t></w:r></w:rt>"
            "<w:rubyBase><w:r><w:t>base</w:t></w:r></w:rubyBase></w:ruby>"
            "<w:t>tail</w:t>"
        )
    )

    assert signature == (
        (("other", "w:ruby"), ("t", "tail", False)),
        (("t", "guide", False),),
        (("t", "base", False),),
    )


def test_tab_stop_definitions_are_not_run_tabs():
    # w:pPr/w:tabs/w:tab uses the run tab's element name. Format-only may
    # strip a direct w:tabs the architect style supplies, so counting it here
    # would fail every such paragraph.
    runs = _run("<w:t>Title</w:t><w:tab/><w:t>Page</w:t>")
    with_stops = (
        "<w:pPr><w:tabs><w:tab w:val='left' w:pos='720'/>"
        '<w:tab w:val="right" w:leader="dot" w:pos="9360"/></w:tabs>'
        '<w:pPrChange w:id="9" w:author="A"><w:pPr><w:tabs>'
        '<w:tab w:val="center" w:pos="4680"/></w:tabs></w:pPr></w:pPrChange>'
        "</w:pPr>"
    )

    assert _signature(with_stops + runs) == _signature(runs) == (
        (("t", "Title", False), ("tab",), ("t", "Page", False)),
    )


def test_out_of_scope_subtrees_are_excluded():
    # Drawings and text boxes are compared byte-for-byte elsewhere; their
    # paragraphs are not the host paragraph's run content.
    signature = _signature(
        _run(
            "<mc:AlternateContent><mc:Choice Requires='wps'><w:drawing>"
            "<wps:txbx><w:txbxContent><w:p><w:r><w:t>box</w:t></w:r></w:p>"
            "</w:txbxContent></wps:txbx></w:drawing></mc:Choice>"
            "<mc:Fallback><w:pict><v:textbox><w:txbxContent><w:p><w:r>"
            "<w:t>box</w:t></w:r></w:p></w:txbxContent></v:textbox></w:pict>"
            "</mc:Fallback></mc:AlternateContent><w:t>host</w:t>"
        )
        + _run('<w:object w:dxaOrig="10"><w:t>ole</w:t></w:object>')
    )

    assert signature == (
        (("other", "mc:AlternateContent"), ("t", "host", False)),
        (),
    )


# ── reporting a difference ─────────────────────────────────────────────────

# Each row is a mutation of the Format-only gate probe
# (docs/docx_method_hardening/probes/probe_format_only_gate.py): identical by
# normalized text, different by run content.
PROBE_MUTATIONS = [
    pytest.param(
        _run('<w:t xml:space="preserve">Section </w:t><w:tab/><w:t>Title</w:t>'),
        _run('<w:t xml:space="preserve">Section </w:t><w:t>Title</w:t>'),
        "tab",
        id="tab-dropped-beside-a-space",
    ),
    pytest.param(
        _run('<w:t xml:space="preserve">Old </w:t>')
        + '<w:del w:id="1" w:author="A" w:date="2026-01-01T00:00:00Z">'
        + _run("<w:delText>deleted words</w:delText>")
        + "</w:del>"
        + _run("<w:t>kept</w:t>"),
        _run('<w:t xml:space="preserve">Old </w:t>')
        + '<w:del w:id="1" w:author="A" w:date="2026-01-01T00:00:00Z">'
        + _run("<w:delText></w:delText>")
        + "</w:del>"
        + _run("<w:t>kept</w:t>"),
        "deleted_text",
        id="deleted-text-emptied",
    ),
    pytest.param(
        _run('<w:t xml:space="preserve">Trailing </w:t>') + _run("<w:t>text</w:t>"),
        _run("<w:t>Trailing </w:t>") + _run("<w:t>text</w:t>"),
        "preserve_space",
        id="preserve-removed",
    ),
    pytest.param(
        _run("<w:t>SECTION 21 13 13</w:t>"),
        _run("<w:t>SECTION 21 13 13</w:t>"),
        "text",
        id="non-breaking-spaces-made-plain",
    ),
    pytest.param(
        _run("<w:t>Double  space</w:t>"),
        _run("<w:t>Double space</w:t>"),
        "text",
        id="double-space-collapsed",
    ),
    pytest.param(
        _run("<w:t>Fire</w:t><w:softHyphen/><w:t>proofing</w:t>"),
        _run("<w:t>Fire</w:t><w:t>proofing</w:t>"),
        "soft_hyphen",
        id="soft-hyphen-dropped",
    ),
    pytest.param(
        _run('<w:t xml:space="preserve">line one </w:t><w:br/><w:t>line two</w:t>'),
        _run('<w:t xml:space="preserve">line one </w:t><w:t>line two</w:t>'),
        "break",
        id="break-dropped-beside-a-space",
    ),
]


@pytest.mark.parametrize(("source", "output", "kind"), PROBE_MUTATIONS)
def test_signature_sees_what_normalized_text_cannot(source, output, kind):
    source_paragraph = f"<w:p>{source}</w:p>"
    output_paragraph = f"<w:p>{output}</w:p>"
    assert paragraph_text_from_block(source_paragraph) == paragraph_text_from_block(
        output_paragraph
    )

    before = paragraph_run_content_signature(source_paragraph)
    after = paragraph_run_content_signature(output_paragraph)

    assert before != after
    assert run_content_difference(before, after) == kind


@pytest.mark.parametrize(
    ("source", "output", "kind"),
    [
        (_run("<w:t>shall</w:t>"), _run("<w:t>should</w:t>"), "text"),
        (
            _run('<w:delText xml:space="preserve">gone </w:delText>'),
            _run("<w:delText>gone </w:delText>"),
            "preserve_space",
        ),
        (
            _run("<w:instrText> PAGE </w:instrText>"),
            _run("<w:instrText> NUMPAGES </w:instrText>"),
            "field_instruction",
        ),
        (
            _run("<w:delInstrText> PAGE </w:delInstrText>"),
            _run("<w:delInstrText>PAGE</w:delInstrText>"),
            "deleted_field_instruction",
        ),
        (
            _run("<w:t>a</w:t><w:t>b</w:t>"),
            _run("<w:t>a</w:t><w:tab/><w:t>b</w:t>"),
            "tab",
        ),
        (
            _run('<w:ptab w:alignment="right" w:relativeTo="margin" w:leader="dot"/>'),
            _run('<w:ptab w:alignment="right" w:relativeTo="margin" w:leader="none"/>'),
            "positional_tab",
        ),
        (_run('<w:br w:type="page"/>'), _run("<w:br/>"), "break"),
        (_run("<w:t>a</w:t><w:cr/>"), _run("<w:t>a</w:t>"), "carriage_return"),
        (
            _run("<w:t>a</w:t><w:noBreakHyphen/><w:t>b</w:t>"),
            _run("<w:t>a</w:t><w:t>b</w:t>"),
            "non_breaking_hyphen",
        ),
        (
            _run('<w:sym w:font="Symbol" w:char="F0B0"/>'),
            _run('<w:sym w:font="Symbol" w:char="F0B1"/>'),
            "symbol",
        ),
        (
            _run('<w:fldChar w:fldCharType="separate"/>'),
            _run('<w:fldChar w:fldCharType="end"/>'),
            "field_character",
        ),
        (
            _run('<w:footnoteReference w:id="2"/>'),
            _run('<w:footnoteReference w:id="3"/>'),
            "footnote_reference",
        ),
        (
            _run('<w:endnoteReference w:id="2"/>'),
            _run(""),
            "endnote_reference",
        ),
        (
            _run('<w:commentReference w:id="7"/>'),
            _run('<w:commentReference w:id="8"/>'),
            "comment_reference",
        ),
        (_run("<w:separator/>"), _run("<w:continuationSeparator/>"), "separator"),
        (_run("<w:footnoteRef/>"), _run(""), "footnote_ref"),
        (_run("<w:t>p</w:t><w:pgNum/>"), _run("<w:t>p</w:t>"), "page_number"),
        (_run("<w:dayLong/>"), _run("<w:dayShort/>"), "day_long"),
        (
            _run('<w:contentPart r:id="rId9"/>'),
            _run(""),
            "other_run_content",
        ),
        # The same content, regrouped: a run split in two, or an empty run
        # added, changes nothing a text node records.
        (_run("<w:t>a</w:t><w:t>b</w:t>"), _run("<w:t>a</w:t>") + _run("<w:t>b</w:t>"), "run_boundary"),
        (_run("<w:t>a</w:t>"), _run("<w:t>a</w:t>") + "<w:r/>", "run_boundary"),
    ],
)
def test_difference_names_the_kind_of_item_that_changed(source, output, kind):
    before = _signature(source)
    after = _signature(output)

    assert run_content_difference(before, after) == kind
    assert kind in RUN_CONTENT_DIFFERENCE_KINDS


def test_a_removed_item_is_named_rather_than_the_item_that_slid_into_its_place():
    before = _signature(_run("<w:t>a</w:t><w:br/><w:t>b</w:t><w:tab/><w:t>c</w:t>"))
    after = _signature(_run("<w:t>a</w:t><w:t>b</w:t><w:tab/><w:t>c</w:t>"))

    assert run_content_difference(before, after) == "break"


def test_identical_signatures_report_no_difference():
    signature = _signature(
        _run('<w:t xml:space="preserve">a </w:t><w:tab/>')
        + '<w:del w:id="1" w:author="A">'
        + _run("<w:delText>b</w:delText>")
        + "</w:del>"
    )

    assert run_content_difference(signature, signature) is None
    assert run_content_difference((), ()) is None


def test_the_reported_kind_is_a_closed_identifier_never_document_text():
    before = _signature(_run("<w:t>Confidential requirement wording</w:t>"))
    after = _signature(_run("<w:t>Confidential requirement wordin</w:t>"))

    kind = run_content_difference(before, after)

    assert kind == "text"
    assert all(kind.replace("_", "").isalpha() for kind in RUN_CONTENT_DIFFERENCE_KINDS)
    assert "Confidential" not in "".join(RUN_CONTENT_DIFFERENCE_KINDS)
