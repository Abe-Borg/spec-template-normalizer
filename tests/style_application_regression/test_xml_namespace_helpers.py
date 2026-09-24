"""The lexical namespace helpers that let fragments move between parts.

A fragment taken out of one part and written into another keeps its prefixes
but not the declarations that gave them meaning. These helpers read the
declarations on a part's root, find the prefixes a fragment needs from its
surroundings, and add the missing ones to a destination root -- refusing,
rather than guessing, when the destination already means something else by a
prefix.
"""

from __future__ import annotations

import xml.etree.ElementTree as ET

import pytest

from spec_formatter.style_application.core.xml_helpers import (
    MCE_NS,
    NamespaceReconciliationError,
    RootNamespaces,
    declare_fragment_namespaces,
    ensure_root_declarations,
    prefixes_used,
    root_ignorable_prefixes,
    root_namespace_additions,
    root_namespace_declarations,
    root_opening_tag,
    xml_unescape,
)

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
W14_NS = "http://schemas.microsoft.com/office/word/2010/wordml"
W15_NS = "http://schemas.microsoft.com/office/word/2012/wordml"
FOREIGN_NS = "urn:spec-template-normalizer:test:not-word-2010"


# ── root_namespace_declarations ────────────────────────────────────────────


def test_root_declarations_skip_the_prolog_and_read_either_quoting():
    xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        "<!-- <x:decoy xmlns:x='urn:decoy'/> -->"
        f"<w:styles xmlns:w=\"{W_NS}\" xmlns:w14 = '{W14_NS}' xmlns=\"urn:default\" "
        f'xmlns:mc="{MCE_NS}" mc:Ignorable="w14">'
        '<w:style xmlns:nested="urn:nested"/></w:styles>'
    )

    assert root_namespace_declarations(xml) == {
        "w": W_NS,
        "w14": W14_NS,
        "": "urn:default",
        "mc": MCE_NS,
    }


def test_root_declarations_decode_character_references():
    xml = '<w:styles xmlns:w="urn:a&amp;b&#x26;c" xmlns:x="urn:&#65;"/>'

    assert root_namespace_declarations(xml) == {"w": "urn:a&b&c", "x": "urn:A"}


@pytest.mark.parametrize("xml", ["", "   ", "<!-- only a comment -->", "</w:styles>"])
def test_a_part_without_a_root_element_is_refused(xml):
    with pytest.raises(ValueError):
        root_namespace_declarations(xml)


# ── root_ignorable_prefixes ────────────────────────────────────────────────


def test_ignorable_prefixes_are_read_from_the_mce_attribute():
    xml = (
        f'<w:styles xmlns:w="{W_NS}" xmlns:mc="{MCE_NS}" '
        'mc:Ignorable="w14  w15\twp14"/>'
    )

    assert root_ignorable_prefixes(xml) == frozenset({"w14", "w15", "wp14"})


def test_ignorable_prefixes_follow_the_namespace_not_the_literal_prefix():
    word_2007 = f'<w:styles xmlns:w="{W_NS}" xmlns:ve="{MCE_NS}" ve:Ignorable="w14"/>'
    impostor = f'<w:styles xmlns:w="{W_NS}" xmlns:mc="{FOREIGN_NS}" mc:Ignorable="w14"/>'

    assert root_ignorable_prefixes(word_2007) == frozenset({"w14"})
    assert root_ignorable_prefixes(impostor) == frozenset()
    assert root_ignorable_prefixes(f'<w:styles xmlns:w="{W_NS}"/>') == frozenset()


def test_root_namespaces_bundles_declarations_and_ignorable_prefixes():
    xml = (
        f'<w:styles xmlns:w="{W_NS}" xmlns:w14="{W14_NS}" xmlns:mc="{MCE_NS}" '
        'mc:Ignorable="w14"/>'
    )

    namespaces = RootNamespaces.of(xml)

    assert dict(namespaces.declarations) == {"w": W_NS, "w14": W14_NS, "mc": MCE_NS}
    assert namespaces.ignorable == frozenset({"w14"})


# ── prefixes_used ──────────────────────────────────────────────────────────


def test_prefixes_used_reads_element_and_attribute_names_but_not_xml_or_xmlns():
    fragment = (
        '<w:rPr><w14:ligatures w14:val="standardContextual"/>'
        '<w:b w15:note="1"/><w:t xml:space="preserve">a</w:t></w:rPr>'
    )

    assert prefixes_used(fragment) == {"w", "w14", "w15"}


def test_a_prefix_declared_inside_the_fragment_is_needed_only_outside_its_scope():
    in_scope = '<x:a xmlns:x="urn:x"><x:b x:c="1"/></x:a>'
    out_of_scope = '<w:r><x:a xmlns:x="urn:x"/><x:b/></w:r>'

    assert prefixes_used(in_scope) == set()
    assert prefixes_used(out_of_scope) == {"w", "x"}


def test_unprefixed_elements_need_the_default_namespace():
    assert prefixes_used("<a><b/></a>") == {""}
    assert prefixes_used('<a xmlns="urn:x"><b/></a>') == set()
    # An unprefixed attribute is in no namespace and needs nothing.
    assert prefixes_used('<w:b val="1"/>') == {"w"}


def test_text_that_only_looks_like_markup_is_not_read():
    fragment = (
        "<w:r><!-- <x:y/> --><?pi <p:q/> ?>"
        '<w:t w:val="a:b > c:d"><![CDATA[<z:q/>]]> e:f </w:t></w:r>'
    )

    assert prefixes_used(fragment) == {"w"}


# ── ensure_root_declarations ───────────────────────────────────────────────

TARGET = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
    f'<w:styles xmlns:w="{W_NS}"><w:style w:styleId="Normal"/></w:styles>'
)


def test_missing_declarations_are_added_to_the_root_and_nothing_else_changes():
    result = ensure_root_declarations(TARGET, {"w": W_NS, "w14": W14_NS})

    assert result == TARGET.replace(
        f'<w:styles xmlns:w="{W_NS}">',
        f'<w:styles xmlns:w="{W_NS}" xmlns:w14="{W14_NS}">',
    )
    assert root_namespace_declarations(result)["w14"] == W14_NS
    ET.fromstring(result.encode("utf-8"))


def test_nothing_to_add_returns_the_part_unchanged():
    assert ensure_root_declarations(TARGET, {"w": W_NS}) == TARGET
    assert ensure_root_declarations(TARGET, {}) == TARGET


def test_a_prefix_bound_to_another_namespace_is_a_conflict():
    target = TARGET.replace(
        f'xmlns:w="{W_NS}"', f'xmlns:w="{W_NS}" xmlns:w14="{FOREIGN_NS}"'
    )

    with pytest.raises(NamespaceReconciliationError) as raised:
        ensure_root_declarations(target, {"w": W_NS, "w14": W14_NS})

    assert isinstance(raised.value, ValueError)
    assert raised.value.prefix == "w14"
    assert raised.value.reason == "conflict"


def test_the_default_namespace_is_compared_and_never_added():
    with_default = TARGET.replace(f'xmlns:w="{W_NS}"', f'xmlns:w="{W_NS}" xmlns="urn:d"')

    # Unprefixed names in no namespace mean the same in a root without a
    # default declaration.
    assert ensure_root_declarations(TARGET, {"": ""}) == TARGET
    assert ensure_root_declarations(with_default, {"": "urn:d"}) == with_default
    for part, needed in ((with_default, ""), (TARGET, "urn:d")):
        with pytest.raises(NamespaceReconciliationError) as raised:
            ensure_root_declarations(part, {"": needed})
        assert raised.value.prefix == ""


def test_an_ignorable_prefix_gets_a_markup_compatibility_declaration_and_token():
    result = ensure_root_declarations(TARGET, {"w14": W14_NS}, ignorable={"w14"})

    assert (
        f'<w:styles xmlns:w="{W_NS}" xmlns:w14="{W14_NS}" xmlns:mc="{MCE_NS}" '
        'mc:Ignorable="w14">'
    ) in result
    assert root_ignorable_prefixes(result) == frozenset({"w14"})
    ET.fromstring(result.encode("utf-8"))


def test_an_existing_token_is_not_repeated_and_new_tokens_are_appended():
    target = TARGET.replace(
        f'xmlns:w="{W_NS}"',
        f"xmlns:w=\"{W_NS}\" xmlns:mc=\"{MCE_NS}\" xmlns:w14=\"{W14_NS}\" "
        f"xmlns:w15=\"{W15_NS}\" mc:Ignorable='w14'",
    )

    unchanged = ensure_root_declarations(target, {"w14": W14_NS}, ignorable={"w14"})
    extended = ensure_root_declarations(
        target, {"w14": W14_NS, "w15": W15_NS}, ignorable={"w14", "w15"}
    )

    assert unchanged == target
    assert "mc:Ignorable='w14 w15'" in extended
    assert root_ignorable_prefixes(extended) == frozenset({"w14", "w15"})


def test_an_mc_prefix_bound_elsewhere_cannot_carry_ignorable_tokens():
    target = TARGET.replace(f'xmlns:w="{W_NS}"', f'xmlns:w="{W_NS}" xmlns:mc="{FOREIGN_NS}"')

    with pytest.raises(NamespaceReconciliationError) as raised:
        ensure_root_declarations(target, {"w14": W14_NS}, ignorable={"w14"})

    assert raised.value.prefix == "mc"
    assert raised.value.reason == "conflict"


def test_a_self_closing_root_and_escaped_uris_are_written_correctly():
    uri = 'urn:a&b"c<d'

    result = ensure_root_declarations(f'<w:styles xmlns:w="{W_NS}"/>', {"x": uri})

    assert result == f'<w:styles xmlns:w="{W_NS}" xmlns:x="urn:a&amp;b&quot;c&lt;d"/>'
    assert root_namespace_declarations(result)["x"] == uri


# ── declare_fragment_namespaces ────────────────────────────────────────────


def test_fragments_take_their_prefixes_meaning_from_the_source_root():
    source = RootNamespaces.of(
        f'<w:styles xmlns:w="{W_NS}" xmlns:w14="{W14_NS}" xmlns:w15="{W15_NS}" '
        f'xmlns:mc="{MCE_NS}" mc:Ignorable="w14 w15"/>'
    )
    fragments = ['<w:rPr><w14:ligatures w14:val="standard"/></w:rPr>', "<w:b/>"]

    result = declare_fragment_namespaces(TARGET, fragments, source)

    # w15 is ignorable in the source but no fragment uses it.
    assert root_namespace_declarations(result) == {"w": W_NS, "w14": W14_NS, "mc": MCE_NS}
    assert root_ignorable_prefixes(result) == frozenset({"w14"})


def test_a_prefix_the_source_root_does_not_declare_is_refused():
    source = RootNamespaces.of(f'<w:styles xmlns:w="{W_NS}"/>')

    with pytest.raises(NamespaceReconciliationError) as raised:
        declare_fragment_namespaces(TARGET, ['<w14:ligatures w14:val="x"/>'], source)

    assert raised.value.prefix == "w14"
    assert raised.value.reason == "undeclared"


# ── small readers ──────────────────────────────────────────────────────────


def test_root_opening_tag_is_returned_exactly_as_written():
    xml = f"<?xml version='1.0'?>\n<w:styles xmlns:w='{W_NS}' a=\"x > y\"><w:style/></w:styles>"

    assert root_opening_tag(xml) == f"<w:styles xmlns:w='{W_NS}' a=\"x > y\">"


def test_additions_report_new_declarations_and_ignorable_tokens_only():
    before = f'<w:styles xmlns:w="{W_NS}" xmlns:mc="{MCE_NS}" mc:Ignorable="w15"/>'
    after = ensure_root_declarations(
        before, {"w14": W14_NS, "w15": W15_NS}, ignorable={"w14", "w15"}
    )

    assert root_namespace_additions(before, after) == ("w14", "w15", "mc:Ignorable=w14")
    assert root_namespace_additions(after, after) == ()


def test_xml_unescape_expands_xml_references_and_nothing_else():
    assert xml_unescape("&lt;&gt;&amp;&quot;&apos;&#65;&#x42;") == "<>&\"'AB"
    # HTML entities are not XML; they are left exactly as written.
    assert xml_unescape("&nbsp;&copy;") == "&nbsp;&copy;"

