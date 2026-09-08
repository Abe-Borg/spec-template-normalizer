"""DOCTYPE/ENTITY rejection for every untrusted XML parse in the target engine."""

from __future__ import annotations

import zipfile
from pathlib import Path

import pytest

from spec_formatter.style_application.core.classification import _build_numbering_catalog
from spec_formatter.style_application.core.untrusted_xml import (
    UntrustedXmlError,
    parse_untrusted_xml,
)
from spec_formatter.style_application.docx_patch import validate_xml_wellformedness
from spec_formatter.style_application.phase2_invariants import validate_docx_package


W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
PKG_REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"

# A small internal-entity payload: expat would expand it into 10**6 copies of
# "lol" if the declaration were ever allowed to reach the parser.
ENTITY_DOCTYPE = (
    "<!DOCTYPE lolz [<!ENTITY lol \"lol\">"
    "<!ENTITY lol1 \"&lol;&lol;&lol;&lol;&lol;&lol;&lol;&lol;&lol;&lol;\">"
    "<!ENTITY lol2 \"&lol1;&lol1;&lol1;&lol1;&lol1;&lol1;&lol1;&lol1;&lol1;&lol1;\">"
    "<!ENTITY lol3 \"&lol2;&lol2;&lol2;&lol2;&lol2;&lol2;&lol2;&lol2;&lol2;&lol2;\">"
    "<!ENTITY lol4 \"&lol3;&lol3;&lol3;&lol3;&lol3;&lol3;&lol3;&lol3;&lol3;&lol3;\">"
    "<!ENTITY lol5 \"&lol4;&lol4;&lol4;&lol4;&lol4;&lol4;&lol4;&lol4;&lol4;&lol4;\">"
    "<!ENTITY lol6 \"&lol5;&lol5;&lol5;&lol5;&lol5;&lol5;&lol5;&lol5;&lol5;&lol5;\">"
    "]>"
)


def _document_with_entities() -> str:
    return (
        '<?xml version="1.0" encoding="UTF-8"?>'
        f"{ENTITY_DOCTYPE}"
        f'<w:document xmlns:w="{W_NS}"><w:body><w:p><w:r><w:t>&lol6;</w:t></w:r>'
        "</w:p></w:body></w:document>"
    )


def test_parse_untrusted_xml_returns_the_root_for_clean_input():
    root = parse_untrusted_xml(f'<w:styles xmlns:w="{W_NS}"/>', "word/styles.xml")
    assert root.tag == f"{{{W_NS}}}styles"
    root = parse_untrusted_xml(b"<Types/>", "[Content_Types].xml")
    assert root.tag == "Types"


@pytest.mark.parametrize(
    "payload",
    [
        _document_with_entities(),
        _document_with_entities().encode("utf-8"),
        "<!doctype html><w:p/>",
        "\n\n<!-- prolog comment -->\n" + ENTITY_DOCTYPE + "<w:p/>",
        # Any ENTITY declaration is rejected even outside a leading DOCTYPE.
        "<w:p>" + "x" * 2048 + "<!ENTITY a \"b\"></w:p>",
    ],
)
def test_doctype_and_entity_declarations_are_rejected_before_parsing(payload):
    with pytest.raises(UntrustedXmlError, match="word/document.xml: DOCTYPE/ENTITY"):
        parse_untrusted_xml(payload, "word/document.xml")


def test_parse_errors_are_wrapped_with_the_part_name():
    with pytest.raises(UntrustedXmlError, match="word/numbering.xml: XML parse error") as raised:
        parse_untrusted_xml("<w:numbering><w:num></w:numbering>", "word/numbering.xml")
    assert isinstance(raised.value, ValueError)


def test_numbering_catalog_rejects_entity_doctype():
    numbering = (
        f"{ENTITY_DOCTYPE}<w:numbering xmlns:w=\"{W_NS}\">"
        '<w:abstractNum w:abstractNumId="1"><w:lvl w:ilvl="0">'
        "<w:lvlText w:val=\"&lol6;\"/></w:lvl></w:abstractNum></w:numbering>"
    )
    with pytest.raises(UntrustedXmlError, match="word/numbering.xml"):
        _build_numbering_catalog(numbering)


def test_replacement_wellformedness_check_names_the_doctype_part():
    errors = validate_xml_wellformedness(
        {"word/document.xml": _document_with_entities().encode("utf-8")}
    )
    assert errors == [
        "word/document.xml: DOCTYPE/ENTITY declarations are not allowed in an "
        "untrusted OOXML part"
    ]


def test_package_validation_rejects_doctype_in_content_types(tmp_path: Path):
    docx = tmp_path / "doctype.docx"
    parts = {
        "[Content_Types].xml": (
            '<?xml version="1.0" encoding="UTF-8"?>'
            "<!DOCTYPE Types [<!ENTITY x \"application/xml\">]>"
            '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
            '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
            '<Default Extension="xml" ContentType="&x;"/>'
            "</Types>"
        ),
        "_rels/.rels": (
            f'<Relationships xmlns="{PKG_REL_NS}">'
            f'<Relationship Id="rId1" Type="{R_NS}/officeDocument" Target="word/document.xml"/>'
            "</Relationships>"
        ),
        "word/document.xml": (
            f'<w:document xmlns:w="{W_NS}"><w:body><w:p/><w:sectPr/></w:body></w:document>'
        ),
        "word/_rels/document.xml.rels": (
            f'<Relationships xmlns="{PKG_REL_NS}">'
            f'<Relationship Id="rId1" Type="{R_NS}/styles" Target="styles.xml"/>'
            "</Relationships>"
        ),
        "word/styles.xml": f'<w:styles xmlns:w="{W_NS}"/>',
    }
    with zipfile.ZipFile(docx, "w") as zf:
        for name, value in parts.items():
            zf.writestr(name, value)

    with pytest.raises(Exception, match=r"\[Content_Types\]\.xml: DOCTYPE/ENTITY"):
        validate_docx_package(docx)


# --- Encoding-independent rejection (W1) -----------------------------------
#
# The byte scan matches ASCII, so a UTF-16 part could carry `<!DOCTYPE`
# past it and reach expat with entities intact. These cover every encoding
# the shared reader claims to support, in both directions: a prohibited
# declaration is always rejected, and valid content always survives.

TINY_ENTITY_DOC = '<!DOCTYPE r [<!ENTITY a "AAAA">]><r>&a;</r>'


def _declared(encoding: str, body: str = '<r a="café">naïve — éà</r>') -> str:
    return f'<?xml version="1.0" encoding="{encoding}"?>{body}'


@pytest.mark.parametrize(
    "label,payload",
    [
        ("utf-8", _declared("UTF-8").encode("utf-8")),
        ("utf-8 with BOM", b"\xef\xbb\xbf" + _declared("UTF-8").encode("utf-8")),
        ("utf-16 with BOM", _declared("UTF-16").encode("utf-16")),
        ("utf-16-le with BOM", b"\xff\xfe" + _declared("UTF-16").encode("utf-16-le")),
        ("utf-16-be with BOM", b"\xfe\xff" + _declared("UTF-16").encode("utf-16-be")),
        ("bom-less utf-16-le", _declared("UTF-16").encode("utf-16-le")),
        ("bom-less utf-16-be", _declared("UTF-16").encode("utf-16-be")),
        ("declared windows-1252", _declared("windows-1252").encode("cp1252")),
        ("str", _declared("UTF-8")),
        ("str declaring windows-1252", _declared("windows-1252")),
        ("str declaring utf-16", _declared("UTF-16")),
        ("str with no declaration", '<r a="café">naïve — éà</r>'),
    ],
)
def test_valid_content_keeps_its_characters_across_encodings(label, payload):
    root = parse_untrusted_xml(payload, "word/document.xml")
    assert root.tag == "r"
    assert root.attrib["a"] == "café", label
    assert root.text == "naïve — éà", label


@pytest.mark.parametrize(
    "label,payload",
    [
        ("utf-8", TINY_ENTITY_DOC.encode("utf-8")),
        ("utf-8 with BOM", b"\xef\xbb\xbf" + TINY_ENTITY_DOC.encode("utf-8")),
        ("utf-16 with BOM", TINY_ENTITY_DOC.encode("utf-16")),
        ("utf-16-be with BOM", b"\xfe\xff" + TINY_ENTITY_DOC.encode("utf-16-be")),
        (
            "utf-16 with BOM and declaration",
            ('<?xml version="1.0" encoding="UTF-16"?>' + TINY_ENTITY_DOC).encode("utf-16"),
        ),
        ("bom-less utf-16-le", TINY_ENTITY_DOC.encode("utf-16-le")),
        ("bom-less utf-16-be", TINY_ENTITY_DOC.encode("utf-16-be")),
        ("str", TINY_ENTITY_DOC),
    ],
)
def test_doctype_is_rejected_in_every_supported_encoding(label, payload):
    with pytest.raises(UntrustedXmlError, match="DOCTYPE/ENTITY") as raised:
        parse_untrusted_xml(payload, "word/document.xml")
    assert "word/document.xml" in str(raised.value), label


def test_utf16_entity_is_rejected_before_expansion(monkeypatch):
    """The contract is "before expansion", so prove the payload never parses."""
    import xml.etree.ElementTree as element_tree
    from spec_formatter.style_application.core import untrusted_xml as module

    calls = []
    monkeypatch.setattr(
        module.ET,
        "fromstring",
        lambda payload: calls.append(payload) or element_tree.fromstring(payload),
    )
    with pytest.raises(UntrustedXmlError, match="DOCTYPE/ENTITY"):
        parse_untrusted_xml(TINY_ENTITY_DOC.encode("utf-16"), "word/document.xml")
    assert calls == []


@pytest.mark.parametrize(
    "payload",
    [
        '<!DOCTYPE r SYSTEM "http://example.invalid/evil.dtd"><r/>',
        '<!DOCTYPE r PUBLIC "-//X//EN" "/etc/passwd"><r/>',
        '<!DOCTYPE r SYSTEM "file:///etc/passwd"><r/>',
        '<!DOCTYPE r SYSTEM "http://example.invalid/evil.dtd"><r/>'.encode("utf-16"),
    ],
)
def test_external_declarations_are_rejected_without_dereferencing(payload):
    with pytest.raises(UntrustedXmlError, match="DOCTYPE/ENTITY"):
        parse_untrusted_xml(payload, "word/document.xml")


@pytest.mark.parametrize(
    "payload",
    [
        # Declaration-shaped text the byte scan has always rejected, even
        # though XML would allow it here. Preserved deliberately.
        f'<w:p xmlns:w="{W_NS}"><!-- <!DOCTYPE evil> --></w:p>',
        f'<w:p xmlns:w="{W_NS}"><![CDATA[<!DOCTYPE evil>]]></w:p>',
        f'<w:p xmlns:w="{W_NS}"><![CDATA[<!ENTITY a "b">]]></w:p>',
    ],
)
def test_conservative_screening_of_declaration_shaped_text_is_preserved(payload):
    with pytest.raises(UntrustedXmlError, match="DOCTYPE/ENTITY"):
        parse_untrusted_xml(payload, "word/document.xml")


def test_escaped_declaration_text_is_still_ordinary_content():
    root = parse_untrusted_xml(
        f'<w:p xmlns:w="{W_NS}">&lt;!DOCTYPE evil&gt;</w:p>', "word/document.xml"
    )
    assert root.text == "<!DOCTYPE evil>"


@pytest.mark.parametrize(
    "label,payload",
    [
        ("truncated utf-16", "<r>abc".encode("utf-16")),
        ("literal NUL in content", b"<r>\x00</r>"),
        ("unknown declared encoding", b'<?xml version="1.0" encoding="nope-9000"?><r/>'),
        ("invalid utf-8 bytes", b"<r>\xff\xfe\xfa</r>"),
        ("empty", b""),
    ],
)
def test_malformed_payloads_fail_predictably(label, payload):
    with pytest.raises(UntrustedXmlError) as raised:
        parse_untrusted_xml(payload, "word/document.xml")
    assert isinstance(raised.value, ValueError), label
    assert "word/document.xml" in str(raised.value), label


def test_str_declaration_is_normalized_rather_than_reinterpreted():
    """A decoded str must not be read back through its stale declaration."""
    root = parse_untrusted_xml(_declared("windows-1252", "<x>é</x>"), "part.xml")
    assert root.text == "é"  # not "Ã©"
    root = parse_untrusted_xml(_declared("utf-16", "<x>é</x>"), "part.xml")
    assert root.text == "é"


def test_comments_namespaces_and_attributes_still_parse():
    root = parse_untrusted_xml(
        f'<?xml version="1.0"?><!-- lead --><w:p xmlns:w="{W_NS}" w:rsidR="00Aé">'
        f"<w:t>a &amp; b</w:t></w:p>",
        "word/document.xml",
    )
    assert root.tag == f"{{{W_NS}}}p"
    assert root.attrib[f"{{{W_NS}}}rsidR"] == "00Aé"
    assert root.find(f"{{{W_NS}}}t").text == "a & b"


@pytest.mark.parametrize(
    "label,payload",
    [
        # Both used to escape as bare LookupError / ValueError with no part
        # name, so a caller handling UntrustedXmlError never saw them.
        ("codec Python lacks", b'<?xml version="1.0" encoding="nope-9000"?><r/>'),
        ("multi-byte expat refuses", b'<?xml version="1.0" encoding="utf-7"?><r/>'),
    ],
)
def test_unsupported_encodings_are_wrapped_with_the_part_name(label, payload):
    with pytest.raises(UntrustedXmlError, match="unsupported XML encoding") as raised:
        parse_untrusted_xml(payload, "part.xml")
    assert isinstance(raised.value, ValueError), label
    assert "part.xml" in str(raised.value), label
