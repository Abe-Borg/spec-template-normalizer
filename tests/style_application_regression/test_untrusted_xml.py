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
