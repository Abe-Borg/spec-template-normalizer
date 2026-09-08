"""Prefix-stable OOXML serializers under concurrent use."""

from __future__ import annotations

import xml.etree.ElementTree as ET
from concurrent.futures import ThreadPoolExecutor

from spec_formatter.style_application.core.ooxml_namespaces import (
    CT_NS,
    PKG_REL_NS,
    W_NS,
    serialize_content_types,
    serialize_package_relationships,
    serialize_wordprocessingml,
)


def _content_types_root() -> ET.Element:
    root = ET.Element(f"{{{CT_NS}}}Types")
    ET.SubElement(
        root,
        f"{{{CT_NS}}}Override",
        PartName="/word/theme/theme1.xml",
        ContentType="application/vnd.openxmlformats-officedocument.theme+xml",
    )
    return root


def _relationships_root() -> ET.Element:
    root = ET.Element(f"{{{PKG_REL_NS}}}Relationships")
    ET.SubElement(root, f"{{{PKG_REL_NS}}}Relationship", Id="rId1", Target="styles.xml")
    return root


def test_content_types_serialize_with_a_default_namespace():
    xml = serialize_content_types(_content_types_root()).decode("utf-8")
    assert f'<Types xmlns="{CT_NS}">' in xml
    assert "<Override " in xml
    assert "ns0" not in xml


def test_package_relationships_serialize_with_a_default_namespace():
    xml = serialize_package_relationships(_relationships_root()).decode("utf-8")
    assert f'<Relationships xmlns="{PKG_REL_NS}">' in xml
    assert "<Relationship " in xml
    assert "ns0" not in xml


def test_wordprocessingml_serializes_with_the_w_prefix():
    root = ET.Element(f"{{{W_NS}}}styles")
    ET.SubElement(root, f"{{{W_NS}}}style")
    xml = serialize_wordprocessingml(root).decode("utf-8")
    assert f'<w:styles xmlns:w="{W_NS}">' in xml
    assert "<w:style />" in xml or "<w:style/>" in xml
    assert "ns0" not in xml


def test_concurrent_serializers_never_swap_prefixes():
    # The default prefix can map to one URI at a time in ElementTree's global
    # table, so unlocked serializers racing on a thread pool used to emit
    # ns0-prefixed relationships or content types.
    def content_types(_index: int) -> bytes:
        return serialize_content_types(_content_types_root())

    def relationships(_index: int) -> bytes:
        return serialize_package_relationships(_relationships_root())

    with ThreadPoolExecutor(max_workers=8) as pool:
        ct_outputs = list(pool.map(content_types, range(200)))
        rel_outputs = list(pool.map(relationships, range(200)))
        mixed = list(
            pool.map(
                lambda index: (content_types if index % 2 else relationships)(index),
                range(200),
            )
        )

    for xml_bytes in ct_outputs:
        assert f'<Types xmlns="{CT_NS}">'.encode() in xml_bytes
    for xml_bytes in rel_outputs:
        assert f'<Relationships xmlns="{PKG_REL_NS}">'.encode() in xml_bytes
    for xml_bytes in mixed:
        assert b"ns0" not in xml_bytes
        assert (
            f'<Types xmlns="{CT_NS}">'.encode() in xml_bytes
            or f'<Relationships xmlns="{PKG_REL_NS}">'.encode() in xml_bytes
        )
