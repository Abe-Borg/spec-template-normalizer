from __future__ import annotations

import xml.etree.ElementTree as ET

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
PKG_REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
CT_NS = "http://schemas.openxmlformats.org/package/2006/content-types"


def register_ooxml_namespaces() -> None:
    ET.register_namespace("w", W_NS)
    ET.register_namespace("r", R_NS)
    ET.register_namespace("", PKG_REL_NS)


def serialize_wordprocessingml(root: ET.Element) -> bytes:
    register_ooxml_namespaces()
    xml_bytes = ET.tostring(root, encoding="utf-8", xml_declaration=True)
    if b"<ns0:" in xml_bytes or b"</ns0:" in xml_bytes:
        raise RuntimeError("Unexpected anonymous namespace prefix in WordprocessingML output")
    return xml_bytes


def serialize_package_relationships(root: ET.Element) -> bytes:
    register_ooxml_namespaces()
    return ET.tostring(root, encoding="utf-8", xml_declaration=True)


def serialize_content_types(root: ET.Element) -> bytes:
    register_ooxml_namespaces()
    return ET.tostring(root, encoding="utf-8", xml_declaration=True)
