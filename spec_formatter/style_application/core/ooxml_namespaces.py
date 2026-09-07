"""OOXML namespace constants and prefix-stable ElementTree serializers.

``ET.register_namespace`` mutates a process-global table, and the default
(empty) prefix can map to only one URI at a time. Targets are formatted on a
thread pool, so two serializers racing on that table could swap prefixes and
emit ``ns0:``-prefixed parts. Every serializer here therefore registers its
own prefix map and calls ``tostring`` under one module lock, and every one
refuses to return output that fell back to an anonymous ``ns0`` prefix.
"""

from __future__ import annotations

import threading
import xml.etree.ElementTree as ET
from typing import Mapping

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
PKG_REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
CT_NS = "http://schemas.openxmlformats.org/package/2006/content-types"

_WORDPROCESSINGML_PREFIXES: Mapping[str, str] = {"w": W_NS, "r": R_NS}
_PACKAGE_RELATIONSHIPS_PREFIXES: Mapping[str, str] = {"": PKG_REL_NS}
_CONTENT_TYPES_PREFIXES: Mapping[str, str] = {"": CT_NS}

_SERIALIZE_LOCK = threading.Lock()


def register_ooxml_namespaces() -> None:
    """Register the WordprocessingML prefixes (kept for external callers)."""

    with _SERIALIZE_LOCK:
        for prefix, uri in _WORDPROCESSINGML_PREFIXES.items():
            ET.register_namespace(prefix, uri)


def _serialize(root: ET.Element, prefixes: Mapping[str, str], kind: str) -> bytes:
    with _SERIALIZE_LOCK:
        for prefix, uri in prefixes.items():
            ET.register_namespace(prefix, uri)
        xml_bytes = ET.tostring(root, encoding="utf-8", xml_declaration=True)
    if b"<ns0:" in xml_bytes or b"</ns0:" in xml_bytes or b"xmlns:ns0=" in xml_bytes:
        raise RuntimeError(f"Unexpected anonymous namespace prefix in {kind} output")
    return xml_bytes


def serialize_wordprocessingml(root: ET.Element) -> bytes:
    return _serialize(root, _WORDPROCESSINGML_PREFIXES, "WordprocessingML")


def serialize_package_relationships(root: ET.Element) -> bytes:
    return _serialize(root, _PACKAGE_RELATIONSHIPS_PREFIXES, "package relationships")


def serialize_content_types(root: ET.Element) -> bytes:
    return _serialize(root, _CONTENT_TYPES_PREFIXES, "content types")
