"""Parse XML from untrusted DOCX parts without entity expansion risk.

``xml.etree.ElementTree`` expands internal entities declared in a DOCTYPE, so
a hostile ``word/document.xml`` or ``word/numbering.xml`` could carry an
entity-expansion payload. Word never writes a DOCTYPE into an OOXML part, so
every parser in the target engine goes through :func:`parse_untrusted_xml`,
which rejects any DOCTYPE or ENTITY declaration before the bytes reach expat
and wraps parse errors with the part name for the caller.
"""

from __future__ import annotations

import re
import xml.etree.ElementTree as ET
from typing import Union

_DOCTYPE_RE = re.compile(rb"<!(?:DOCTYPE|ENTITY)", re.IGNORECASE)


class UntrustedXmlError(ValueError):
    """An untrusted XML part was rejected before or during parsing."""


def parse_untrusted_xml(data: Union[bytes, str], context: str) -> ET.Element:
    """Parse ``data`` (bytes or already-prepared text) and return its root.

    ``context`` names the part for error messages (``word/document.xml``).
    Raises :class:`UntrustedXmlError` (a ``ValueError``) when the payload
    declares a DOCTYPE or ENTITY, or when it is not well-formed.
    """

    payload = data.encode("utf-8") if isinstance(data, str) else data
    if _DOCTYPE_RE.search(payload):
        raise UntrustedXmlError(
            f"{context}: DOCTYPE/ENTITY declarations are not allowed in an "
            "untrusted OOXML part"
        )
    try:
        return ET.fromstring(payload)
    except ET.ParseError as exc:
        raise UntrustedXmlError(f"{context}: XML parse error: {exc}") from exc
