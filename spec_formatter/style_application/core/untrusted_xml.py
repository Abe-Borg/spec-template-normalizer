"""Parse XML from untrusted DOCX parts without entity expansion risk.

``xml.etree.ElementTree`` expands internal entities declared in a DOCTYPE, so
a hostile ``word/document.xml`` or ``word/numbering.xml`` could carry an
entity-expansion payload. Word never writes a DOCTYPE into an OOXML part, so
every parser in the target engine goes through :func:`parse_untrusted_xml`,
which rejects any DOCTYPE or ENTITY declaration before the bytes reach expat
and wraps parse errors with the part name for the caller.

Rejection is encoding-independent. Three steps run over one immutable byte
payload, so the representation that is screened is always the representation
that is parsed:

0. A ``str`` is already decoded, so its XML declaration is made truthful
   before encoding. Skipping this parses UTF-8 bytes back through a stale
   declared encoding -- ``windows-1252`` silently yields mojibake and
   ``utf-16`` fails outright.
1. A byte scan rejects ``<!DOCTYPE``/``<!ENTITY`` anywhere in the payload.
   This is deliberately broader than XML requires: it also rejects
   declaration-shaped text inside comments and CDATA, and text that is not
   well-formed at all. That conservatism is the long-standing contract.
2. Expat rejects a real declaration in any encoding it can auto-detect,
   which the byte scan alone cannot do -- UTF-16 encodes ``<!DOCTYPE`` as
   ``<\\x00!\\x00D\\x00...`` and does not match an ASCII pattern.

Well-formedness is not this module's judgement to make twice. When the
Expat pass finds a payload malformed it stays quiet and lets ElementTree
parse the same bytes and word the error, so a caller never sees a message
that depends on which parser noticed first.
"""

from __future__ import annotations

import re
import xml.etree.ElementTree as ET
import xml.parsers.expat as expat
from typing import Union

from .ooxml_text import prepare_xml_text_for_utf8

_DOCTYPE_RE = re.compile(rb"<!(?:DOCTYPE|ENTITY)", re.IGNORECASE)

_DOCTYPE_REJECTED = (
    "{context}: DOCTYPE/ENTITY declarations are not allowed in an "
    "untrusted OOXML part"
)


class UntrustedXmlError(ValueError):
    """An untrusted XML part was rejected before or during parsing."""


class _DoctypeDeclared(Exception):
    """Raised out of expat's handler the moment a DOCTYPE begins."""


def _reject_doctype_in_any_encoding(payload: bytes, context: str) -> None:
    """Reject a DOCTYPE that the byte scan cannot see because of encoding.

    ``StartDoctypeDeclHandler`` fires as expat begins the document-type
    declaration, before any entity is declared or expanded, and expat has
    already auto-detected the encoding by then. A payload malformed *before*
    its DOCTYPE never reaches the handler, but ElementTree fails on the same
    bytes at the same place, so nothing is expanded either way.
    """

    parser = expat.ParserCreate()

    def _start_doctype(name, system_id, public_id, has_internal_subset):
        raise _DoctypeDeclared

    parser.StartDoctypeDeclHandler = _start_doctype
    try:
        parser.Parse(payload, True)
    except _DoctypeDeclared:
        raise UntrustedXmlError(_DOCTYPE_REJECTED.format(context=context)) from None
    except expat.ExpatError:
        return
    except (LookupError, ValueError) as exc:
        # An encoding expat cannot use: a codec Python does not have
        # (LookupError) or a multi-byte one it refuses internally
        # (ValueError). Neither is an ExpatError and neither carries the
        # part name, so both used to escape callers handling
        # UntrustedXmlError.
        raise UntrustedXmlError(
            f"{context}: unsupported XML encoding: {exc}"
        ) from exc


def parse_untrusted_xml(data: Union[bytes, str], context: str) -> ET.Element:
    """Parse ``data`` (bytes or already-decoded text) and return its root.

    ``context`` names the part for error messages (``word/document.xml``).
    Raises :class:`UntrustedXmlError` (a ``ValueError``) when the payload
    declares a DOCTYPE or ENTITY in any encoding, or when it is not
    well-formed.
    """

    if isinstance(data, str):
        payload = prepare_xml_text_for_utf8(data).encode("utf-8")
    else:
        payload = data

    if _DOCTYPE_RE.search(payload):
        raise UntrustedXmlError(_DOCTYPE_REJECTED.format(context=context))

    _reject_doctype_in_any_encoding(payload, context)

    try:
        return ET.fromstring(payload)
    except ET.ParseError as exc:
        raise UntrustedXmlError(f"{context}: XML parse error: {exc}") from exc
