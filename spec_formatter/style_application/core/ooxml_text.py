"""Encoding-aware reads and UTF-8 writes for OOXML XML parts.

OPC permits XML encodings other than UTF-8. Phase 2 decodes target parts from
their BOM or XML declaration, then serializes only modified/generated parts as
UTF-8 with a truthful declaration.
"""

from __future__ import annotations

import codecs
import re
from pathlib import Path


_DECLARED_ENCODING = re.compile(
    br"<\?xml[^>]*\bencoding\s*=\s*['\"]([A-Za-z0-9._:-]+)['\"]",
    re.IGNORECASE,
)
_TEXT_DECLARED_ENCODING = re.compile(
    r"(<\?xml[^>]*\bencoding\s*=\s*['\"])([^'\"]+)(['\"])",
    re.IGNORECASE,
)


def decode_xml_bytes(data: bytes, *, part_name: str = "XML part") -> str:
    """Decode XML bytes according to their BOM, byte order, or declaration."""
    if data.startswith((codecs.BOM_UTF32_LE, codecs.BOM_UTF32_BE)):
        encoding = "utf-32"
    elif data.startswith((codecs.BOM_UTF16_LE, codecs.BOM_UTF16_BE)):
        encoding = "utf-16"
    elif data.startswith(codecs.BOM_UTF8):
        encoding = "utf-8-sig"
    elif data.startswith(b"\x00\x00\x00<"):
        encoding = "utf-32-be"
    elif data.startswith(b"<\x00\x00\x00"):
        encoding = "utf-32-le"
    elif data.startswith(b"\x00<\x00?"):
        encoding = "utf-16-be"
    elif data.startswith(b"<\x00?\x00"):
        encoding = "utf-16-le"
    else:
        match = _DECLARED_ENCODING.search(data[:512])
        encoding = match.group(1).decode("ascii") if match else "utf-8"

    try:
        codecs.lookup(encoding)
        return data.decode(encoding)
    except (LookupError, UnicodeDecodeError) as exc:
        raise ValueError(
            f"Could not decode {part_name} using {encoding!r}: {exc}"
        ) from exc


def read_xml_text(path: Path) -> str:
    """Read an OOXML XML part without assuming UTF-8."""
    path = Path(path)
    return decode_xml_bytes(path.read_bytes(), part_name=str(path))


def prepare_xml_text_for_utf8(text: str) -> str:
    """Make an existing XML declaration truthful before UTF-8 serialization."""
    return _TEXT_DECLARED_ENCODING.sub(r"\1UTF-8\3", text, count=1)


def write_xml_text(path: Path, text: str) -> None:
    """Serialize an OOXML XML part as UTF-8 with a truthful declaration."""
    if not isinstance(text, str):
        raise TypeError("OOXML XML text must be a string")
    Path(path).write_bytes(prepare_xml_text_for_utf8(text).encode("utf-8"))
