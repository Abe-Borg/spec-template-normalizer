"""Tests for encoding-aware OOXML reads and truthful UTF-8 writes."""

from __future__ import annotations

import codecs
from pathlib import Path

import pytest

from spec_formatter.style_application.core.ooxml_text import (
    decode_xml_bytes,
    prepare_xml_text_for_utf8,
    read_xml_text,
    write_xml_text,
)


@pytest.mark.parametrize(
    "payload",
    [
        codecs.BOM_UTF16_LE
        + '<?xml version="1.0" encoding="UTF-16"?><root>check</root>'.encode(
            "utf-16-le"
        ),
        codecs.BOM_UTF16_BE
        + '<?xml version="1.0" encoding="UTF-16"?><root>check</root>'.encode(
            "utf-16-be"
        ),
        codecs.BOM_UTF8
        + b'<?xml version="1.0" encoding="UTF-8"?><root>check</root>',
        b'<?xml version="1.0" encoding="iso-8859-1"?><root>caf\xe9</root>',
    ],
)
def test_decode_xml_bytes_honors_bom_and_declared_encoding(payload: bytes) -> None:
    assert "<root>" in decode_xml_bytes(payload)


def test_read_xml_text_accepts_bomless_utf16_byte_order(tmp_path: Path) -> None:
    path = tmp_path / "document.xml"
    text = '<?xml version="1.0" encoding="UTF-16"?><root>value</root>'
    path.write_bytes(text.encode("utf-16-be"))

    assert read_xml_text(path) == text


def test_prepare_and_write_make_declaration_match_utf8_bytes(tmp_path: Path) -> None:
    path = tmp_path / "settings.xml"
    text = "<?xml version='1.0' encoding='UTF-16'?><root>caf\u00e9</root>"

    assert "encoding='UTF-8'" in prepare_xml_text_for_utf8(text)
    write_xml_text(path, text)

    payload = path.read_bytes()
    assert not payload.startswith((codecs.BOM_UTF8, codecs.BOM_UTF16_LE, codecs.BOM_UTF16_BE))
    assert b"encoding='UTF-8'" in payload
    assert payload.decode("utf-8").endswith("<root>caf\u00e9</root>")


def test_unknown_or_incorrect_declared_encoding_is_rejected() -> None:
    with pytest.raises(ValueError, match="Could not decode"):
        decode_xml_bytes(b'<?xml version="1.0" encoding="not-real"?><root/>')
