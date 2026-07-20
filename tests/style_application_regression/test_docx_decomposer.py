"""Fail-closed extraction tests for Phase 2 target DOCX packages."""

from __future__ import annotations

import stat
import zipfile
from pathlib import Path

import pytest

from spec_formatter.style_application import docx_decomposer
from spec_formatter.style_application.docx_decomposer import DocxDecomposer


REQUIRED_PARTS = {
    "[Content_Types].xml": b"<Types/>",
    "_rels/.rels": b"<Relationships/>",
    "word/document.xml": b"<w:document/>",
    "word/styles.xml": b"<w:styles/>",
}


def _write_package(path: Path, extra=()) -> None:
    with zipfile.ZipFile(path, "w", compression=zipfile.ZIP_STORED) as archive:
        for name, payload in REQUIRED_PARTS.items():
            archive.writestr(name, payload)
        for item in extra:
            if len(item) == 2:
                archive.writestr(item[0], item[1])
            else:
                archive.writestr(item[0], item[1], compress_type=item[2])

    # ``zipfile`` normalizes backslashes while writing on Windows. Replace
    # equal-length names in both headers so this fixture contains the hostile
    # bytes a cross-platform producer can put in a real archive.
    for item in extra:
        raw_name = item[0]
        if "\\" not in raw_name:
            continue
        normalized_name = raw_name.replace("\\", "/")
        payload = path.read_bytes()
        assert payload.count(normalized_name.encode("ascii")) == 2
        path.write_bytes(
            payload.replace(
                normalized_name.encode("ascii"),
                raw_name.encode("ascii"),
            )
        )


def test_extracts_valid_package_into_new_directory(tmp_path: Path) -> None:
    source = tmp_path / "input.docx"
    output = tmp_path / "extracted"
    _write_package(source, [("word/media/image.png", b"png")])

    result = DocxDecomposer(source).extract(output)

    assert result == output
    assert (output / "word" / "media" / "image.png").read_bytes() == b"png"


def test_refuses_existing_extraction_target_without_deleting_it(tmp_path: Path) -> None:
    source = tmp_path / "input.docx"
    output = tmp_path / "extracted"
    _write_package(source)
    output.mkdir()
    sentinel = output / "keep.txt"
    sentinel.write_text("keep", encoding="utf-8")

    with pytest.raises(FileExistsError, match="already exists"):
        DocxDecomposer(source).extract(output)

    assert sentinel.read_text(encoding="utf-8") == "keep"


@pytest.mark.parametrize(
    "unsafe_name",
    [
        "../outside.xml",
        "/absolute.xml",
        r"word\outside.xml",
        "word/../../outside.xml",
    ],
)
def test_rejects_unsafe_member_and_removes_partial_tree(
    tmp_path: Path, unsafe_name: str
) -> None:
    source = tmp_path / "unsafe.docx"
    output = tmp_path / "extracted"
    _write_package(source, [(unsafe_name, b"bad")])

    with pytest.raises(ValueError, match="Unsafe|escapes"):
        DocxDecomposer(source).extract(output)

    assert not output.exists()
    assert not (tmp_path / "outside.xml").exists()


def test_rejects_casefold_duplicate_members(tmp_path: Path) -> None:
    source = tmp_path / "duplicate.docx"
    output = tmp_path / "extracted"
    _write_package(source, [("WORD/DOCUMENT.XML", b"duplicate")])

    with pytest.raises(ValueError, match="duplicate member"):
        DocxDecomposer(source).extract(output)

    assert not output.exists()


def test_rejects_zip_symlink(tmp_path: Path) -> None:
    source = tmp_path / "symlink.docx"
    output = tmp_path / "extracted"
    symlink = zipfile.ZipInfo("word/media/link.png")
    symlink.create_system = 3
    symlink.external_attr = (stat.S_IFLNK | 0o777) << 16
    with zipfile.ZipFile(source, "w") as archive:
        for name, payload in REQUIRED_PARTS.items():
            archive.writestr(name, payload)
        archive.writestr(symlink, "../../outside")

    with pytest.raises(ValueError, match="symbolic link"):
        DocxDecomposer(source).extract(output)

    assert not output.exists()


def test_enforces_entry_part_and_total_limits(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    source = tmp_path / "limited.docx"
    output = tmp_path / "extracted"
    _write_package(source, [("word/extra.bin", b"12345")])

    monkeypatch.setattr(docx_decomposer, "MAX_PACKAGE_ENTRIES", 4)
    with pytest.raises(ValueError, match="entries"):
        DocxDecomposer(source).extract(output)
    assert not output.exists()

    monkeypatch.setattr(docx_decomposer, "MAX_PACKAGE_ENTRIES", 10)
    monkeypatch.setattr(docx_decomposer, "MAX_PACKAGE_PART_BYTES", 4)
    with pytest.raises(ValueError, match="per-part limit"):
        DocxDecomposer(source).extract(output)
    assert not output.exists()

    monkeypatch.setattr(docx_decomposer, "MAX_PACKAGE_PART_BYTES", 1024)
    monkeypatch.setattr(docx_decomposer, "MAX_PACKAGE_UNCOMPRESSED_BYTES", 20)
    with pytest.raises(ValueError, match="expands to"):
        DocxDecomposer(source).extract(output)
    assert not output.exists()


def test_rejects_suspicious_compression_ratio(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    source = tmp_path / "compression-bomb.docx"
    output = tmp_path / "extracted"
    _write_package(
        source,
        [("word/compression-bomb.bin", b"0" * 100_000, zipfile.ZIP_DEFLATED)],
    )
    monkeypatch.setattr(docx_decomposer, "MAX_COMPRESSION_RATIO", 10)

    with pytest.raises(ValueError, match="suspicious compression ratio"):
        DocxDecomposer(source).extract(output)

    assert not output.exists()


def test_rejects_missing_required_parts_and_removes_tree(tmp_path: Path) -> None:
    source = tmp_path / "incomplete.docx"
    output = tmp_path / "extracted"
    with zipfile.ZipFile(source, "w") as archive:
        archive.writestr("word/document.xml", b"<w:document/>")

    with pytest.raises(ValueError, match="missing required parts"):
        DocxDecomposer(source).extract(output)

    assert not output.exists()
