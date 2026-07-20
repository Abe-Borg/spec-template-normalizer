"""Tests for docx_patch — XML well-formedness validation."""

import zipfile
from pathlib import Path

import pytest

from spec_formatter.style_application.docx_patch import (
    is_allowed_header_footer_part_name,
    is_allowed_header_footer_rels_name,
    patch_docx,
    validate_xml_wellformedness,
)


# ── validate_xml_wellformedness ─────────────────────────────────────────────


_W_NS = b' xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'


class TestValidateXmlWellformedness:
    def test_valid_xml_passes(self):
        parts = {
            "word/styles.xml": b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles' + _W_NS + b"></w:styles>",
            "word/document.xml": b"<w:document" + _W_NS + b"><w:body/></w:document>",
        }
        assert validate_xml_wellformedness(parts) == []

    def test_unclosed_tag_detected(self):
        parts = {
            "word/styles.xml": b"<styles><style>",
        }
        errors = validate_xml_wellformedness(parts)
        assert len(errors) == 1
        assert "word/styles.xml" in errors[0]
        assert "XML parse error" in errors[0]

    def test_empty_bytes_detected(self):
        parts = {"word/document.xml": b""}
        errors = validate_xml_wellformedness(parts)
        assert len(errors) == 1
        assert "word/document.xml" in errors[0]

    def test_xml_declaration_with_valid_content_passes(self):
        content = b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/></Types>'
        assert validate_xml_wellformedness({"[Content_Types].xml": content}) == []

    def test_mixed_valid_and_invalid(self):
        parts = {
            "word/styles.xml": b"<root/>",
            "word/settings.xml": b"<broken><",
            "word/document.xml": b"<root/>",
        }
        errors = validate_xml_wellformedness(parts)
        assert len(errors) == 1
        assert "word/settings.xml" in errors[0]


# ── patch_docx integration ──────────────────────────────────────────────────


class TestPatchDocxXmlValidation:
    def _make_minimal_docx(self, path: Path) -> None:
        """Create a minimal valid DOCX (ZIP with one XML entry)."""
        with zipfile.ZipFile(path, "w") as zf:
            zf.writestr("word/document.xml", b"<w:document" + _W_NS + b"/>")

    def test_malformed_xml_prevents_repack(self, tmp_path):
        src = tmp_path / "input.docx"
        out = tmp_path / "output.docx"
        self._make_minimal_docx(src)

        with pytest.raises(RuntimeError, match="XML well-formedness check failed"):
            patch_docx(
                src_docx=src,
                out_docx=out,
                replacements={"word/document.xml": b"<broken><"},
            )
        assert not out.exists()

    def test_valid_xml_allows_repack(self, tmp_path):
        src = tmp_path / "input.docx"
        out = tmp_path / "output.docx"
        self._make_minimal_docx(src)

        patch_docx(
            src_docx=src,
            out_docx=out,
            replacements={"word/document.xml": b"<w:document" + _W_NS + b"><w:body/></w:document>"},
        )
        assert out.exists()

    def test_string_replacement_gets_truthful_utf8_declaration(self, tmp_path):
        src = tmp_path / "input.docx"
        out = tmp_path / "output.docx"
        self._make_minimal_docx(src)

        patch_docx(
            src_docx=src,
            out_docx=out,
            replacements={
                "word/document.xml": (
                    '<?xml version="1.0" encoding="UTF-16"?>'
                    '<w:document xmlns:w="http://schemas.openxmlformats.org/'
                    'wordprocessingml/2006/main"/>'
                )
            },
        )

        with zipfile.ZipFile(out) as archive:
            payload = archive.read("word/document.xml")
        assert b'encoding="UTF-8"' in payload
        payload.decode("utf-8")

    def test_duplicate_source_members_raise_explicitly_before_output_replacement(
        self, tmp_path
    ):
        src = tmp_path / "input.docx"
        out = tmp_path / "output.docx"
        with zipfile.ZipFile(src, "w") as archive:
            archive.writestr("word/document.xml", b"<first/>")
            archive.writestr("WORD/DOCUMENT.XML", b"<second/>")
        out.write_bytes(b"existing-output")

        with pytest.raises(RuntimeError, match="duplicate ZIP member"):
            patch_docx(
                src_docx=src,
                out_docx=out,
                replacements={"word/document.xml": b"<replacement/>"},
            )

        assert out.read_bytes() == b"existing-output"

    def test_replacement_rejects_case_insensitive_source_collision(self, tmp_path):
        src = tmp_path / "input.docx"
        out = tmp_path / "output.docx"
        with zipfile.ZipFile(src, "w") as archive:
            archive.writestr("WORD/DOCUMENT.XML", b"<source/>")
        out.write_bytes(b"existing-output")

        with pytest.raises(RuntimeError, match="collides with a source ZIP member"):
            patch_docx(
                src_docx=src,
                out_docx=out,
                replacements={"word/document.xml": b"<replacement/>"},
            )

        assert out.read_bytes() == b"existing-output"

    def test_replacements_reject_case_insensitive_duplicates(self, tmp_path):
        src = tmp_path / "input.docx"
        out = tmp_path / "output.docx"
        self._make_minimal_docx(src)

        with pytest.raises(RuntimeError, match="Replacement parts contain duplicate"):
            patch_docx(
                src_docx=src,
                out_docx=out,
                replacements={
                    "word/header1.xml": b"<header/>",
                    "WORD/HEADER1.XML": b"<other-header/>",
                },
            )

        assert not out.exists()


class TestPatchDocxHeaderFooterSupport:
    def _make_minimal_docx(self, path: Path) -> None:
        with zipfile.ZipFile(path, "w") as zf:
            zf.writestr("word/document.xml", b"<w:document" + _W_NS + b"/>")

    def _make_docx_with_header(self, path: Path) -> None:
        with zipfile.ZipFile(path, "w") as zf:
            zf.writestr("word/document.xml", b"<w:document" + _W_NS + b"/>")
            zf.writestr("word/header1.xml", b"<w:hdr" + _W_NS + b"/>")
            zf.writestr("word/_rels/header1.xml.rels", b"<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\"/>")

    def test_allows_media_binary_replacement(self, tmp_path):
        src = tmp_path / "input.docx"
        out = tmp_path / "output.docx"
        self._make_minimal_docx(src)

        patch_docx(
            src_docx=src,
            out_docx=out,
            replacements={
                "word/document.xml": b"<w:document" + _W_NS + b"><w:body/></w:document>",
                "word/media/image1.png": b"\x89PNG",
            },
        )
        with zipfile.ZipFile(out, "r") as zf:
            assert zf.read("word/media/image1.png") == b"\x89PNG"

    def test_allows_validated_nested_header_part_via_dynamic_allow_list(self, tmp_path):
        src = tmp_path / "input.docx"
        out = tmp_path / "output.docx"
        self._make_minimal_docx(src)
        part_name = "word/headers/default.xml"

        patch_docx(
            src_docx=src,
            out_docx=out,
            replacements={part_name: b"<w:hdr" + _W_NS + b"/>"},
            allowed_dynamic_parts={part_name},
        )

        with zipfile.ZipFile(out) as archive:
            assert archive.read(part_name).startswith(b"<w:hdr")

    @pytest.mark.parametrize(
        "part_name",
        [
            "word/headerFirst.xml",
            "word/footerEven.xml",
            "word/_rels/headerFirst.xml.rels",
            "word/_rels/footerEven.xml.rels",
        ],
    )
    def test_allows_safe_nonnumeric_header_footer_part_names(self, tmp_path, part_name):
        src = tmp_path / "input.docx"
        out = tmp_path / "output.docx"
        self._make_minimal_docx(src)
        replacement = (
            b'<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"/>'
            if part_name.endswith(".rels")
            else b"<w:hdr" + _W_NS + b"/>"
        )

        patch_docx(
            src_docx=src,
            out_docx=out,
            replacements={
                "word/document.xml": b"<w:document" + _W_NS + b"/>",
                part_name: replacement,
            },
        )

        with zipfile.ZipFile(out, "r") as zf:
            assert zf.read(part_name) == replacement

    @pytest.mark.parametrize(
        "part_name",
        [
            "word/../evil.xml",
            "word/header../evil.xml",
            r"word/header..\evil.xml",
            "word/header:evil.xml",
            "word/headerFirst.xml.bak",
        ],
    )
    def test_rejects_unsafe_header_footer_part_name_variants(self, part_name):
        assert not is_allowed_header_footer_part_name(part_name)

    def test_allows_safe_nested_custom_header_footer_part_name(self):
        assert is_allowed_header_footer_part_name(
            "word/headers/default.xml",
            expected_kind="header",
        )
        assert is_allowed_header_footer_rels_name(
            "word/headers/_rels/default.xml.rels",
            owner_part="word/headers/default.xml",
        )

    @pytest.mark.parametrize(
        "rels_name",
        [
            "word/_rels/header../evil.xml.rels",
            r"word/_rels/header..\evil.xml.rels",
            "word/_rels/header:evil.xml.rels",
            "word/_rels/headerFirst.xml.rels.bak",
            "word/_rels/../headerFirst.xml.rels",
        ],
    )
    def test_rejects_unsafe_header_footer_rels_name_variants(self, rels_name):
        assert not is_allowed_header_footer_rels_name(rels_name)

    def test_exclude_parts_drops_old_header_parts(self, tmp_path):
        src = tmp_path / "input.docx"
        out = tmp_path / "output.docx"
        self._make_docx_with_header(src)

        patch_docx(
            src_docx=src,
            out_docx=out,
            replacements={"word/document.xml": b"<w:document" + _W_NS + b"><w:body/></w:document>"},
            exclude_parts={"word/header1.xml", "word/_rels/header1.xml.rels"},
        )
        with zipfile.ZipFile(out, "r") as zf:
            assert "word/header1.xml" not in zf.namelist()
            assert "word/_rels/header1.xml.rels" not in zf.namelist()
