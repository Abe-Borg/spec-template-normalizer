"""Tests for core.registry.preflight_validate_registries()."""

import pytest
import spec_formatter.style_application.core.registry as registry_module
from spec_formatter.style_application.core.registry import preflight_validate_registries


# ---------------------------------------------------------------------------
# Helpers to build minimal valid registries
# ---------------------------------------------------------------------------

def _minimal_style_registry():
    """role -> styleId mapping (output of load_arch_style_registry)."""
    return {"PART": "CSIPart", "ARTICLE": "CSIArticle"}


def _minimal_template_registry():
    """Minimal valid arch_template_registry dict."""
    return {
        "styles": {
            "style_defs": [
                {
                    "style_id": "CSIPart",
                    "type": "paragraph",
                    "name": "CSI Part",
                },
                {
                    "style_id": "CSIArticle",
                    "type": "paragraph",
                    "name": "CSI Article",
                },
            ]
        },
        "page_layout": {
            "default_section": {
                "sectPr": "<w:sectPr><w:pgSz w:w=\"12240\" w:h=\"15840\"/><w:pgMar w:top=\"1800\" w:right=\"1080\" w:bottom=\"1440\" w:left=\"2160\" w:header=\"900\" w:footer=\"720\"/></w:sectPr>"
            },
            "section_chain": [],
        },
    }


# ---------------------------------------------------------------------------
# 1. Valid registries → no errors
# ---------------------------------------------------------------------------

def test_valid_registries_no_errors():
    errors = preflight_validate_registries(
        _minimal_style_registry(), _minimal_template_registry()
    )
    assert errors == []


# ---------------------------------------------------------------------------
# 2. Wrong section types
# ---------------------------------------------------------------------------

def test_wrong_section_types():
    tmpl = _minimal_template_registry()
    tmpl["theme"] = "not-a-dict"
    errors = preflight_validate_registries(_minimal_style_registry(), tmpl)
    assert any("'theme' must be dict" in e for e in errors)


# ---------------------------------------------------------------------------
# 3. style_defs not a list
# ---------------------------------------------------------------------------

def test_style_defs_not_list():
    tmpl = {
        "styles": {"style_defs": {"bad": True}},
        "page_layout": {"default_section": {"sectPr": "<w:sectPr/>"}},
    }
    errors = preflight_validate_registries({}, tmpl)
    assert any("style_defs must be a list" in e for e in errors)


# ---------------------------------------------------------------------------
# 4. Duplicate style_id
# ---------------------------------------------------------------------------

def test_duplicate_style_ids():
    tmpl = {
        "styles": {
            "style_defs": [
                {"style_id": "Dup", "type": "paragraph", "name": "A"},
                {"style_id": "Dup", "type": "paragraph", "name": "B"},
            ]
        },
        "page_layout": {"default_section": {"sectPr": "<w:sectPr/>"}},
    }
    errors = preflight_validate_registries({}, tmpl)
    assert any("Duplicate style_id 'Dup'" in e for e in errors)


# ---------------------------------------------------------------------------
# 5. Malformed XML fragment in style_def
# ---------------------------------------------------------------------------

def test_malformed_xml_fragment():
    tmpl = {
        "styles": {
            "style_defs": [
                {
                    "style_id": "Bad",
                    "type": "paragraph",
                    "name": "Bad Style",
                    "pPr": "<w:pPr><w:spacing w:after='200'/>",  # missing close
                },
            ]
        },
        "page_layout": {"default_section": {"sectPr": "<w:sectPr/>"}},
    }
    errors = preflight_validate_registries({}, tmpl)
    assert any("w:pPr" in e and "malformed" in e for e in errors)


def test_self_closing_xml_fragment_is_valid():
    tmpl = {
        "styles": {
            "style_defs": [
                {
                    "style_id": "Good",
                    "type": "paragraph",
                    "name": "Good Style",
                    "pPr": '<w:pPr w:val="x"/>',
                },
            ]
        },
        "page_layout": {"default_section": {"sectPr": "<w:sectPr/>"}},
    }
    errors = preflight_validate_registries({}, tmpl)
    assert errors == []


# ---------------------------------------------------------------------------
# 6. Invalid compat_xml
# ---------------------------------------------------------------------------

def test_invalid_compat_xml():
    tmpl = _minimal_template_registry()
    tmpl["settings"] = {"compat": {"compat_xml": "<w:compat><w:useFELayout/>"}}
    errors = preflight_validate_registries(_minimal_style_registry(), tmpl)
    assert any("compat_xml" in e and "malformed" in e for e in errors)


def test_valid_compat_xml():
    tmpl = _minimal_template_registry()
    tmpl["settings"] = {
        "compat": {"compat_xml": "<w:compat><w:useFELayout/></w:compat>"}
    }
    errors = preflight_validate_registries(_minimal_style_registry(), tmpl)
    assert errors == []


# ---------------------------------------------------------------------------
# 7. Style ID missing from template
# ---------------------------------------------------------------------------

def test_style_id_missing_from_template():
    style_reg = {"PART": "CSIPart", "ARTICLE": "MissingStyle"}
    tmpl = {
        "styles": {
            "style_defs": [
                {"style_id": "CSIPart", "type": "paragraph", "name": "Part"},
            ]
        },
        "page_layout": {"default_section": {"sectPr": "<w:sectPr/>"}},
    }
    errors = preflight_validate_registries(style_reg, tmpl)
    assert any("MissingStyle" in e and "ARTICLE" in e for e in errors)


# ---------------------------------------------------------------------------
# 8. Numbering abstractNumId ref missing
# ---------------------------------------------------------------------------

def test_numbering_abstract_ref_missing():
    tmpl = _minimal_template_registry()
    tmpl["numbering"] = {
        "abstract_nums": [{"abstractNumId": 1, "xml": "<w:abstractNum/>"}],
        "nums": [{"numId": 5, "abstractNumId": 99, "xml": "<w:num/>"}],
    }
    errors = preflight_validate_registries(_minimal_style_registry(), tmpl)
    assert any("numId=5" in e and "abstractNumId=99" in e for e in errors)


def test_numbering_consistent():
    tmpl = _minimal_template_registry()
    tmpl["numbering"] = {
        "abstract_nums": [{"abstractNumId": 1, "xml": "<w:abstractNum/>"}],
        "nums": [{"numId": 5, "abstractNumId": 1, "xml": "<w:num/>"}],
    }
    errors = preflight_validate_registries(_minimal_style_registry(), tmpl)
    assert errors == []


# ---------------------------------------------------------------------------
# 9. Empty style_defs list is valid
# ---------------------------------------------------------------------------

def test_empty_style_defs_is_ok():
    tmpl = {
        "styles": {"style_defs": []},
        "page_layout": {"default_section": {"sectPr": "<w:sectPr/>"}},
    }
    errors = preflight_validate_registries({}, tmpl)
    assert errors == []


# ---------------------------------------------------------------------------
# 10. Missing optional sections is valid
# ---------------------------------------------------------------------------

def test_missing_page_layout_is_error():
    tmpl = {"styles": _minimal_template_registry()["styles"]}
    errors = preflight_validate_registries({}, tmpl)
    assert any("missing page_layout" in e for e in errors)


# ---------------------------------------------------------------------------
# Additional edge cases
# ---------------------------------------------------------------------------

def test_style_def_not_dict():
    tmpl = {
        "styles": {"style_defs": ["not-a-dict"]},
        "page_layout": {"default_section": {"sectPr": "<w:sectPr/>"}},
    }
    errors = preflight_validate_registries({}, tmpl)
    assert any("must be a dict" in e for e in errors)


def test_malformed_theme_xml():
    tmpl = _minimal_template_registry()
    tmpl["theme"] = {"theme1_xml": "<a:theme><a:themeElements/>"}
    errors = preflight_validate_registries(_minimal_style_registry(), tmpl)
    assert any("theme1_xml" in e and "malformed" in e for e in errors)


def test_malformed_font_table_xml():
    tmpl = _minimal_template_registry()
    tmpl["fonts"] = {"font_table_xml": "<w:fonts>incomplete"}
    errors = preflight_validate_registries(_minimal_style_registry(), tmpl)
    assert any("font_table_xml" in e and "malformed" in e for e in errors)


def test_numbering_nums_not_list():
    tmpl = _minimal_template_registry()
    tmpl["numbering"] = {"abstract_nums": [], "nums": "bad"}
    errors = preflight_validate_registries(_minimal_style_registry(), tmpl)
    assert any("numbering.nums must be a list" in e for e in errors)


def test_collects_all_errors_at_once():
    """Verify that multiple errors are collected in a single pass."""
    style_reg = {"PART": "MissingA", "ARTICLE": "MissingB"}
    tmpl = {
        "styles": "not-a-dict",  # wrong type
        "theme": 42,             # wrong type
    }
    errors = preflight_validate_registries(style_reg, tmpl)
    # At minimum: styles wrong type + theme wrong type + 2 missing cross-refs
    assert len(errors) >= 3


def _header_entry(*, part_xml, rels_xml, media=None, rel_id="rIdHeader"):
    return {
        "part_name": "word/header1.xml",
        "rel_id": rel_id,
        "xml": part_xml,
        "rels_xml": rels_xml,
        "media": list(media or []),
    }


def _with_header(template, entry):
    template["headers_footers"] = {
        "headers": [entry],
        "footers": [],
        "header_footer_media": [
            item["target"] for item in entry.get("media", [])
        ],
    }
    return template


def test_rejects_embedded_font_references_without_payload_contract():
    tmpl = _minimal_template_registry()
    tmpl["fonts"] = {
        "font_table_xml": (
            '<w:fonts xmlns:w="http://schemas.openxmlformats.org/'
            'wordprocessingml/2006/main" '
            'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/'
            'relationships"><w:font w:name="Embedded">'
            '<w:embedRegular r:id="rId1"/></w:font></w:fonts>'
        )
    }

    errors = preflight_validate_registries(_minimal_style_registry(), tmpl)

    assert any("embedded-font references" in error for error in errors)


def test_allows_external_header_relationship_without_dereferencing_it():
    tmpl = _with_header(
        _minimal_template_registry(),
        _header_entry(
            part_xml=(
                '<w:hdr xmlns:w="http://schemas.openxmlformats.org/'
                'wordprocessingml/2006/main" '
                'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/'
                'relationships"><w:hyperlink r:id="rId1"/></w:hdr>'
            ),
            rels_xml=(
                '<Relationships xmlns="http://schemas.openxmlformats.org/package/'
                '2006/relationships"><Relationship Id="rId1" '
                'Type="http://schemas.openxmlformats.org/officeDocument/2006/'
                'relationships/hyperlink" Target="https://example.com" '
                'TargetMode="External"/></Relationships>'
            ),
        ),
    )

    assert preflight_validate_registries(_minimal_style_registry(), tmpl) == []


def test_rejects_unsupported_internal_header_relationship():
    tmpl = _with_header(
        _minimal_template_registry(),
        _header_entry(
            part_xml=(
                '<w:hdr xmlns:w="http://schemas.openxmlformats.org/'
                'wordprocessingml/2006/main"/>'
            ),
            rels_xml=(
                '<Relationships xmlns="http://schemas.openxmlformats.org/package/'
                '2006/relationships"><Relationship Id="rId1" '
                'Type="http://schemas.openxmlformats.org/officeDocument/2006/'
                'relationships/chart" Target="charts/chart1.xml"/>'
                '</Relationships>'
            ),
        ),
    )

    errors = preflight_validate_registries(_minimal_style_registry(), tmpl)

    assert any("unsupported internal non-image relationship" in error for error in errors)


def test_rejects_internal_image_without_captured_payload():
    tmpl = _with_header(
        _minimal_template_registry(),
        _header_entry(
            part_xml=(
                '<w:hdr xmlns:w="http://schemas.openxmlformats.org/'
                'wordprocessingml/2006/main" '
                'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/'
                'relationships"><w:drawing r:embed="rIdImage"/></w:hdr>'
            ),
            rels_xml=(
                '<Relationships xmlns="http://schemas.openxmlformats.org/package/'
                '2006/relationships"><Relationship Id="rIdImage" '
                'Type="http://schemas.openxmlformats.org/officeDocument/2006/'
                'relationships/image" Target="media/image1.png"/>'
                '</Relationships>'
            ),
        ),
    )

    errors = preflight_validate_registries(_minimal_style_registry(), tmpl)

    assert any("has no captured media payload" in error for error in errors)


def test_accepts_nested_internal_image_target_with_matching_payload():
    media = {
        "rel_id": "rIdImage",
        "target": "media/linked/image1.png",
        "content_type": "image/png",
        "data_base64": "UE5H",
    }
    entry = _header_entry(
            part_xml=(
                '<w:hdr xmlns:w="http://schemas.openxmlformats.org/'
                'wordprocessingml/2006/main" '
                'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/'
                'relationships"><w:drawing r:embed="rIdImage"/></w:hdr>'
            ),
            rels_xml=(
                '<Relationships xmlns="http://schemas.openxmlformats.org/package/'
                '2006/relationships"><Relationship Id="rIdImage" '
                'Type="http://schemas.openxmlformats.org/officeDocument/2006/'
                'relationships/image" Target="../media/linked/image1.png"/>'
                '</Relationships>'
            ),
            media=[media],
    )
    entry["part_name"] = "word/headers/default.xml"
    entry["rels_part_name"] = "word/headers/_rels/default.xml.rels"
    tmpl = _with_header(_minimal_template_registry(), entry)

    assert preflight_validate_registries(_minimal_style_registry(), tmpl) == []


def test_rejects_case_insensitive_header_footer_owner_collision():
    first = _header_entry(
        part_xml=(
            '<w:hdr xmlns:w="http://schemas.openxmlformats.org/'
            'wordprocessingml/2006/main"/>'
        ),
        rels_xml=None,
        rel_id="rIdHeaderA",
    )
    second = dict(first)
    second["part_name"] = "word/HEADER1.xml"
    second["rel_id"] = "rIdHeaderB"
    tmpl = _minimal_template_registry()
    tmpl["headers_footers"] = {
        "headers": [first, second],
        "footers": [],
    }

    errors = preflight_validate_registries(_minimal_style_registry(), tmpl)

    assert any("collides case-insensitively" in error for error in errors)


def test_rejects_unknown_header_footer_reference_type():
    entry = _header_entry(
        part_xml=(
            '<w:hdr xmlns:w="http://schemas.openxmlformats.org/'
            'wordprocessingml/2006/main"/>'
        ),
        rels_xml=None,
    )
    tmpl = _with_header(_minimal_template_registry(), entry)
    tmpl["page_layout"]["default_section"]["header_refs"] = {
        "default-or-first": "rIdHeader"
    }

    errors = preflight_validate_registries(_minimal_style_registry(), tmpl)

    assert any("unsupported reference type" in error for error in errors)


def test_rejects_dangling_header_relationship_id_and_media_target_mismatch():
    media = {
        "rel_id": "rIdImage",
        "target": "media/other.png",
        "content_type": "image/png",
        "data_base64": "UE5H",
    }
    tmpl = _with_header(
        _minimal_template_registry(),
        _header_entry(
            part_xml=(
                '<w:hdr xmlns:w="http://schemas.openxmlformats.org/'
                'wordprocessingml/2006/main" '
                'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/'
                'relationships"><w:hyperlink r:id="rIdMissing"/></w:hdr>'
            ),
            rels_xml=(
                '<Relationships xmlns="http://schemas.openxmlformats.org/package/'
                '2006/relationships"><Relationship Id="rIdImage" '
                'Type="http://schemas.openxmlformats.org/officeDocument/2006/'
                'relationships/image" Target="media/image1.png"/>'
                '</Relationships>'
            ),
            media=[media],
        ),
    )

    errors = preflight_validate_registries(_minimal_style_registry(), tmpl)

    assert any("references missing relationship Id 'rIdMissing'" in error for error in errors)
    assert any("target does not match relationship" in error for error in errors)


def test_rejects_section_reference_to_missing_header_entry():
    tmpl = _minimal_template_registry()
    tmpl["headers_footers"] = {
        "headers": [],
        "footers": [],
        "header_footer_media": [],
    }
    tmpl["page_layout"]["default_section"]["header_refs"] = {
        "default": "rIdMissing"
    }

    errors = preflight_validate_registries(_minimal_style_registry(), tmpl)

    assert any("missing header relationship Id 'rIdMissing'" in error for error in errors)


@pytest.mark.parametrize(
    "unsafe_target",
    [
        "%2e%2e/media/image.png",
        "file%3A///outside/image.png",
        "media/image.png%3Fdownload=1",
        "media/image.png%23fragment",
    ],
)
def test_rejects_encoded_traversal_scheme_query_and_fragment_targets(unsafe_target):
    media = {
        "rel_id": "rIdImage",
        "target": "media/image.png",
        "content_type": "image/png",
        "data_base64": "UE5H",
    }
    tmpl = _with_header(
        _minimal_template_registry(),
        _header_entry(
            part_xml=(
                '<w:hdr xmlns:w="http://schemas.openxmlformats.org/'
                'wordprocessingml/2006/main" '
                'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/'
                'relationships"><w:drawing r:embed="rIdImage"/></w:hdr>'
            ),
            rels_xml=(
                '<Relationships xmlns="http://schemas.openxmlformats.org/package/'
                '2006/relationships"><Relationship Id="rIdImage" '
                'Type="http://schemas.openxmlformats.org/officeDocument/2006/'
                f'relationships/image" Target="{unsafe_target}"/>'
                '</Relationships>'
            ),
            media=[media],
        ),
    )

    errors = preflight_validate_registries(_minimal_style_registry(), tmpl)

    assert any("unsafe" in error.lower() for error in errors)


def test_rejects_empty_relationship_reference_in_header_xml():
    tmpl = _with_header(
        _minimal_template_registry(),
        _header_entry(
            part_xml=(
                '<w:hdr xmlns:w="http://schemas.openxmlformats.org/'
                'wordprocessingml/2006/main" '
                'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/'
                'relationships"><w:drawing r:embed=""/></w:hdr>'
            ),
            rels_xml=(
                '<Relationships xmlns="http://schemas.openxmlformats.org/package/'
                '2006/relationships"/>'
            ),
        ),
    )

    errors = preflight_validate_registries(_minimal_style_registry(), tmpl)

    assert any("empty r:embed relationship reference" in error for error in errors)


def test_rejects_percent_encoded_traversal_in_captured_media_target():
    media = {
        "rel_id": "rIdImage",
        "target": "%2e%2e/media/image.png",
        "content_type": "image/png",
        "data_base64": "UE5H",
    }
    tmpl = _with_header(
        _minimal_template_registry(),
        _header_entry(
            part_xml=(
                '<w:hdr xmlns:w="http://schemas.openxmlformats.org/'
                'wordprocessingml/2006/main" '
                'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/'
                'relationships"><w:drawing r:embed="rIdImage"/></w:hdr>'
            ),
            rels_xml=(
                '<Relationships xmlns="http://schemas.openxmlformats.org/package/'
                '2006/relationships"><Relationship Id="rIdImage" '
                'Type="http://schemas.openxmlformats.org/officeDocument/2006/'
                'relationships/image" Target="media/image.png"/>'
                '</Relationships>'
            ),
            media=[media],
        ),
    )

    errors = preflight_validate_registries(_minimal_style_registry(), tmpl)

    assert any("media[0].target is unsafe" in error for error in errors)


@pytest.mark.parametrize(
    "payload, content_type, expected",
    [
        ("not@@base64", "image/png", "invalid base64"),
        ("UE5H", "image/png\r\nInjected", "image media type"),
        ("UE5H", "text/plain", "image media type"),
    ],
)
def test_rejects_invalid_media_payload_or_mime(payload, content_type, expected):
    media = {
        "rel_id": "rIdImage",
        "target": "media/image.png",
        "content_type": content_type,
        "data_base64": payload,
    }
    tmpl = _with_header(
        _minimal_template_registry(),
        _header_entry(
            part_xml=(
                '<w:hdr xmlns:w="http://schemas.openxmlformats.org/'
                'wordprocessingml/2006/main" '
                'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/'
                'relationships"><w:drawing r:embed="rIdImage"/></w:hdr>'
            ),
            rels_xml=(
                '<Relationships xmlns="http://schemas.openxmlformats.org/package/'
                '2006/relationships"><Relationship Id="rIdImage" '
                'Type="http://schemas.openxmlformats.org/officeDocument/2006/'
                'relationships/image" Target="media/image.png"/>'
                '</Relationships>'
            ),
            media=[media],
        ),
    )

    errors = preflight_validate_registries(_minimal_style_registry(), tmpl)

    assert any(expected in error for error in errors)


def test_enforces_per_image_and_total_decoded_media_bounds(monkeypatch):
    monkeypatch.setattr(registry_module, "MAX_HEADER_FOOTER_MEDIA_BYTES", 2)
    monkeypatch.setattr(registry_module, "MAX_HEADER_FOOTER_MEDIA_TOTAL_BYTES", 3)

    oversized = {
        "rel_id": "rIdOne",
        "target": "media/one.png",
        "content_type": "image/x-custom",
        "data_base64": "MTIz",
    }
    tmpl = _with_header(
        _minimal_template_registry(),
        _header_entry(
            part_xml=(
                '<w:hdr xmlns:w="http://schemas.openxmlformats.org/'
                'wordprocessingml/2006/main" '
                'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/'
                'relationships"><w:drawing r:embed="rIdOne"/></w:hdr>'
            ),
            rels_xml=(
                '<Relationships xmlns="http://schemas.openxmlformats.org/package/'
                '2006/relationships"><Relationship Id="rIdOne" '
                'Type="http://schemas.openxmlformats.org/officeDocument/2006/'
                'relationships/image" Target="media/one.png"/>'
                '</Relationships>'
            ),
            media=[oversized],
        ),
    )

    errors = preflight_validate_registries(_minimal_style_registry(), tmpl)
    assert any("decoded media limit" in error or "limit is 2 bytes" in error for error in errors)

    media_items = [
        {
            "rel_id": "rIdOne",
            "target": "media/one.png",
            "content_type": "image/png",
            "data_base64": "MTI=",
        },
        {
            "rel_id": "rIdTwo",
            "target": "media/two.png",
            "content_type": "image/x-custom",
            "data_base64": "MzQ=",
        },
    ]
    tmpl = _with_header(
        _minimal_template_registry(),
        _header_entry(
            part_xml=(
                '<w:hdr xmlns:w="http://schemas.openxmlformats.org/'
                'wordprocessingml/2006/main" '
                'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/'
                'relationships"><w:drawing r:embed="rIdOne"/>'
                '<w:drawing r:embed="rIdTwo"/></w:hdr>'
            ),
            rels_xml=(
                '<Relationships xmlns="http://schemas.openxmlformats.org/package/'
                '2006/relationships"><Relationship Id="rIdOne" '
                'Type="http://schemas.openxmlformats.org/officeDocument/2006/'
                'relationships/image" Target="media/one.png"/>'
                '<Relationship Id="rIdTwo" '
                'Type="http://schemas.openxmlformats.org/officeDocument/2006/'
                'relationships/image" Target="media/two.png"/>'
                '</Relationships>'
            ),
            media=media_items,
        ),
    )

    errors = preflight_validate_registries(_minimal_style_registry(), tmpl)
    assert any("captured media exceeds the total 3-byte limit" in error for error in errors)
