"""
Phase 1 contract validation — validates both registries before writing to disk.

Public API:
    validate_template_registry(registry)   — arch_template_registry.json shape + XML
    validate_style_registry(registry)      — arch_style_registry.json shape + roles
    validate_cross_registry(style_registry, template_registry) — cross-check IDs
    validate_phase1_contracts(style_registry, template_registry) — all of the above
"""

from __future__ import annotations

import xml.etree.ElementTree as ET
from typing import Any, Dict, List, Set


# ---------------------------------------------------------------------------
# OOXML namespace declarations for wrapping XML fragments
# ---------------------------------------------------------------------------

_NS_DECLS = (
    'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
    'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" '
    'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" '
    'xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" '
    'xmlns:o="urn:schemas-microsoft-com:office:office" '
    'xmlns:v="urn:schemas-microsoft-com:vml" '
    'xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math" '
    'xmlns:wpc="http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas" '
    'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" '
    'xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" '
    'xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml" '
    'xmlns:w16se="http://schemas.microsoft.com/office/word/2015/wordml/symex"'
)

_REQUIRED_TOP_LEVEL_KEYS = {
    "meta",
    "package_inventory",
    "doc_defaults",
    "styles",
    "theme",
    "settings",
    "page_layout",
    "headers_footers",
    "numbering",
    "fonts",
}

# All supported roles for the architect style contract.
_ALLOWED_ROLES = {
    "SectionID",
    "SectionTitle",
    "PART",
    "ARTICLE",
    "PARAGRAPH",
    "SUBPARAGRAPH",
    "SUBSUBPARAGRAPH",
}


# ---------------------------------------------------------------------------
# XML fragment parser
# ---------------------------------------------------------------------------

def _parse_xml_fragment(xml_str: str, context: str) -> None:
    """Parse an XML fragment by wrapping it in a namespace-aware synthetic root.

    Raises ValueError if the fragment is not well-formed.
    """
    stripped = xml_str.strip()
    if not stripped:
        return  # empty string is a no-op

    # Strip XML declaration if present
    if stripped.startswith("<?xml"):
        idx = stripped.index("?>")
        stripped = stripped[idx + 2:].strip()

    try:
        ET.fromstring(f"<root {_NS_DECLS}>{stripped}</root>")
    except ET.ParseError as exc:
        raise ValueError(
            f"Malformed XML fragment at {context}: {exc}"
        ) from exc


def _validate_xml_field(obj: Any, field: str, context: str) -> None:
    """Validate a single XML field on a dict if it is present and non-None."""
    val = obj.get(field)
    if val is None:
        return
    if not isinstance(val, str):
        raise ValueError(f"{context}.{field} must be a string or null, got {type(val).__name__}")
    _parse_xml_fragment(val, f"{context}.{field}")


# ---------------------------------------------------------------------------
# Template registry validation
# ---------------------------------------------------------------------------

def validate_template_registry(registry: Dict[str, Any]) -> None:
    """Validate arch_template_registry.json structure and XML fragments.

    Raises ValueError on the first problem found.
    """
    if not isinstance(registry, dict):
        raise ValueError("Template registry must be a JSON object")

    # -- top-level shape --
    missing = _REQUIRED_TOP_LEVEL_KEYS - set(registry.keys())
    if missing:
        raise ValueError(f"Template registry missing required keys: {sorted(missing)}")

    # -- styles.style_defs is list --
    styles = registry["styles"]
    if not isinstance(styles, dict):
        raise ValueError("Template registry 'styles' must be an object")

    style_defs = styles.get("style_defs")
    if not isinstance(style_defs, list):
        raise ValueError(
            f"styles.style_defs must be a list, got {type(style_defs).__name__}"
        )

    # -- style IDs unique and non-empty --
    seen_ids: Set[str] = set()
    for i, sdef in enumerate(style_defs):
        sid = sdef.get("style_id")
        if not isinstance(sid, str) or not sid:
            raise ValueError(
                f"styles.style_defs[{i}] has empty or missing style_id"
            )
        if sid in seen_ids:
            raise ValueError(f"Duplicate style_id in template registry: '{sid}'")
        seen_ids.add(sid)

    # -- XML fragment validation --

    # doc_defaults
    dd = registry["doc_defaults"]
    drp = dd.get("default_run_props")
    if isinstance(drp, dict):
        _validate_xml_field(drp, "rPr", "doc_defaults.default_run_props")
    dpp = dd.get("default_paragraph_props")
    if isinstance(dpp, dict):
        _validate_xml_field(dpp, "pPr", "doc_defaults.default_paragraph_props")

    # styles.latent_styles
    ls = styles.get("latent_styles")
    if isinstance(ls, dict):
        _validate_xml_field(ls, "latentStyles_xml", "styles.latent_styles")

    # each style_defs[*] XML properties
    for i, sdef in enumerate(style_defs):
        ctx = f"styles.style_defs[{i}] (style_id={sdef.get('style_id', '?')})"
        for prop in ("raw_style_xml", "pPr", "rPr", "tblPr", "trPr", "tcPr"):
            _validate_xml_field(sdef, prop, ctx)

    # settings.compat
    settings = registry["settings"]
    if isinstance(settings, dict):
        compat = settings.get("compat")
        if isinstance(compat, dict):
            _validate_xml_field(compat, "compat_xml", "settings.compat")

    # numbering blocks
    numbering = registry["numbering"]
    if isinstance(numbering, dict):
        for key in ("abstract_nums", "nums"):
            items = numbering.get(key)
            if isinstance(items, list):
                for i, entry in enumerate(items):
                    if isinstance(entry, dict):
                        _validate_xml_field(entry, "xml", f"numbering.{key}[{i}]")

    # page_layout sectPr blocks
    pl = registry["page_layout"]
    if isinstance(pl, dict):
        ds = pl.get("default_section")
        if isinstance(ds, dict):
            _validate_xml_field(ds, "sectPr", "page_layout.default_section")
        chain = pl.get("section_chain")
        if isinstance(chain, list):
            for i, sec in enumerate(chain):
                if isinstance(sec, dict):
                    _validate_xml_field(sec, "sectPr", f"page_layout.section_chain[{i}]")

    # theme
    theme = registry["theme"]
    if isinstance(theme, dict):
        _validate_xml_field(theme, "theme1_xml", "theme")

    # headers/footers + relationship/media payloads
    hf = registry["headers_footers"]
    if isinstance(hf, dict):
        for section_name in ("headers", "footers"):
            items = hf.get(section_name)
            if not isinstance(items, list):
                continue
            for i, entry in enumerate(items):
                if not isinstance(entry, dict):
                    raise ValueError(f"headers_footers.{section_name}[{i}] must be an object")

                _validate_xml_field(entry, "xml", f"headers_footers.{section_name}[{i}]")

                rels_xml = entry.get("rels_xml")
                if rels_xml is not None and not isinstance(rels_xml, str):
                    raise ValueError(
                        f"headers_footers.{section_name}[{i}].rels_xml must be a string or null"
                    )

                media = entry.get("media")
                if media is None:
                    continue
                if not isinstance(media, list):
                    raise ValueError(f"headers_footers.{section_name}[{i}].media must be a list")
                for j, media_entry in enumerate(media):
                    if not isinstance(media_entry, dict):
                        raise ValueError(
                            f"headers_footers.{section_name}[{i}].media[{j}] must be an object"
                        )
                    for field in ("rel_id", "target", "content_type", "data_base64"):
                        val = media_entry.get(field)
                        if not isinstance(val, str):
                            raise ValueError(
                                f"headers_footers.{section_name}[{i}].media[{j}].{field} must be a string"
                            )

        header_footer_media = hf.get("header_footer_media")
        if header_footer_media is not None:
            if not isinstance(header_footer_media, list):
                raise ValueError("headers_footers.header_footer_media must be a list when present")
            for i, item in enumerate(header_footer_media):
                if not isinstance(item, str):
                    raise ValueError(
                        f"headers_footers.header_footer_media[{i}] must be a string"
                    )

    # fonts
    fonts = registry["fonts"]
    if isinstance(fonts, dict):
        _validate_xml_field(fonts, "font_table_xml", "fonts")


# ---------------------------------------------------------------------------
# Style registry validation
# ---------------------------------------------------------------------------

def validate_style_registry(registry: Dict[str, Any]) -> None:
    """Validate arch_style_registry.json structure.

    Raises ValueError on the first problem found.
    Does NOT require style_name (style_id is the hard requirement).
    """
    if not isinstance(registry, dict):
        raise ValueError("Style registry must be a JSON object")

    if registry.get("version") not in (1, 2):
        raise ValueError("style registry version must be 1 or 2")

    src = registry.get("source_docx")
    if not isinstance(src, str) or not src:
        raise ValueError("style registry source_docx must be a non-empty string")

    roles = registry.get("roles")
    if not isinstance(roles, dict):
        raise ValueError("style registry roles must be an object")

    for role, spec in roles.items():
        if role not in _ALLOWED_ROLES:
            raise ValueError(f"style registry contains unknown role '{role}'")
        if not isinstance(spec, dict):
            raise ValueError(f"roles['{role}'] must be an object")
        sid = spec.get("style_id")
        if not isinstance(sid, str) or not sid:
            raise ValueError(f"roles['{role}'].style_id must be a non-empty string")
        epi = spec.get("exemplar_paragraph_index")
        if not isinstance(epi, int) or epi < 0:
            raise ValueError(
                f"roles['{role}'].exemplar_paragraph_index must be a non-negative int"
            )
        resolved = spec.get("resolved_formatting")
        if resolved is not None and not isinstance(resolved, dict):
            raise ValueError(f"roles['{role}'].resolved_formatting must be an object when present")

        provenance = spec.get("numbering_provenance")
        if provenance is not None:
            allowed = ("style_numpr", "direct_numpr", "text_literal", "none")
            if provenance not in allowed:
                raise ValueError(
                    f"roles['{role}'].numbering_provenance must be one of: "
                    "style_numpr, direct_numpr, text_literal, none"
                )

        numbering_pattern = spec.get("numbering_pattern")
        if numbering_pattern is not None:
            if not isinstance(numbering_pattern, dict):
                raise ValueError(f"roles['{role}'].numbering_pattern must be an object when present")


# ---------------------------------------------------------------------------
# Cross-registry consistency
# ---------------------------------------------------------------------------

def validate_cross_registry(
    style_registry: Dict[str, Any],
    template_registry: Dict[str, Any],
) -> None:
    """Every role style_id in the style registry must exist in the template registry.

    Raises ValueError if any role references a style_id not present in
    template_registry.styles.style_defs.
    """
    template_ids: Set[str] = set()
    style_defs = template_registry.get("styles", {}).get("style_defs", [])
    for sdef in style_defs:
        sid = sdef.get("style_id")
        if sid:
            template_ids.add(sid)

    roles = style_registry.get("roles", {})
    missing: List[str] = []
    for role, spec in roles.items():
        sid = spec.get("style_id", "")
        if sid and sid not in template_ids:
            missing.append(f"  role '{role}' references style_id '{sid}'")

    if missing:
        raise ValueError(
            "Style registry references style_id(s) not found in template registry "
            "style_defs:\n" + "\n".join(missing)
        )


# ---------------------------------------------------------------------------
# Top-level entry point
# ---------------------------------------------------------------------------

def validate_phase1_contracts(
    style_registry: Dict[str, Any],
    template_registry: Dict[str, Any],
) -> None:
    """Validate both registries and their cross-consistency.

    Raises ValueError on the first problem found.
    Call this after building both registry dicts but before writing to disk.
    """
    validate_template_registry(template_registry)
    validate_style_registry(style_registry)
    validate_cross_registry(style_registry, template_registry)
