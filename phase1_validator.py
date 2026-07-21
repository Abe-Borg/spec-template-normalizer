"""
Phase 1 contract validation — validates both registries before writing to disk.

Public API:
    validate_template_registry(registry)   — arch_template_registry.json shape + XML
    validate_style_registry(registry)      — arch_style_registry.json shape + roles
    validate_cross_registry(style_registry, template_registry) — cross-check IDs
    validate_phase1_contracts(style_registry, template_registry) — all of the above
"""

from __future__ import annotations

import re
import xml.etree.ElementTree as ET
from typing import Any, Dict, Iterable, List, Optional, Set

from spec_formatter.role_contract import (
    ALLOWED_ARCH_STYLE_IDS as _ALLOWED_ARCH_STYLE_IDS,
    ALLOWED_ROLES as _ALLOWED_ROLES,
)


# ---------------------------------------------------------------------------
# OOXML namespace declarations for wrapping XML fragments
# ---------------------------------------------------------------------------

_KNOWN_NS = {
    "w":      "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
    "a":      "http://schemas.openxmlformats.org/drawingml/2006/main",
    "r":      "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "mc":     "http://schemas.openxmlformats.org/markup-compatibility/2006",
    "o":      "urn:schemas-microsoft-com:office:office",
    "v":      "urn:schemas-microsoft-com:vml",
    "m":      "http://schemas.openxmlformats.org/officeDocument/2006/math",
    "wpc":    "http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas",
    "wp":     "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing",
    "w14":    "http://schemas.microsoft.com/office/word/2010/wordml",
    "w15":    "http://schemas.microsoft.com/office/word/2012/wordml",
    "w16se":  "http://schemas.microsoft.com/office/word/2015/wordml/symex",
    "w16cid": "http://schemas.microsoft.com/office/word/2016/wordml/cid",
    "w16":    "http://schemas.microsoft.com/office/word/2018/wordml",
    "w16cex": "http://schemas.microsoft.com/office/word/2018/wordml/cex",
    "w16sdtdh": "http://schemas.microsoft.com/office/word/2020/wordml/sdtdatahash",
    "wp14":   "http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing",
    "wps":    "http://schemas.microsoft.com/office/word/2010/wordprocessingShape",
    "wne":    "http://schemas.microsoft.com/office/word/2006/wordml",
    "cx":     "http://schemas.microsoft.com/office/drawing/2014/chartex",
    "cx1":    "http://schemas.microsoft.com/office/drawing/2015/9/8/chartex",
    "sl":     "http://schemas.openxmlformats.org/schemaLibrary/2006/main",
}

# Pre-built declaration string from the known set
_NS_DECLS = " ".join(f'xmlns:{k}="{v}"' for k, v in _KNOWN_NS.items())

# Regex to find all namespace prefixes used in an XML fragment
# Matches element names like <prefix:local and attribute names like prefix:attr="
_PREFIX_RE = re.compile(r'(?:</?|[\s])([A-Za-z][A-Za-z0-9_]*):[A-Za-z]')

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
    "custom_xml",
    "capture_policy",
}
_TEMPLATE_SCHEMA_VERSION = "1.0.0"

# Public immutable views for callers that need to build prompts or perform
# additional semantic checks without maintaining a second role list.
ALLOWED_ROLES = _ALLOWED_ROLES
ALLOWED_ARCH_STYLE_IDS = _ALLOWED_ARCH_STYLE_IDS


# ---------------------------------------------------------------------------
# Phase 1 instruction validation
# ---------------------------------------------------------------------------

def _require_plain_int(value: Any, context: str) -> int:
    """Return *value* as an int, rejecting bool (which subclasses int)."""
    if type(value) is not int or value < 0:
        raise ValueError(f"{context} must be a non-negative int")
    return value


def validate_instruction_contract(
    instructions: Dict[str, Any],
    expected_paragraph_indices: Optional[Iterable[int]] = None,
) -> None:
    """Validate the formal Phase 1 instruction wire contract.

    This deliberately covers shape and coverage partitioning only.  Document-
    aware semantic checks (for example, whether a selected exemplar really is
    an ARTICLE) remain the responsibility of the classifier/runtime.

    When ``expected_paragraph_indices`` is supplied, every expected paragraph
    must occur exactly once in either ``apply_pStyle`` or
    ``ignored_paragraphs``.  The two collections are always required to be
    disjoint, even when no expected index set is supplied.
    """
    if not isinstance(instructions, dict):
        raise ValueError("Phase 1 instructions must be a JSON object")

    allowed_keys = {
        "create_styles",
        "apply_pStyle",
        "ignored_paragraphs",
        "roles",
        "notes",
    }
    extra = set(instructions) - allowed_keys
    if extra:
        raise ValueError(f"Invalid instruction keys: {sorted(extra)}")

    roles = instructions.get("roles")
    if not isinstance(roles, dict):
        raise ValueError("Missing/invalid required key: roles")

    create_styles = instructions.get("create_styles", [])
    if not isinstance(create_styles, list):
        raise ValueError("create_styles must be an array when present")
    created: Dict[str, int] = {}
    for i, style in enumerate(create_styles):
        context = f"create_styles[{i}]"
        if not isinstance(style, dict):
            raise ValueError(f"{context} must be an object")
        allowed = {
            "styleId", "name", "type", "derive_from_paragraph_index",
            "basedOn", "role",
        }
        extra_fields = set(style) - allowed
        if extra_fields:
            raise ValueError(f"{context} has invalid fields: {sorted(extra_fields)}")
        sid = style.get("styleId")
        if sid not in _ALLOWED_ARCH_STYLE_IDS:
            raise ValueError(f"{context}.styleId is not an allowed reserved ARCH style")
        if sid in created:
            raise ValueError(f"Duplicate create_styles styleId: {sid}")
        source_index = _require_plain_int(
            style.get("derive_from_paragraph_index"),
            f"{context}.derive_from_paragraph_index",
        )
        if style.get("type", "paragraph") != "paragraph":
            raise ValueError(f"{context}.type must be 'paragraph'")
        if "name" in style and not isinstance(style["name"], str):
            raise ValueError(f"{context}.name must be a string")
        if "basedOn" in style and style["basedOn"] is not None and not isinstance(style["basedOn"], str):
            raise ValueError(f"{context}.basedOn must be a string or null")
        if "role" in style and not isinstance(style["role"], str):
            raise ValueError(f"{context}.role must be a string")
        created[sid] = source_index

    apply_items = instructions.get("apply_pStyle", [])
    if not isinstance(apply_items, list):
        raise ValueError("apply_pStyle must be an array when present")
    applied: Set[int] = set()
    for i, item in enumerate(apply_items):
        context = f"apply_pStyle[{i}]"
        if not isinstance(item, dict):
            raise ValueError(f"{context} must be an object")
        extra_fields = set(item) - {"paragraph_index", "styleId"}
        if extra_fields:
            raise ValueError(f"{context} has invalid fields: {sorted(extra_fields)}")
        index = _require_plain_int(item.get("paragraph_index"), f"{context}.paragraph_index")
        sid = item.get("styleId")
        if not isinstance(sid, str) or not sid:
            raise ValueError(f"{context}.styleId must be a non-empty string")
        if index in applied:
            raise ValueError(f"Duplicate paragraph_index in apply_pStyle: {index}")
        applied.add(index)

    ignored_items = instructions.get("ignored_paragraphs", [])
    if not isinstance(ignored_items, list):
        raise ValueError("ignored_paragraphs must be an array when present")
    ignored: Set[int] = set()
    for i, item in enumerate(ignored_items):
        context = f"ignored_paragraphs[{i}]"
        if not isinstance(item, dict):
            raise ValueError(f"{context} must be an object")
        extra_fields = set(item) - {"paragraph_index", "reason"}
        if extra_fields:
            raise ValueError(f"{context} has invalid fields: {sorted(extra_fields)}")
        index = _require_plain_int(item.get("paragraph_index"), f"{context}.paragraph_index")
        reason = item.get("reason")
        if not isinstance(reason, str) or not reason.strip():
            raise ValueError(f"{context}.reason must be a non-empty string")
        if index in ignored:
            raise ValueError(f"Duplicate paragraph_index in ignored_paragraphs: {index}")
        ignored.add(index)

    overlap = sorted(applied & ignored)
    if overlap:
        raise ValueError(
            "apply_pStyle and ignored_paragraphs must be disjoint; "
            f"overlap={overlap[:20]}{'...' if len(overlap) > 20 else ''}"
        )

    for role, spec in roles.items():
        if role not in _ALLOWED_ROLES:
            raise ValueError(f"Unknown role '{role}'")
        if not isinstance(spec, dict):
            raise ValueError(f"roles['{role}'] must be an object")
        extra_fields = set(spec) - {"styleId", "exemplar_paragraph_index", "style_name"}
        if extra_fields:
            raise ValueError(f"roles['{role}'] has invalid fields: {sorted(extra_fields)}")
        sid = spec.get("styleId")
        if not isinstance(sid, str) or not sid:
            raise ValueError(f"roles['{role}'].styleId must be a non-empty string")
        exemplar = _require_plain_int(
            spec.get("exemplar_paragraph_index"),
            f"roles['{role}'].exemplar_paragraph_index",
        )
        if "style_name" in spec and spec["style_name"] is not None and not isinstance(spec["style_name"], str):
            raise ValueError(f"roles['{role}'].style_name must be a string or null")
        if sid in created and exemplar != created[sid]:
            raise ValueError(
                f"roles['{role}'].exemplar_paragraph_index ({exemplar}) must equal "
                f"derive_from_paragraph_index ({created[sid]}) for '{sid}'"
            )

    if "notes" in instructions:
        notes = instructions["notes"]
        if not isinstance(notes, list) or any(not isinstance(note, str) for note in notes):
            raise ValueError("notes must be an array of strings")

    if expected_paragraph_indices is not None:
        expected: Set[int] = set()
        for raw_index in expected_paragraph_indices:
            index = _require_plain_int(raw_index, "expected paragraph index")
            if index in expected:
                raise ValueError(f"Duplicate expected paragraph index: {index}")
            expected.add(index)
        covered = applied | ignored
        missing = sorted(expected - covered)
        unexpected = sorted(covered - expected)
        if missing or unexpected:
            raise ValueError(
                "instruction coverage partition mismatch; "
                f"missing={missing[:20]}{'...' if len(missing) > 20 else ''}, "
                f"unexpected={unexpected[:20]}{'...' if len(unexpected) > 20 else ''}"
            )


# ---------------------------------------------------------------------------
# XML fragment parser
# ---------------------------------------------------------------------------

def _build_ns_decls(xml_str: str) -> str:
    """Build namespace declarations that cover every prefix used in *xml_str*.

    Starts with the known OOXML set.  For any prefix found in the fragment
    that is NOT in the known set, a synthetic placeholder URI is generated
    so that :func:`ET.fromstring` never fails with 'unbound prefix'.

    This is safe because we are only checking well-formedness, not resolving
    namespace semantics.
    """
    used_prefixes = set(_PREFIX_RE.findall(xml_str))
    # 'xml' and 'xmlns' are reserved prefixes that must never be (re)declared
    used_prefixes.discard("xml")
    used_prefixes.discard("xmlns")

    unknown = used_prefixes - set(_KNOWN_NS.keys())
    if not unknown:
        return _NS_DECLS

    extra = " ".join(
        f'xmlns:{p}="urn:unknown-ooxml-ns:{p}"' for p in sorted(unknown)
    )
    return f"{_NS_DECLS} {extra}"


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

    ns_decls = _build_ns_decls(stripped)

    try:
        ET.fromstring(f"<root {ns_decls}>{stripped}</root>")
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

    meta = registry["meta"]
    if not isinstance(meta, dict):
        raise ValueError("Template registry 'meta' must be an object")
    if meta.get("schema_version") != _TEMPLATE_SCHEMA_VERSION:
        raise ValueError(
            f"Unsupported template registry schema_version: {meta.get('schema_version')!r}"
        )
    if not isinstance(registry["package_inventory"], dict):
        raise ValueError("Template registry 'package_inventory' must be an object")
    if not isinstance(registry["headers_footers"], dict):
        raise ValueError("Template registry 'headers_footers' must be an object")
    if not isinstance(registry["custom_xml"], dict):
        raise ValueError("Template registry 'custom_xml' must be an object")
    if not isinstance(registry["capture_policy"], dict):
        raise ValueError("Template registry 'capture_policy' must be an object")

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

    version = registry.get("version")
    if version not in (1, 2):
        raise ValueError("style registry version must be 1 or 2")

    allowed_top_level = {"version", "source_docx", "roles"}
    if version == 2:
        allowed_top_level.update({"source_tokens", "source_sha256"})
    extra_top_level = set(registry) - allowed_top_level
    if extra_top_level:
        raise ValueError(f"style registry contains unknown keys: {sorted(extra_top_level)}")

    src = registry.get("source_docx")
    if not isinstance(src, str) or not src:
        raise ValueError("style registry source_docx must be a non-empty string")

    source_sha256 = registry.get("source_sha256")
    if version == 2 and source_sha256 is None:
        raise ValueError("style registry version 2 requires source_sha256")
    if source_sha256 is not None:
        if version != 2:
            raise ValueError("style registry source_sha256 is supported only in version 2")
        if not isinstance(source_sha256, str) or re.fullmatch(r"[0-9a-f]{64}", source_sha256) is None:
            raise ValueError("style registry source_sha256 must be a lowercase SHA-256 digest")

    roles = registry.get("roles")
    if not isinstance(roles, dict):
        raise ValueError("style registry roles must be an object")
    if version == 2 and not roles:
        raise ValueError("style registry version 2 roles must not be empty")

    source_tokens = registry.get("source_tokens")
    if version == 2 and source_tokens is None:
        raise ValueError("style registry version 2 requires source_tokens")
    if source_tokens is not None:
        if version != 2:
            raise ValueError("style registry source_tokens is supported only in version 2")
        if not isinstance(source_tokens, dict):
            raise ValueError("style registry source_tokens must be an object")
        unknown_tokens = set(source_tokens) - {"SectionID", "SectionTitle"}
        if unknown_tokens:
            raise ValueError(f"style registry source_tokens contains unknown keys: {sorted(unknown_tokens)}")
        for token_role, token in source_tokens.items():
            if not isinstance(token, str) or not token:
                raise ValueError(
                    f"style registry source_tokens['{token_role}'] must be a non-empty string"
                )

    for role, spec in roles.items():
        if role not in _ALLOWED_ROLES:
            raise ValueError(f"style registry contains unknown role '{role}'")
        if not isinstance(spec, dict):
            raise ValueError(f"roles['{role}'] must be an object")
        allowed_role_fields = {
            "style_id",
            "exemplar_paragraph_index",
            "style_name",
            "resolved_formatting",
            "warning",
            "numbering_provenance",
            "numbering_pattern",
        }
        extra_role_fields = set(spec) - allowed_role_fields
        if extra_role_fields:
            raise ValueError(
                f"roles['{role}'] contains unknown fields: {sorted(extra_role_fields)}"
            )
        sid = spec.get("style_id")
        if not isinstance(sid, str) or not sid:
            raise ValueError(f"roles['{role}'].style_id must be a non-empty string")
        epi = spec.get("exemplar_paragraph_index")
        if type(epi) is not int or epi < 0:
            raise ValueError(
                f"roles['{role}'].exemplar_paragraph_index must be a non-negative int"
            )
        style_name = spec.get("style_name")
        if style_name is not None and not isinstance(style_name, str):
            raise ValueError(f"roles['{role}'].style_name must be a string or null")
        resolved = spec.get("resolved_formatting")
        if resolved is not None and not isinstance(resolved, dict):
            raise ValueError(f"roles['{role}'].resolved_formatting must be an object when present")
        warning = spec.get("warning")
        if warning is not None and not isinstance(warning, str):
            raise ValueError(f"roles['{role}'].warning must be a string when present")

        provenance = spec.get("numbering_provenance")
        if version == 2 and provenance is None:
            raise ValueError(
                f"roles['{role}'].numbering_provenance is required in version 2"
            )
        if provenance is not None:
            allowed = (
                ("style_numpr", "text_literal", "none")
                if version == 1
                else ("style_numpr", "direct_numpr", "text_literal", "none")
            )
            if provenance not in allowed:
                raise ValueError(
                    f"roles['{role}'].numbering_provenance must be one of: "
                    "style_numpr, direct_numpr, text_literal, none"
                )

        numbering_pattern = spec.get("numbering_pattern")
        if numbering_pattern is not None:
            if not isinstance(numbering_pattern, dict):
                raise ValueError(f"roles['{role}'].numbering_pattern must be an object when present")
            allowed_pattern_keys = {"numId", "ilvl", "numFmt", "lvlText"}
            if version == 2:
                allowed_pattern_keys.update({
                    "abstractNumId",
                    "start",
                    "lvlRestart",
                    "suff",
                    "isLgl",
                    "startOverride",
                })
            unknown_pattern_keys = set(numbering_pattern) - allowed_pattern_keys
            if unknown_pattern_keys:
                raise ValueError(
                    f"roles['{role}'].numbering_pattern contains unknown fields: "
                    f"{sorted(unknown_pattern_keys)}"
                )
            for field, value in numbering_pattern.items():
                if not isinstance(value, str):
                    raise ValueError(
                        f"roles['{role}'].numbering_pattern.{field} must be a string"
                    )
        if version == 2 and provenance in {"style_numpr", "direct_numpr"}:
            if not isinstance(numbering_pattern, dict) or not numbering_pattern.get("numId"):
                raise ValueError(
                    f"roles['{role}'] provenance {provenance} requires numbering_pattern.numId"
                )
        elif version == 2 and numbering_pattern is not None:
            raise ValueError(
                f"roles['{role}'] provenance {provenance} must not define numbering_pattern"
            )


# ---------------------------------------------------------------------------
# Cross-registry consistency
# ---------------------------------------------------------------------------

def validate_cross_registry(
    style_registry: Dict[str, Any],
    template_registry: Dict[str, Any],
) -> None:
    """Every role style must be a source style or a canonical generated role style.

    Generated ``CSI_*__ARCH`` role styles intentionally live only in
    portable_styles.xml, so the source-only template registry cannot contain
    them yet.
    """
    template_ids: Set[str] = set()
    style_defs = template_registry.get("styles", {}).get("style_defs", [])
    for sdef in style_defs:
        sid = sdef.get("style_id")
        if sid:
            template_ids.add(sid)

    roles = style_registry.get("roles", {})
    allows_portable_generated_styles = style_registry.get("version") == 2
    missing: List[str] = []
    for role, spec in roles.items():
        sid = spec.get("style_id", "")
        if sid and sid not in template_ids and not (
            allows_portable_generated_styles and sid in _ALLOWED_ARCH_STYLE_IDS
        ):
            missing.append(f"  role '{role}' references style_id '{sid}'")

    if missing:
        raise ValueError(
            "Style registry references style_id(s) not found in template registry "
            "or declared as version 2 canonical generated ARCH styles:\n" + "\n".join(missing)
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
