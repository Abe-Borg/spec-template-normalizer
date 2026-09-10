"""The committed Canadian CSC PageFormat numbering scheme.

Canadian conversion never invents numbering: it retargets a target's
paragraphs onto a multilevel Word list that something else already proved.
With an architect template that proof comes from the architect's own
``numbering.xml``.  This module is the other source -- a nine-level CSC
PageFormat list built from constants in this file, for runs where no architect
enters the picture.

Two properties make that safe enough to publish from:

* It is **generated, not analyzed**.  No API call, no cache, no disk
  round-trip, so there is no gap between deciding what the scheme is and using
  it, and nothing to checksum against tampering.  That is why this is not
  dressed up as a synthetic ``.phase1`` bundle: a manifest asserting a
  classifier and prompt hashes that never ran would make the manifest mean
  less, not more.
* It is **validated, not trusted**.  ``tests/test_builtin_scheme.py`` runs the
  generated role specs and numbering through the *unmodified* architect
  validators in ``core/csi_to_canadian.py``.  The built-in scheme has to earn
  its way past exactly the contract a real architect template must satisfy, so
  an edit here that breaks the Canadian hierarchy fails a test rather than a
  user's document.

The level shapes below are what those validators require:

===== ======================== ============ ==================
ilvl  role                     ``lvlText``  rendered
===== ======================== ============ ==================
0     ``PART``                 ``PART %1``  ``PART 1``
1     ``ARTICLE``              ``%1.%2``    ``1.1``
2     ``PARAGRAPH``            ``.%3``      ``.1``
3     ``SUBPARAGRAPH``         ``.%4``      ``.1``
4     ``SUBSUBPARAGRAPH``      ``.%5``      ``.1``
5-8   ``SUBPARAGRAPH_LEVEL_*``  ``.%6``-``.%9``  ``.1``
===== ======================== ============ ==================

Every level is ``decimal``, starts at 1, and declares no ``lvlRestart``,
because the converter refuses source sequences it cannot prove and must not be
handed a scheme whose counters it could not prove either.
"""

from __future__ import annotations

import hashlib
import json
import xml.etree.ElementTree as ET
from typing import Any, Dict, List

from .role_contract import (
    BODY_HIERARCHY_ROLES,
    ROLE_LEVEL,
    ROLE_ORDER,
    ROLE_TO_ARCH_STYLE,
    ROLE_TO_STYLE_NAME,
)


#: Bump when the rendered numbering or the role styles change in a way a user
#: would see. Recorded in ``run.json`` beside the digest.
BUILTIN_SCHEME_VERSION = "1"

#: Word list identifiers for the generated list. ``numbering_importer`` remaps
#: these on import when the target already uses them, so they only need to be
#: internally consistent.
BUILTIN_ABSTRACT_NUM_ID = "900"
BUILTIN_NUM_ID = "900"

#: Left indent of the first numbered level, and the step added per level, in
#: twentieths of a point (1440 = one inch).
_INDENT_STEP_TWIPS = 720

_W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"

_STYLES_OPEN = (
    '<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
    ' xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"'
    ' xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"'
    ' xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml">'
)
_NUMBERING_OPEN = (
    '<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
    ' xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"'
    ' xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
)
_XML_DECL = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'

#: Roles that carry no automatic numbering. They still get a style so the
#: section identity and the closing line are formatted consistently.
_UNNUMBERED_ROLES = ("SectionID", "SectionTitle", "END_OF_SECTION")


def _level_text(role: str) -> str:
    """The ``lvlText`` the Canadian contract requires for *role*."""

    if role == "PART":
        return "PART %1"
    if role == "ARTICLE":
        return "%1.%2"
    return f".%{ROLE_LEVEL[role] + 1}"


def _level_indent(level: int) -> tuple[int, int]:
    """Return ``(left, hanging)`` twips for *level*.

    ``PART`` sits flush left with its title after a tab; every level below it
    steps in by one ``_INDENT_STEP_TWIPS`` and hangs its marker back out, which
    is how a spec section reads on the page.
    """

    if level == 0:
        return 0, 0
    return _INDENT_STEP_TWIPS * level, _INDENT_STEP_TWIPS


def _spacing(role: str) -> str:
    if role in {"PART", "ARTICLE", "SectionID", "SectionTitle"}:
        return '<w:spacing w:before="240" w:after="120"/>'
    return '<w:spacing w:before="0" w:after="120"/>'


def build_numbering_xml() -> str:
    """Return the generated ``word/numbering.xml`` for the built-in scheme."""

    parts: List[str] = [
        _XML_DECL,
        _NUMBERING_OPEN,
        f'<w:abstractNum w:abstractNumId="{BUILTIN_ABSTRACT_NUM_ID}">',
        '<w:multiLevelType w:val="multilevel"/>',
        '<w:name w:val="CSC PageFormat (built-in)"/>',
    ]
    for role in BODY_HIERARCHY_ROLES:
        level = ROLE_LEVEL[role]
        left, hanging = _level_indent(level)
        parts.extend(
            [
                f'<w:lvl w:ilvl="{level}">',
                '<w:start w:val="1"/>',
                '<w:numFmt w:val="decimal"/>',
                f'<w:lvlText w:val="{_level_text(role)}"/>',
                '<w:lvlJc w:val="left"/>',
                "<w:pPr>",
                f'<w:ind w:left="{left}" w:hanging="{hanging}"/>',
                "</w:pPr>",
                "</w:lvl>",
            ]
        )
    parts.extend(
        [
            "</w:abstractNum>",
            f'<w:num w:numId="{BUILTIN_NUM_ID}">',
            f'<w:abstractNumId w:val="{BUILTIN_ABSTRACT_NUM_ID}"/>',
            "</w:num>",
            "</w:numbering>",
        ]
    )
    return _validated("\n".join(parts), "numbering.xml")


def build_styles_xml() -> str:
    """Return the generated stylesheet holding the twelve CSI role styles.

    Deliberately free of ``rPr``: no font, size, or colour is imposed. The
    target keeps its own theme and document defaults, which is what makes
    "the shell is left untouched" true on the page and not merely true of the
    headers and footers.
    """

    parts: List[str] = [
        _XML_DECL,
        _STYLES_OPEN,
        "<w:docDefaults>",
        "<w:rPrDefault><w:rPr/></w:rPrDefault>",
        "<w:pPrDefault><w:pPr/></w:pPrDefault>",
        "</w:docDefaults>",
    ]
    for role in ROLE_ORDER:
        style_id = ROLE_TO_ARCH_STYLE[role]
        parts.extend(
            [
                f'<w:style w:type="paragraph" w:styleId="{style_id}">',
                f'<w:name w:val="{ROLE_TO_STYLE_NAME[role]}"/>',
                "<w:qFormat/>",
                "<w:pPr>",
            ]
        )
        if role in BODY_HIERARCHY_ROLES:
            level = ROLE_LEVEL[role]
            left, hanging = _level_indent(level)
            parts.append(
                "<w:numPr>"
                f'<w:ilvl w:val="{level}"/>'
                f'<w:numId w:val="{BUILTIN_NUM_ID}"/>'
                "</w:numPr>"
            )
            parts.append(f'<w:ind w:left="{left}" w:hanging="{hanging}"/>')
        parts.append(_spacing(role))
        parts.extend(["</w:pPr>", "</w:style>"])
    parts.append("</w:styles>")
    return _validated("\n".join(parts), "styles.xml")


def build_role_specs() -> Dict[str, Dict[str, Any]]:
    """Return role contracts in the shape ``load_role_specs_from_registry`` gives.

    ``exemplar_paragraph_index`` is 0 for every role: the field records which
    architect paragraph demonstrated a role, and the built-in scheme has no
    document behind it. Nothing in the Canadian path reads it, and the value is
    never published as though a real exemplar existed.
    """

    specs: Dict[str, Dict[str, Any]] = {}
    for role in ROLE_ORDER:
        spec: Dict[str, Any] = {
            "style_id": ROLE_TO_ARCH_STYLE[role],
            "style_name": ROLE_TO_STYLE_NAME[role],
            "exemplar_paragraph_index": 0,
        }
        if role in BODY_HIERARCHY_ROLES:
            spec["numbering_provenance"] = "style_numpr"
            spec["numbering_pattern"] = {
                "numId": BUILTIN_NUM_ID,
                "ilvl": str(ROLE_LEVEL[role]),
                "abstractNumId": BUILTIN_ABSTRACT_NUM_ID,
                "start": "1",
                "numFmt": "decimal",
                "lvlText": _level_text(role),
            }
        else:
            spec["numbering_provenance"] = "none"
        specs[role] = spec
    return specs


def _abstract_num_xml() -> str:
    """The ``w:abstractNum`` element alone, as the registry stores it."""

    whole = build_numbering_xml()
    start = whole.index("<w:abstractNum ")
    end = whole.index("</w:abstractNum>") + len("</w:abstractNum>")
    return whole[start:end]


def _num_xml() -> str:
    """The ``w:num`` element alone, as the registry stores it."""

    whole = build_numbering_xml()
    start = whole.index("<w:num ")
    end = whole.index("</w:num>") + len("</w:num>")
    return whole[start:end]


def build_style_defs() -> List[Dict[str, Any]]:
    """Registry-shaped records for the generated role styles."""

    defs: List[Dict[str, Any]] = []
    for role in ROLE_ORDER:
        style_id = ROLE_TO_ARCH_STYLE[role]
        block = build_styles_xml()
        opening = f'<w:style w:type="paragraph" w:styleId="{style_id}">'
        start = block.index(opening)
        end = block.index("</w:style>", start) + len("</w:style>")
        body = block[start:end]
        ppr_start = body.index("<w:pPr>")
        ppr_end = body.index("</w:pPr>") + len("</w:pPr>")
        defs.append(
            {
                "style_id": style_id,
                "name": ROLE_TO_STYLE_NAME[role],
                "type": "paragraph",
                "qformat": True,
                "pPr": body[ppr_start:ppr_end],
            }
        )
    return defs


def build_env_registry() -> Dict[str, Any]:
    """A template-environment registry carrying numbering and styles only.

    Every shell section -- theme, settings, page layout, headers and footers --
    is deliberately empty. These modes do not apply an architect shell, and a
    registry that carried one would invite a later change to start applying it
    without anyone deciding to.
    """

    return {
        "meta": {
            "schema_version": "1.0.0",
            "source_docx": {
                "filename": f"built-in CSC PageFormat v{BUILTIN_SCHEME_VERSION}",
                "sha256": scheme_digest(),
            },
        },
        "package_inventory": {
            "has_numbering": True,
            "has_styles": True,
            "has_theme": False,
            "has_settings": False,
            "has_header_parts": False,
            "has_footer_parts": False,
            "has_footnotes": False,
            "has_endnotes": False,
        },
        "doc_defaults": {
            "default_run_props": {"rPr": ""},
            "default_paragraph_props": {"pPr": ""},
        },
        "styles": {"style_defs": build_style_defs(), "table_styles": []},
        "numbering": {
            "numbering_xml": build_numbering_xml(),
            "abstract_nums": [
                {
                    "abstractNumId": int(BUILTIN_ABSTRACT_NUM_ID),
                    "xml": _abstract_num_xml(),
                }
            ],
            "nums": [
                {
                    "numId": int(BUILTIN_NUM_ID),
                    "abstractNumId": int(BUILTIN_ABSTRACT_NUM_ID),
                    "xml": _num_xml(),
                }
            ],
        },
        "page_layout": {"default_section": {}, "section_chain": []},
        "headers_footers": {"headers": [], "footers": [], "header_footer_media": []},
        "theme": {},
        "settings": {},
    }


def build_arch_registry() -> Dict[str, str]:
    """Return the role -> style-id map for the built-in scheme."""

    return {role: ROLE_TO_ARCH_STYLE[role] for role in ROLE_ORDER}


def available_roles() -> List[str]:
    """Every role the built-in scheme can style."""

    return list(ROLE_ORDER)


def scheme_digest() -> str:
    """SHA-256 over the generated parts, recorded as the run's provenance.

    It stands where an architect template's source hash would, so a published
    run always names what its numbering came from.
    """

    payload = json.dumps(
        {
            "version": BUILTIN_SCHEME_VERSION,
            "numbering": build_numbering_xml(),
            "styles": build_styles_xml(),
            "roles": build_role_specs(),
        },
        sort_keys=True,
        separators=(",", ":"),
    )
    return hashlib.sha256(payload.encode("utf-8")).hexdigest()


def _validated(xml_text: str, part_name: str) -> str:
    try:
        ET.fromstring(xml_text.encode("utf-8"))
    except ET.ParseError as exc:  # pragma: no cover - a constant-only bug
        raise ValueError(
            f"Built-in scheme produced malformed {part_name}: {exc}"
        ) from exc
    return xml_text


__all__ = [
    "BUILTIN_ABSTRACT_NUM_ID",
    "BUILTIN_NUM_ID",
    "BUILTIN_SCHEME_VERSION",
    "available_roles",
    "build_arch_registry",
    "build_env_registry",
    "build_numbering_xml",
    "build_role_specs",
    "build_style_defs",
    "build_styles_xml",
    "scheme_digest",
]
