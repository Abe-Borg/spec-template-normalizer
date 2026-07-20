#!/usr/bin/env python3
"""
numbering_importer.py

Imports architect's numbering definitions (abstractNum + num) into target document's
numbering.xml, handling ID collisions by remapping.

This allows imported styles to reference the architect's exact numbering definitions,
preserving list number formatting (fonts, indents, prefixes).
"""

import re
import json
import hashlib
import xml.etree.ElementTree as ET
from pathlib import Path
from typing import Dict, List, Tuple, Optional, Any
from copy import deepcopy

from .core.ooxml_namespaces import (
    CT_NS,
    PKG_REL_NS,
    W_NS,
    serialize_content_types,
    serialize_package_relationships,
)
from .core.ooxml_text import read_xml_text, write_xml_text
from .core.style_import import (
    _find_style_numpr_in_chain,
    collect_style_dependency_closure,
)


_MINIMAL_NUMBERING_XML = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    f'<w:numbering xmlns:w="{W_NS}"></w:numbering>'
)
_NUMBERING_CONTENT_TYPE = (
    "application/vnd.openxmlformats-officedocument."
    "wordprocessingml.numbering+xml"
)
_NUMBERING_REL_TYPE = (
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering"
)


def _ensure_numbering_package_wiring(target_extract_dir: Path, log: List[str]) -> None:
    """Add the OPC declarations required by a newly-created numbering part."""
    content_types_path = target_extract_dir / "[Content_Types].xml"
    document_rels_path = target_extract_dir / "word" / "_rels" / "document.xml.rels"
    if not content_types_path.is_file():
        raise FileNotFoundError(
            f"Cannot create word/numbering.xml without required package part: {content_types_path}"
        )
    if not document_rels_path.is_file():
        raise FileNotFoundError(
            f"Cannot create word/numbering.xml without required package part: {document_rels_path}"
        )

    try:
        content_types_root = ET.fromstring(content_types_path.read_bytes())
    except ET.ParseError as exc:
        raise ValueError(f"Invalid [Content_Types].xml: {exc}") from exc
    if content_types_root.tag != f"{{{CT_NS}}}Types":
        raise ValueError("Invalid [Content_Types].xml root element")

    numbering_overrides = [
        node
        for node in content_types_root.findall(f"{{{CT_NS}}}Override")
        if node.attrib.get("PartName") == "/word/numbering.xml"
    ]
    if len(numbering_overrides) > 1:
        raise ValueError("[Content_Types].xml contains duplicate numbering.xml overrides")
    if numbering_overrides:
        numbering_overrides[0].set("ContentType", _NUMBERING_CONTENT_TYPE)
    else:
        ET.SubElement(
            content_types_root,
            f"{{{CT_NS}}}Override",
            {
                "PartName": "/word/numbering.xml",
                "ContentType": _NUMBERING_CONTENT_TYPE,
            },
        )

    try:
        document_rels_root = ET.fromstring(document_rels_path.read_bytes())
    except ET.ParseError as exc:
        raise ValueError(f"Invalid document.xml.rels: {exc}") from exc
    if document_rels_root.tag != f"{{{PKG_REL_NS}}}Relationships":
        raise ValueError("Invalid document.xml.rels root element")

    numbering_rels = [
        node
        for node in document_rels_root.findall(f"{{{PKG_REL_NS}}}Relationship")
        if node.attrib.get("Type", "").endswith("/numbering")
    ]
    if len(numbering_rels) > 1:
        raise ValueError("document.xml.rels contains multiple numbering relationships")
    if numbering_rels:
        numbering_rels[0].set("Target", "numbering.xml")
        numbering_rels[0].attrib.pop("TargetMode", None)
    else:
        existing_ids = {
            node.attrib.get("Id", "")
            for node in document_rels_root.findall(f"{{{PKG_REL_NS}}}Relationship")
        }
        numeric_rids = [
            int(match.group(1))
            for rid in existing_ids
            if (match := re.fullmatch(r"rId(\d+)", rid))
        ]
        next_rid = (max(numeric_rids) if numeric_rids else 0) + 1
        while f"rId{next_rid}" in existing_ids:
            next_rid += 1
        ET.SubElement(
            document_rels_root,
            f"{{{PKG_REL_NS}}}Relationship",
            {
                "Id": f"rId{next_rid}",
                "Type": _NUMBERING_REL_TYPE,
                "Target": "numbering.xml",
            },
        )

    # Both source package parts are parsed and validated before either is
    # changed, so malformed base wiring cannot leave a half-created part.
    content_types_path.write_bytes(serialize_content_types(content_types_root))
    document_rels_path.write_bytes(serialize_package_relationships(document_rels_root))
    log.append("Added numbering.xml content type and document relationship")


def _canonicalize_xml(xml_content: str) -> str:
    return re.sub(r"\s+", " ", xml_content or "").strip()


def _generate_unique_nsid(abstract_num_xml_content: str) -> str:
    """Generate deterministic nsid (8 hex chars) for abstractNum."""
    digest = hashlib.sha1(_canonicalize_xml(abstract_num_xml_content).encode("utf-8")).hexdigest()
    return digest[:8].upper()


def _generate_unique_durable_id(num_xml_content: str) -> str:
    """Generate deterministic durableId for num."""
    digest = hashlib.sha1(_canonicalize_xml(num_xml_content).encode("utf-8")).hexdigest()
    return str((int(digest[:8], 16) % 2147483646) + 1)


def _generate_collision_safe_nsid(abstract_num_xml: str, target_numbering_xml: str) -> str:
    existing = set(re.findall(r'<w:nsid\s+w:val="([^"]+)"', target_numbering_xml or ""))
    candidate = _generate_unique_nsid(abstract_num_xml)
    if candidate not in existing:
        return candidate
    for i in range(1, 1000):
        candidate = _generate_unique_nsid(abstract_num_xml + f"__collision_{i}")
        if candidate not in existing:
            return candidate
    raise ValueError("Could not generate unique nsid after 1000 attempts")


def _generate_collision_safe_durable_id(num_xml: str, target_numbering_xml: str) -> str:
    existing = set(re.findall(r'w16cid:durableId="([^"]+)"', target_numbering_xml or ""))
    candidate = _generate_unique_durable_id(num_xml)
    if candidate not in existing:
        return candidate
    for i in range(1, 1000):
        candidate = _generate_unique_durable_id(num_xml + f"__collision_{i}")
        if candidate not in existing:
            return candidate
    raise ValueError("Could not generate unique durableId after 1000 attempts")


def find_max_ids_in_numbering(numbering_xml: str) -> Tuple[int, int]:
    """
    Find the maximum abstractNumId and numId in existing numbering.xml.
    Returns (max_abstract_num_id, max_num_id).
    """
    abstract_ids = [int(m) for m in re.findall(r'w:abstractNumId="(\d+)"', numbering_xml)]
    num_ids = [int(m) for m in re.findall(r'<w:num\s+w:numId="(\d+)"', numbering_xml)]
    
    max_abstract = max(abstract_ids) if abstract_ids else -1
    max_num = max(num_ids) if num_ids else 0
    
    return max_abstract, max_num


def extract_used_num_ids_from_styles(styles_xml: str) -> Dict[str, int]:
    """
    Extract which numIds are referenced by which styles.
    Returns dict of styleId -> numId.
    """
    result = {}
    # Find all style definitions with numPr
    style_pattern = r'<w:style[^>]*w:styleId="([^"]+)"[^>]*>[\s\S]*?</w:style>'
    for match in re.finditer(style_pattern, styles_xml):
        style_xml = match.group(0)
        style_id = match.group(1)
        
        # Look for numId in this style
        num_match = re.search(r'<w:numId\s+w:val="(\d+)"', style_xml)
        if num_match:
            result[style_id] = int(num_match.group(1))
    
    return result


def _numbering_signature_from_registry(
    numbering: Dict[str, Any],
    num_id: int,
    ilvl: str,
) -> Dict[str, str]:
    nums = {item.get("numId"): item for item in numbering.get("nums", []) if isinstance(item, dict)}
    abstracts = {
        item.get("abstractNumId"): item
        for item in numbering.get("abstract_nums", [])
        if isinstance(item, dict)
    }
    num = nums.get(num_id)
    if not isinstance(num, dict):
        return {}
    abstract_id = num.get("abstractNumId")
    signature = {"abstractNumId": str(abstract_id)} if isinstance(abstract_id, int) else {}
    abstract = abstracts.get(abstract_id)
    abstract_xml = abstract.get("xml", "") if isinstance(abstract, dict) else ""
    level_match = re.search(
        rf'<w:lvl\b[^>]*w:ilvl="{re.escape(ilvl)}"[^>]*>[\s\S]*?</w:lvl>',
        abstract_xml,
    )
    level_xml = level_match.group(0) if level_match else ""
    for tag, key in (("numFmt", "numFmt"), ("lvlText", "lvlText")):
        match = re.search(rf'<w:{tag}\b[^>]*w:val="([^"]+)"', level_xml)
        if match:
            signature[key] = match.group(1)

    num_xml = num.get("xml", "")
    override_match = re.search(
        rf'<w:lvlOverride\b[^>]*w:ilvl="{re.escape(ilvl)}"[^>]*>[\s\S]*?</w:lvlOverride>',
        num_xml,
    )
    if override_match:
        override_xml = override_match.group(0)
        start = re.search(r'<w:startOverride\b[^>]*w:val="([^"]+)"', override_xml)
        if start:
            signature["startOverride"] = start.group(1)
        for tag, key in (("numFmt", "numFmt"), ("lvlText", "lvlText")):
            match = re.search(rf'<w:{tag}\b[^>]*w:val="([^"]+)"', override_xml)
            if match:
                signature[key] = match.group(1)
    return signature


def build_numbering_import_plan(
    arch_template_registry: Dict[str, Any],
    arch_styles_xml: str,
    target_numbering_xml: str,
    style_ids_to_import: List[str],
    role_specs: Optional[Dict[str, Dict[str, Any]]] = None,
    roles_to_apply: Optional[List[str]] = None,
    additional_num_ids: Optional[List[int]] = None,
) -> Dict[str, Any]:
    """
    Build a plan for importing numbering definitions.
    
    Returns:
    {
        "abstract_nums_to_import": [
            {"old_id": 6, "new_id": 15, "xml": "..."},
            ...
        ],
        "nums_to_import": [
            {"old_id": 2, "new_id": 12, "old_abstract_id": 6, "new_abstract_id": 15, "xml": "..."},
            ...
        ],
        "style_numid_remap": {
            "CSILevel1": {"old_numId": 2, "new_numId": 12},
            ...
        }
    }
    """
    # Numbering can be inherited through a style dependency.  Import numbering
    # for the complete closure, because style import brings that same closure
    # into the target document.
    expanded_style_ids = collect_style_dependency_closure(
        arch_styles_xml,
        style_ids_to_import,
    )

    # Find which numIds the styles we're importing reference
    style_to_numid = extract_used_num_ids_from_styles(arch_styles_xml)
    
    # Filter to only styles we're importing
    relevant_numids = set()
    style_numid_usage = {}
    for style_id in expanded_style_ids:
        if style_id in style_to_numid:
            num_id = style_to_numid[style_id]
            relevant_numids.add(num_id)
            style_numid_usage[style_id] = num_id

    if additional_num_ids:
        for num_id in additional_num_ids:
            if type(num_id) is not int or num_id <= 0:
                raise ValueError(f"Invalid additional numId requirement: {num_id!r}")
            relevant_numids.add(num_id)
    
    role_numid_usage: Dict[str, Tuple[int, str, Dict[str, str]]] = {}
    if role_specs and roles_to_apply:
        for role in roles_to_apply:
            spec = role_specs.get(role)
            if not isinstance(spec, dict):
                raise ValueError(f"Missing full role contract for classified role: {role}")
            provenance = spec.get("numbering_provenance")
            pattern = spec.get("numbering_pattern")
            if provenance in {"style_numpr", "direct_numpr"}:
                if not isinstance(pattern, dict) or not re.fullmatch(r"\d+", pattern.get("numId", "")):
                    raise ValueError(
                        f"Role {role} provenance {provenance} requires a decimal numbering_pattern.numId"
                    )
                ilvl = pattern.get("ilvl", "0")
                if not re.fullmatch(r"\d+", ilvl):
                    raise ValueError(f"Role {role} has invalid numbering_pattern.ilvl: {ilvl!r}")
                pattern_num_id = int(pattern["numId"])
                relevant_numids.add(pattern_num_id)
                if provenance == "style_numpr":
                    style_id = spec.get("style_id", "")
                    numpr = _find_style_numpr_in_chain(arch_styles_xml, style_id)
                    actual_match = re.search(r'<w:numId\b[^>]*w:val="(\d+)"', numpr or "")
                    if not actual_match or int(actual_match.group(1)) != pattern_num_id:
                        raise ValueError(
                            f"Role {role} numbering pattern does not match style {style_id!r} numPr"
                        )
                else:
                    role_numid_usage[role] = (pattern_num_id, ilvl, dict(pattern))
            elif provenance not in {"text_literal", "none"}:
                raise ValueError(f"Role {role} has invalid numbering_provenance: {provenance!r}")

    if not relevant_numids:
        return {
            "abstract_nums_to_import": [],
            "nums_to_import": [],
            "style_numid_remap": {},
            "role_numpr_remap": {},
            "num_id_remap": {},
        }

    # Get the numbering data from arch_template_registry
    numbering = arch_template_registry.get("numbering", {})
    abstract_nums = {an["abstractNumId"]: an for an in numbering.get("abstract_nums", [])}
    nums = {n["numId"]: n for n in numbering.get("nums", [])}

    # --- Fail-fast: every referenced numId must exist in the registry ---
    missing_nums = sorted(relevant_numids - set(nums.keys()))
    if missing_nums:
        styles_for_missing = [
            f"{sid} -> numId {nid}"
            for sid, nid in style_numid_usage.items()
            if nid in missing_nums
        ]
        raise ValueError(
            f"Architect registry is missing required numId definitions: {missing_nums}. "
            f"Referenced by styles: {styles_for_missing}"
        )

    # First, determine which abstractNums we need (referenced by the nums we need)
    needed_abstract_ids = set()
    for num_id in relevant_numids:
        needed_abstract_ids.add(nums[num_id]["abstractNumId"])

    # --- Fail-fast: every referenced abstractNumId must exist in the registry ---
    missing_abstracts = sorted(needed_abstract_ids - set(abstract_nums.keys()))
    if missing_abstracts:
        raise ValueError(
            f"Architect registry is missing required abstractNum definitions: "
            f"{missing_abstracts}. Referenced by numIds: "
            f"{sorted(nid for nid in relevant_numids if nums[nid]['abstractNumId'] in missing_abstracts)}"
        )

    for role, (num_id, ilvl, pattern) in role_numid_usage.items():
        actual_signature = _numbering_signature_from_registry(numbering, num_id, ilvl)
        for key in ("abstractNumId", "startOverride", "numFmt", "lvlText"):
            if key in pattern and pattern[key] != actual_signature.get(key):
                raise ValueError(
                    f"Role {role} numbering_pattern.{key}={pattern[key]!r} does not match "
                    f"the template numbering definition {actual_signature.get(key)!r}"
                )

    # Find max IDs in target to avoid collisions
    max_abstract_id, max_num_id = find_max_ids_in_numbering(target_numbering_xml)

    # Build import lists
    abstract_num_id_remap = {}  # old_id -> new_id
    num_id_remap = {}  # old_id -> new_id

    abstract_nums_to_import = []
    nums_to_import = []
    
    # Assign new IDs to abstractNums (all validated to exist above)
    next_abstract_id = max_abstract_id + 1
    for old_abstract_id in sorted(needed_abstract_ids):
        new_abstract_id = next_abstract_id
        abstract_num_id_remap[old_abstract_id] = new_abstract_id

        # Get XML and remap the abstractNumId
        xml = abstract_nums[old_abstract_id]["xml"]
        source_xml_for_hash = xml
        xml = re.sub(
            r'w:abstractNumId="' + str(old_abstract_id) + '"',
            f'w:abstractNumId="{new_abstract_id}"',
            xml
        )
        # Generate new nsid to avoid conflicts
        xml = re.sub(
            r'<w:nsid\s+w:val="[^"]+"/>',
            f'<w:nsid w:val="{_generate_collision_safe_nsid(source_xml_for_hash, target_numbering_xml)}"/>',
            xml
        )

        abstract_nums_to_import.append({
            "old_id": old_abstract_id,
            "new_id": new_abstract_id,
            "xml": xml
        })
        next_abstract_id += 1

    # Assign new IDs to nums (all validated to exist above)
    next_num_id = max_num_id + 1
    for old_num_id in sorted(relevant_numids):
        new_num_id = next_num_id
        num_id_remap[old_num_id] = new_num_id

        num_data = nums[old_num_id]
        old_abstract_id = num_data["abstractNumId"]
        new_abstract_id = abstract_num_id_remap.get(old_abstract_id, old_abstract_id)

        # Get XML and remap IDs
        xml = num_data["xml"]
        source_xml_for_hash = xml
        xml = re.sub(
            r'w:numId="' + str(old_num_id) + '"',
            f'w:numId="{new_num_id}"',
            xml
        )
        xml = re.sub(
            r'<w:abstractNumId\s+w:val="' + str(old_abstract_id) + '"',
            f'<w:abstractNumId w:val="{new_abstract_id}"',
            xml
        )
        # Generate new durableId
        xml = re.sub(
            r'w16cid:durableId="[^"]*"',
            f'w16cid:durableId="{_generate_collision_safe_durable_id(source_xml_for_hash, target_numbering_xml)}"',
            xml
        )

        nums_to_import.append({
            "old_id": old_num_id,
            "new_id": new_num_id,
            "old_abstract_id": old_abstract_id,
            "new_abstract_id": new_abstract_id,
            "xml": xml
        })
        next_num_id += 1
    
    # Build style remap
    style_numid_remap = {}
    for style_id, old_num_id in style_numid_usage.items():
        if old_num_id in num_id_remap:
            style_numid_remap[style_id] = {
                "old_numId": old_num_id,
                "new_numId": num_id_remap[old_num_id]
            }

    role_numpr_remap = {}
    for role, (old_num_id, ilvl, pattern) in role_numid_usage.items():
        role_numpr_remap[role] = {
            "old_numId": old_num_id,
            "new_numId": num_id_remap[old_num_id],
            "ilvl": int(ilvl),
            "numbering_pattern": pattern,
        }
    
    return {
        "abstract_nums_to_import": abstract_nums_to_import,
        "nums_to_import": nums_to_import,
        "style_numid_remap": style_numid_remap,
        "role_numpr_remap": role_numpr_remap,
        "num_id_remap": num_id_remap,
    }


def inject_numbering_into_xml(
    target_numbering_xml: str,
    abstract_nums_to_import: List[Dict],
    nums_to_import: List[Dict],
    source_numbering_xml: Optional[str] = None,
) -> str:
    """
    Inject imported abstractNums and nums into target numbering.xml.

    abstractNums go before the first <w:num> element.
    nums go at the end, before </w:numbering>.

    Architect numbering XML is preserved exactly as-is (no typography
    normalization).
    """
    result = target_numbering_xml
    if source_numbering_xml:
        target_open = re.search(r'<w:numbering\b[^>]*>', result)
        source_open = re.search(r'<w:numbering\b[^>]*>', source_numbering_xml)
        if target_open and source_open:
            declared = {
                match.group(1) or ""
                for match in re.finditer(r'xmlns(?::([A-Za-z_][\w.-]*))?="[^"]+"', target_open.group(0))
            }
            additions = []
            for match in re.finditer(
                r'xmlns(?::([A-Za-z_][\w.-]*))?="[^"]+"',
                source_open.group(0),
            ):
                prefix = match.group(1) or ""
                if prefix not in declared:
                    additions.append(match.group(0))
                    declared.add(prefix)
            if additions:
                merged_open = target_open.group(0)[:-1] + " " + " ".join(additions) + ">"
                result = result[:target_open.start()] + merged_open + result[target_open.end():]

    first_num_match = re.search(r'<w:num\s', result)
    end_match = re.search(r'</w:numbering>', result)
    abstract_xml = "\n".join(an["xml"] for an in abstract_nums_to_import)
    if abstract_xml:
        if first_num_match:
            insert_pos = first_num_match.start()
        elif end_match:
            insert_pos = end_match.start()
        else:
            raise ValueError("numbering.xml missing </w:numbering> closing tag")
        result = result[:insert_pos] + abstract_xml + "\n" + result[insert_pos:]

    # Find insertion point for nums (before </w:numbering>)
    end_match = re.search(r'</w:numbering>', result)
    if end_match:
        insert_pos = end_match.start()
        num_xml = "\n".join(n["xml"] for n in nums_to_import)
        if num_xml:
            result = result[:insert_pos] + num_xml + "\n" + result[insert_pos:]

    for n in nums_to_import:
        if "new_abstract_id" not in n:
            continue
        aid = str(n.get("new_abstract_id"))
        if not re.search(rf'<w:abstractNum\b[^>]*w:abstractNumId="{re.escape(aid)}"', result):
            raise ValueError(f"Injected num references missing abstractNumId={aid}")

    try:
        ET.fromstring(result.encode("utf-8"))
    except ET.ParseError as exc:
        raise ValueError(f"Imported numbering.xml is not well-formed: {exc}") from exc

    return result


def remap_numid_in_style_xml(style_xml: str, old_num_id: int, new_num_id: int) -> str:
    """
    Update a style's numPr to reference the new numId.
    """
    return re.sub(
        r'(<w:numId\s+w:val=")' + str(old_num_id) + r'"',
        f'\\g<1>{new_num_id}"',
        style_xml
    )


def import_numbering(
    target_extract_dir: Path,
    arch_template_registry: Dict[str, Any],
    arch_styles_xml: str,
    style_ids_to_import: List[str],
    log: List[str],
    role_specs: Optional[Dict[str, Dict[str, Any]]] = None,
    roles_to_apply: Optional[List[str]] = None,
    additional_num_ids: Optional[List[int]] = None,
    return_contract: bool = False,
) -> Dict[str, Any]:
    """
    Main entry point: import architect's numbering into target.

    arch_styles_xml: synthetic or real styles.xml content as a string
    (built from arch_template_registry.json via build_arch_styles_xml_from_registry).

    Returns style_numid_remap for use when importing styles.
    """
    # Determine whether any of the styles being imported actually need numbering
    expanded_style_ids = collect_style_dependency_closure(
        arch_styles_xml,
        style_ids_to_import,
    )
    style_to_numid = extract_used_num_ids_from_styles(arch_styles_xml)
    needed_num_ids = {
        nid for sid, nid in style_to_numid.items() if sid in expanded_style_ids
    }
    if role_specs and roles_to_apply:
        for role in roles_to_apply:
            spec = role_specs.get(role, {})
            if spec.get("numbering_provenance") not in {"style_numpr", "direct_numpr"}:
                continue
            pattern = spec.get("numbering_pattern", {})
            raw_num_id = pattern.get("numId") if isinstance(pattern, dict) else None
            if isinstance(raw_num_id, str) and raw_num_id.isdigit():
                needed_num_ids.add(int(raw_num_id))
    if additional_num_ids:
        for num_id in additional_num_ids:
            if type(num_id) is not int or num_id <= 0:
                raise ValueError(f"Invalid additional numId requirement: {num_id!r}")
            needed_num_ids.add(num_id)

    def _result(
        style_remap: Dict[str, Any],
        role_remap: Optional[Dict[str, Any]] = None,
        num_id_remap: Optional[Dict[int, int]] = None,
    ) -> Dict[str, Any]:
        if return_contract:
            return {
                "style_numid_remap": style_remap,
                "role_numpr_remap": role_remap or {},
                "num_id_remap": num_id_remap or {},
            }
        return style_remap

    # Check if registry has numbering data
    if "numbering" not in arch_template_registry:
        if needed_num_ids:
            raise ValueError(
                f"Architect registry has no numbering data but imported styles require "
                f"numbering definitions (numIds: {sorted(needed_num_ids)})."
            )
        log.append("No numbering data in arch_template_registry, skipping numbering import")
        return _result({})

    numbering_data = arch_template_registry.get("numbering", {})
    if not numbering_data.get("abstract_nums") and not numbering_data.get("nums"):
        if needed_num_ids:
            raise ValueError(
                f"Architect registry has empty numbering definitions but imported styles "
                f"require numbering (numIds: {sorted(needed_num_ids)})."
            )
        log.append("No numbering definitions in arch_template_registry")
        return _result({})

    # Read target's numbering.xml. A valid DOCX is allowed to omit this part;
    # create it when architect styles, direct roles, or header/footer content
    # actually require numbering.
    target_numbering_path = target_extract_dir / "word" / "numbering.xml"
    target_had_numbering = target_numbering_path.exists()
    if not target_had_numbering:
        if not needed_num_ids:
            log.append("Target has no numbering.xml and no styles need numbering, skipping")
            return _result({})
        target_numbering_xml = _MINIMAL_NUMBERING_XML
    else:
        target_numbering_xml = read_xml_text(target_numbering_path)
    
    # Build import plan
    plan = build_numbering_import_plan(
        arch_template_registry,
        arch_styles_xml,
        target_numbering_xml,
        style_ids_to_import,
        role_specs=role_specs,
        roles_to_apply=roles_to_apply,
        additional_num_ids=additional_num_ids,
    )
    
    if not plan["abstract_nums_to_import"] and not plan["nums_to_import"]:
        log.append("No numbering definitions need to be imported")
        return _result(
            plan.get("style_numid_remap", {}),
            plan.get("role_numpr_remap", {}),
            plan.get("num_id_remap", {}),
        )
    
    # Log what we're importing
    log.append(f"Importing {len(plan['abstract_nums_to_import'])} abstractNum definitions")
    log.append(f"Importing {len(plan['nums_to_import'])} num definitions")
    for num in plan["nums_to_import"]:
        log.append(f"  numId {num['old_id']} -> {num['new_id']} (abstractNum {num['old_abstract_id']} -> {num['new_abstract_id']})")
    
    # Inject into target numbering.xml
    new_numbering_xml = inject_numbering_into_xml(
        target_numbering_xml,
        plan["abstract_nums_to_import"],
        plan["nums_to_import"],
        source_numbering_xml=numbering_data.get("numbering_xml"),
    )

    if not target_had_numbering:
        _ensure_numbering_package_wiring(target_extract_dir, log)
        target_numbering_path.parent.mkdir(parents=True, exist_ok=True)
        log.append("Created word/numbering.xml for imported numbering definitions")
    
    # Write updated numbering.xml
    write_xml_text(target_numbering_path, new_numbering_xml)
    log.append(f"Updated {target_numbering_path}")
    
    return _result(
        plan["style_numid_remap"],
        plan.get("role_numpr_remap", {}),
        plan.get("num_id_remap", {}),
    )
