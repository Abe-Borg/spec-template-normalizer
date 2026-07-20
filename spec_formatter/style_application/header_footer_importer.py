from __future__ import annotations

import base64
import hashlib
import html
import posixpath
import re
import xml.etree.ElementTree as ET
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any, Dict, List, Tuple

from .core.ooxml_namespaces import (
    CT_NS,
    PKG_REL_NS,
    R_NS,
    serialize_content_types,
    serialize_package_relationships,
)
from .core.ooxml_text import prepare_xml_text_for_utf8, read_xml_text
from .core.opc_paths import (
    is_safe_package_part_name,
    relationship_part_name_for_owner,
    relationship_target_for_part,
    resolve_internal_relationship_target,
)
from .core.section_mapping import choose_section_sources
from .core.sectpr_tools import (
    canonical_sectpr_order_index,
    child_tag_name,
    extract_all_sectpr_blocks,
    extract_sectpr_children,
    has_body_level_sectpr,
    replace_nth_sectpr_block,
    strip_tag_block,
)
from .core.xml_helpers import (
    edit_preserving_out_of_scope_subtrees,
    iter_paragraph_xml_blocks,
    strip_out_of_scope_subtrees,
)
from .docx_patch import (
    is_allowed_header_footer_part_name,
    is_allowed_header_footer_rels_name,
)


@dataclass
class HeaderFooterImportResult:
    part_names: set[str] = field(default_factory=set)
    rels_names: set[str] = field(default_factory=set)
    media_names: set[str] = field(default_factory=set)
    media_content_types: Dict[str, str] = field(default_factory=dict)
    removed_part_names: set[str] = field(default_factory=set)
    removed_rels_names: set[str] = field(default_factory=set)
    style_ids: set[str] = field(default_factory=set)
    direct_num_ids: set[int] = field(default_factory=set)


def _iter_hf_entries(registry: Dict[str, Any]) -> List[Tuple[str, Dict[str, Any]]]:
    hf_data = registry.get("headers_footers", {}) if isinstance(registry, dict) else {}
    headers = hf_data.get("headers", []) if isinstance(hf_data, dict) else []
    footers = hf_data.get("footers", []) if isinstance(hf_data, dict) else []
    entries: List[Tuple[str, Dict[str, Any]]] = []
    entries.extend(("header", e) for e in headers if isinstance(e, dict))
    entries.extend(("footer", e) for e in footers if isinstance(e, dict))
    return entries


def _resolve_media_items(entry: Dict[str, Any]) -> List[Dict[str, Any]]:
    media = entry.get("media") or entry.get("media_files") or []
    return [m for m in media if isinstance(m, dict)]


def _resolve_media_filename(media_item: Dict[str, Any]) -> str | None:
    for key in ("path", "target", "name", "filename"):
        val = media_item.get(key)
        if isinstance(val, str) and val.strip():
            path = val.strip()
            return path.split("/")[-1]
    return None


def _resolve_media_bytes(media_item: Dict[str, Any]) -> bytes | None:
    for key in ("content_base64", "base64", "data", "data_base64"):
        val = media_item.get(key)
        if isinstance(val, str) and val:
            return base64.b64decode(val)
    return None


def _remove_existing_hf_files(
    target_extract_dir: Path,
    log: List[str],
) -> Tuple[set[str], set[str]]:
    """Remove target parts identified by document relationship type."""
    rels_path = target_extract_dir / "word" / "_rels" / "document.xml.rels"
    if not rels_path.is_file():
        return set(), set()
    root = ET.fromstring(rels_path.read_bytes())
    part_names: set[str] = set()
    for rel in root.findall(f"{{{PKG_REL_NS}}}Relationship"):
        rel_type = rel.attrib.get("Type", "")
        if not rel_type.endswith(("/header", "/footer")):
            continue
        if rel.attrib.get("TargetMode", "").casefold() == "external":
            continue
        try:
            part_names.add(
                resolve_internal_relationship_target(
                    "word/document.xml",
                    rel.attrib.get("Target", ""),
                )
            )
        except ValueError as exc:
            raise ValueError(f"Unsafe target header/footer relationship: {exc}") from exc

    rels_names = {relationship_part_name_for_owner(name) for name in part_names}
    for name in sorted(part_names | rels_names):
        path = target_extract_dir / name
        if not path.exists():
            continue
        path.unlink()
        if name in part_names:
            log.append(f"Removed old part: {name}")
        else:
            log.append(f"Removed old rels: {name}")
    return part_names, rels_names


def _allocate_unique_media_name(part_name: str, index: int, original_name: str, payload: bytes, used: set[str]) -> str:
    stem = Path(part_name).stem
    suffix = hashlib.sha1(payload).hexdigest()[:8]
    ext = Path(original_name).suffix.lower() or ".bin"
    candidate = f"hf_{stem}_{index:02d}_{suffix}{ext}"
    n = 1
    while candidate.casefold() in used:
        candidate = f"hf_{stem}_{index:02d}_{suffix}_{n}{ext}"
        n += 1
    used.add(candidate.casefold())
    return candidate


_IMAGE_EXTENSION_CONTENT_TYPES = {
    "png": "image/png",
    "jpg": "image/jpeg",
    "jpeg": "image/jpeg",
    "gif": "image/gif",
    "bmp": "image/bmp",
    "tif": "image/tiff",
    "tiff": "image/tiff",
    "wmf": "image/x-wmf",
    "emf": "image/x-emf",
}


def _resolve_media_content_type(media_item: Dict[str, Any], filename: str) -> str | None:
    declared = media_item.get("content_type")
    if isinstance(declared, str) and declared.strip():
        return declared.strip()
    extension = Path(filename).suffix.lower().lstrip(".")
    return _IMAGE_EXTENSION_CONTENT_TYPES.get(extension)


def _normalize_rel_target(target: str) -> str:
    value = target.strip().replace("\\", "/")
    if value.startswith("./"):
        value = value[2:]
    return posixpath.normpath(value)


def _media_original_package_targets(media_item: Dict[str, Any]) -> set[str]:
    """Return full package-part interpretations of a registry media target.

    Current Phase 1 registries store parts below ``word/`` relative to that
    directory, while older trusted registries may already include ``word/``.
    Matching these complete normalized names avoids the old basename aliasing
    bug and still lets relationship IDs remain authoritative.
    """
    for key in ("target", "path"):
        value = media_item.get(key)
        if isinstance(value, str) and value.strip():
            normalized = _normalize_rel_target(value)
            if not is_safe_package_part_name(normalized):
                return set()
            candidates = {normalized}
            if not normalized.casefold().startswith("word/"):
                word_candidate = f"word/{normalized}"
                if is_safe_package_part_name(word_candidate):
                    candidates.add(word_candidate)
            return candidates
    return set()


def _relationship_ids_referenced_by_part(xml_content: str) -> set[str]:
    try:
        root = ET.fromstring(prepare_xml_text_for_utf8(xml_content).encode("utf-8"))
    except ET.ParseError as exc:
        raise ValueError(f"Malformed header/footer XML: {exc}") from exc
    relationship_attributes = {
        f"{{{R_NS}}}id",
        f"{{{R_NS}}}embed",
        f"{{{R_NS}}}link",
    }
    referenced: set[str] = set()
    for element in root.iter():
        for attribute, value in element.attrib.items():
            if attribute not in relationship_attributes:
                continue
            if not value:
                raise ValueError("Header/footer XML contains an empty relationship reference")
            referenced.add(value)
    return referenced


def _validate_part_relationship_references(
    *,
    part_name: str,
    xml_content: str,
    rels_xml: str | None,
) -> None:
    referenced = _relationship_ids_referenced_by_part(xml_content)
    if not referenced:
        return
    if not rels_xml:
        raise ValueError(
            f"{part_name} contains relationship references but has no relationships part"
        )
    try:
        root = ET.fromstring(prepare_xml_text_for_utf8(rels_xml).encode("utf-8"))
    except ET.ParseError as exc:
        raise ValueError(f"Malformed relationships XML for {part_name}: {exc}") from exc
    declared = {
        rel.attrib.get("Id", "")
        for rel in root.findall(f"{{{PKG_REL_NS}}}Relationship")
        if rel.attrib.get("Id")
    }
    missing = sorted(referenced - declared)
    if missing:
        raise ValueError(
            f"{part_name} references relationship IDs missing from its .rels part: {missing}"
        )


def _write_hf_parts(
    target_extract_dir: Path,
    entries: List[Tuple[str, Dict[str, Any]]],
    log: List[str],
    result: HeaderFooterImportResult,
) -> Dict[str, str]:
    part_to_type: Dict[str, str] = {}
    media_out = target_extract_dir / "word" / "media"
    media_out.mkdir(parents=True, exist_ok=True)
    written_media: set[str] = {
        p.name.casefold() for p in media_out.iterdir() if p.is_file()
    } if media_out.exists() else set()

    for kind, entry in entries:
        part_name = entry.get("part_name")
        xml_content = entry.get("xml") or entry.get("part_xml")
        if not isinstance(part_name, str) or not isinstance(xml_content, str):
            log.append(f"WARNING: Skipping malformed {kind} entry (missing part_name/xml)")
            continue
        if not is_allowed_header_footer_part_name(part_name, expected_kind=kind):
            raise ValueError(f"Unsafe or invalid {kind} part_name: {part_name!r}")

        rels_xml = entry.get("rels_xml") or entry.get("relationships_xml")
        if rels_xml is not None and not isinstance(rels_xml, str):
            raise ValueError(f"Invalid relationships XML for {part_name}")
        if isinstance(rels_xml, str):
            # Registry strings are Unicode already.  Normalize a stale source
            # declaration before every UTF-8 parse and before writing bytes.
            rels_xml = prepare_xml_text_for_utf8(rels_xml)
        _validate_part_relationship_references(
            part_name=part_name,
            xml_content=xml_content,
            rels_xml=rels_xml,
        )

        part_path = target_extract_dir / part_name
        part_path.parent.mkdir(parents=True, exist_ok=True)
        part_path.write_text(prepare_xml_text_for_utf8(xml_content), encoding="utf-8")
        result.part_names.add(part_name)
        result.style_ids.update(
            re.findall(r'<w:(?:pStyle|rStyle|tblStyle)\b[^>]*w:val="([^"]+)"', xml_content)
        )
        result.direct_num_ids.update(
            int(value)
            for value in re.findall(r'<w:numId\b[^>]*w:val="(\d+)"', xml_content)
            if int(value) != 0
        )
        part_to_type[part_name] = kind
        log.append(f"Wrote {kind} part: {part_name}")

        rels_name = entry.get("rels_part_name")
        target_by_rid: Dict[str, str] = {}
        target_by_original: Dict[str, str] = {}
        ambiguous_original_targets: set[str] = set()
        unkeyed_original_targets: set[str] = set()

        media_items = _resolve_media_items(entry)
        for idx, media_item in enumerate(media_items, start=1):
            filename = _resolve_media_filename(media_item)
            payload = _resolve_media_bytes(media_item)
            if not filename or payload is None:
                continue
            new_name = _allocate_unique_media_name(part_name, idx, filename, payload, written_media)
            media_part_name = f"word/media/{new_name}"
            out_rel = relationship_target_for_part(part_name, media_part_name)
            rel_id = media_item.get("rel_id")
            if isinstance(rel_id, str) and rel_id:
                if rel_id in target_by_rid:
                    raise ValueError(f"Duplicate captured media relationship ID {rel_id!r} in {part_name}")
                target_by_rid[rel_id] = out_rel
            original_targets = _media_original_package_targets(media_item)
            for original_target in original_targets:
                if not isinstance(rel_id, str) or not rel_id:
                    unkeyed_original_targets.add(original_target)
                if original_target in ambiguous_original_targets:
                    continue
                if original_target in target_by_original:
                    # Multiple relationship IDs may legally point at the same
                    # source part. The relationship ID remains authoritative;
                    # disable only the ambiguous legacy target-only fallback.
                    target_by_original.pop(original_target)
                    ambiguous_original_targets.add(original_target)
                else:
                    target_by_original[original_target] = out_rel
            (media_out / new_name).write_bytes(payload)
            result.media_names.add(media_part_name)
            content_type = _resolve_media_content_type(media_item, filename)
            if content_type:
                result.media_content_types[media_part_name] = content_type
            log.append(f"Wrote media asset: {media_part_name}")

        ambiguous_unkeyed = sorted(ambiguous_original_targets & unkeyed_original_targets)
        if ambiguous_unkeyed:
            raise ValueError(
                f"Captured media in {part_name} has duplicate target-only mappings; "
                f"relationship IDs are required for: {ambiguous_unkeyed}"
            )

        if isinstance(rels_xml, str):
            expected_rels_name = relationship_part_name_for_owner(part_name)
            if not isinstance(rels_name, str) or not rels_name:
                rels_name = expected_rels_name
            elif (
                rels_name != expected_rels_name
                or not is_allowed_header_footer_rels_name(
                    rels_name,
                    owner_part=part_name,
                )
            ):
                raise ValueError(
                    f"Invalid rels_part_name for {part_name}: {rels_name!r}"
                )
            if target_by_rid or target_by_original:
                rels_root = ET.fromstring(
                    prepare_xml_text_for_utf8(rels_xml).encode("utf-8")
                )
                matched_rids: set[str] = set()
                matched_targets: set[str] = set()
                for rel in rels_root.findall(f"{{{PKG_REL_NS}}}Relationship"):
                    rel_id = rel.attrib.get("Id", "")
                    old = ""
                    if rel.attrib.get("TargetMode", "").casefold() != "external":
                        try:
                            old = resolve_internal_relationship_target(
                                part_name,
                                rel.attrib.get("Target", ""),
                            )
                        except ValueError:
                            # Current bundles fail preflight before this point.
                            # Retain the precise missing-mapping failure below
                            # for trusted legacy callers.
                            old = ""
                    if rel_id in target_by_rid:
                        rel.set("Target", target_by_rid[rel_id])
                        matched_rids.add(rel_id)
                    elif old in target_by_original:
                        rel.set("Target", target_by_original[old])
                        matched_targets.add(old)
                missing_rids = sorted(set(target_by_rid) - matched_rids)
                matched_outputs = {
                    target_by_original[target]
                    for target in matched_targets
                } | {
                    target_by_rid[rid]
                    for rid in matched_rids
                }
                missing_targets = sorted(
                    target
                    for target in set(target_by_original) - matched_targets
                    if target_by_original[target] not in matched_outputs
                )
                if missing_rids or missing_targets:
                    raise ValueError(
                        f"Captured media for {part_name} does not match its relationships XML: "
                        f"missing IDs={missing_rids}, missing targets={missing_targets}"
                    )
                rels_xml = serialize_package_relationships(rels_root).decode("utf-8")
            rels_path = target_extract_dir / rels_name
            rels_path.parent.mkdir(parents=True, exist_ok=True)
            rels_path.write_text(prepare_xml_text_for_utf8(rels_xml), encoding="utf-8")
            result.rels_names.add(rels_name)
            log.append(f"Wrote rels part: {rels_name}")

    return part_to_type


def _next_rid(rels_root: ET.Element) -> int:
    rids = []
    for rel in rels_root.findall(f"{{{PKG_REL_NS}}}Relationship"):
        rid = rel.attrib.get("Id", "")
        m = re.fullmatch(r"rId(\d+)", rid)
        if m:
            rids.append(int(m.group(1)))
    return (max(rids) if rids else 0) + 1


def _rebuild_document_rels(target_extract_dir: Path, part_to_type: Dict[str, str], log: List[str]) -> Dict[str, str]:
    rels_path = target_extract_dir / "word" / "_rels" / "document.xml.rels"
    if not rels_path.exists():
        raise FileNotFoundError(f"Missing required file: {rels_path}")

    root = ET.fromstring(rels_path.read_bytes())
    for rel in list(root.findall(f"{{{PKG_REL_NS}}}Relationship")):
        rel_type = rel.attrib.get("Type", "")
        if rel_type.endswith("/header") or rel_type.endswith("/footer"):
            root.remove(rel)

    part_to_rid: Dict[str, str] = {}
    rid_num = _next_rid(root)
    for part_name, kind in sorted(part_to_type.items()):
        rid = f"rId{rid_num}"
        rid_num += 1
        target = relationship_target_for_part("word/document.xml", part_name)
        rel_type = f"http://schemas.openxmlformats.org/officeDocument/2006/relationships/{kind}"
        ET.SubElement(root, f"{{{PKG_REL_NS}}}Relationship", {"Id": rid, "Type": rel_type, "Target": target})
        part_to_rid[part_name] = rid

    rels_path.write_bytes(serialize_package_relationships(root))
    log.append(f"Rebuilt document.xml.rels header/footer relationships ({len(part_to_rid)} entries)")
    return part_to_rid


def _extract_arch_hf_refs(page_layout_section: Dict[str, Any]) -> Tuple[Dict[str, str], Dict[str, str]]:
    headers: Dict[str, str] = {}
    footers: Dict[str, str] = {}

    for key in ("header_refs", "headers"):
        val = page_layout_section.get(key)
        if isinstance(val, dict):
            headers = {k: v for k, v in val.items() if isinstance(k, str) and isinstance(v, str)}
            break
    for key in ("footer_refs", "footers"):
        val = page_layout_section.get(key)
        if isinstance(val, dict):
            footers = {k: v for k, v in val.items() if isinstance(k, str) and isinstance(v, str)}
            break

    unsupported = (set(headers) | set(footers)) - {"default", "even", "first"}
    if unsupported:
        raise ValueError(
            f"Unsupported header/footer reference type(s): {sorted(unsupported)}"
        )

    return headers, footers


def _build_arch_rid_to_part(entries: List[Tuple[str, Dict[str, Any]]]) -> Dict[str, str]:
    out: Dict[str, str] = {}
    for _kind, entry in entries:
        part_name = entry.get("part_name")
        if not isinstance(part_name, str):
            continue
        for key in ("rel_id", "rid", "rId", "relationship_id"):
            rid = entry.get(key)
            if isinstance(rid, str):
                previous = out.get(rid)
                if previous is not None and previous != part_name:
                    raise ValueError(
                        f"Architect relationship {rid!r} maps to both "
                        f"{previous!r} and {part_name!r}"
                    )
                out[rid] = part_name
    return out


def _select_reachable_entries(
    target_extract_dir: Path,
    registry: Dict[str, Any],
    entries: List[Tuple[str, Dict[str, Any]]],
    log: List[str],
) -> List[Tuple[str, Dict[str, Any]]]:
    """Return only architect header/footer parts used by mapped target sections."""
    doc_path = target_extract_dir / "word" / "document.xml"
    if not doc_path.is_file():
        raise FileNotFoundError(f"Missing required file: {doc_path}")

    document_xml = read_xml_text(doc_path)
    target_section_count = len(extract_all_sectpr_blocks(document_xml))
    if target_section_count == 0:
        # A body without an explicit final sectPr still represents one section;
        # rewiring will create the missing body-level sectPr below.
        target_section_count = 1

    page_layout = registry.get("page_layout", {}) if isinstance(registry, dict) else {}
    sources = choose_section_sources(
        target_section_count,
        page_layout,
        require_default=True,
        log=log,
    )
    rid_to_part = _build_arch_rid_to_part(entries)
    reachable_parts: set[str] = set()
    for source in sources:
        headers, footers = _extract_arch_hf_refs(source)
        for old_rid in [*headers.values(), *footers.values()]:
            part_name = rid_to_part.get(old_rid)
            if not part_name:
                raise ValueError(
                    f"Mapped architect section references unknown header/footer relationship {old_rid!r}"
                )
            reachable_parts.add(part_name)

    selected = [
        (kind, entry)
        for kind, entry in entries
        if entry.get("part_name") in reachable_parts
    ]
    omitted = len(entries) - len(selected)
    if omitted:
        log.append(f"Skipped {omitted} unreferenced architect header/footer part(s)")
    return selected


def _unique_part_entries(
    entries: List[Tuple[str, Dict[str, Any]]],
) -> List[Tuple[str, Dict[str, Any]]]:
    """Collapse multiple document relationship IDs that share one owner part."""
    unique: Dict[str, Tuple[str, str, Dict[str, Any]]] = {}
    payload_keys = ("xml", "part_xml", "rels_xml", "relationships_xml", "media", "media_files")
    for kind, entry in entries:
        part_name = entry.get("part_name")
        if not isinstance(part_name, str):
            continue
        key = part_name.casefold()
        previous = unique.get(key)
        if previous is None:
            unique[key] = (part_name, kind, entry)
            continue
        previous_name, previous_kind, previous_entry = previous
        if previous_name != part_name or previous_kind != kind or any(
            previous_entry.get(key) != entry.get(key) for key in payload_keys
        ):
            raise ValueError(
                "Architect registry repeats a header/footer owner part "
                f"case-insensitively or with conflicting content: "
                f"{previous_name!r} and {part_name!r}"
            )
    return [(kind, entry) for _name, kind, entry in unique.values()]


def _ensure_body_sectpr(document_xml: str) -> str:
    if has_body_level_sectpr(document_xml):
        return document_xml
    body_close = document_xml.rfind("</w:body>")
    if body_close < 0:
        raise ValueError("word/document.xml is missing </w:body>")
    return document_xml[:body_close] + "<w:sectPr></w:sectPr>" + document_xml[body_close:]


def _raw_ref(kind: str, ref_type: str, rid: str) -> str:
    return f'<w:{kind}Reference w:type="{ref_type}" r:id="{rid}"/>'


def _rewire_document_sectpr(target_extract_dir: Path, registry: Dict[str, Any], entries: List[Tuple[str, Dict[str, Any]]], part_to_rid: Dict[str, str], log: List[str]) -> None:
    doc_path = target_extract_dir / "word" / "document.xml"
    if not doc_path.exists():
        return

    page_layout = registry.get("page_layout", {}) if isinstance(registry, dict) else {}
    rid_to_part = _build_arch_rid_to_part(entries)

    doc_original = _ensure_body_sectpr(read_xml_text(doc_path))
    sectprs = extract_all_sectpr_blocks(doc_original)
    if not sectprs:
        raise ValueError("Target document has no usable sectPr after normalization")

    section_sources = choose_section_sources(len(sectprs), page_layout, require_default=True, log=log)
    updated_xml = doc_original
    order_index = canonical_sectpr_order_index()

    for idx, (target_sectpr, source) in enumerate(zip(sectprs, section_sources)):
        if re.fullmatch(r"<w:sectPr\b[^>]*/>", target_sectpr, flags=re.S):
            target_sectpr = re.sub(r"/\s*>$", ">", target_sectpr) + "</w:sectPr>"
        headers, footers = _extract_arch_hf_refs(source)

        open_tag_m = re.match(r'(<w:sectPr\b[^>]*>)', target_sectpr)
        close_tag = "</w:sectPr>"
        if not open_tag_m or not target_sectpr.endswith(close_tag):
            continue
        open_tag = open_tag_m.group(1)
        inner = target_sectpr[len(open_tag):-len(close_tag)]

        for tag in ("headerReference", "footerReference", "titlePg"):
            inner = strip_tag_block(inner, tag)
        children = extract_sectpr_children(inner)

        insert_nodes: List[str] = []
        source_sectpr_xml = source.get("sectPr", "") if isinstance(source, dict) else ""
        source_has_titlepg = "<w:titlePg" in source_sectpr_xml
        for ref_type, old_rid in headers.items():
            part_name = rid_to_part.get(old_rid)
            new_rid = part_to_rid.get(part_name) if part_name else None
            if not new_rid:
                continue
            insert_nodes.append(_raw_ref("header", ref_type, new_rid))

        for ref_type, old_rid in footers.items():
            part_name = rid_to_part.get(old_rid)
            new_rid = part_to_rid.get(part_name) if part_name else None
            if not new_rid:
                continue
            insert_nodes.append(_raw_ref("footer", ref_type, new_rid))

        if source_has_titlepg:
            insert_nodes.append("<w:titlePg/>")

        for node in insert_nodes:
            node_tag = child_tag_name(node)
            node_order = order_index.get(node_tag or "", 10_000)
            insert_at = len(children)
            for i, child in enumerate(children):
                ctag = child_tag_name(child)
                if order_index.get(ctag or "", 10_000) > node_order:
                    insert_at = i
                    break
            children.insert(insert_at, node)

        updated_sectpr = f"{open_tag}{''.join(children)}{close_tag}"
        updated_xml = replace_nth_sectpr_block(updated_xml, idx, updated_sectpr)

    updated_xml = prepare_xml_text_for_utf8(updated_xml)
    ET.fromstring(updated_xml.encode("utf-8"))
    doc_path.write_text(updated_xml, encoding="utf-8")
    log.append(f"Rewired sectPr header/footer references in {len(sectprs)} sections")


def _ensure_content_types(
    target_extract_dir: Path,
    part_to_type: Dict[str, str],
    result: HeaderFooterImportResult,
    log: List[str],
) -> None:
    ct_path = target_extract_dir / "[Content_Types].xml"
    if not ct_path.exists():
        return

    root = ET.fromstring(ct_path.read_bytes())
    imported_parts = {name.casefold() for name in part_to_type}
    for node in list(root.findall(f"{{{CT_NS}}}Override")):
        part_name = node.attrib.get("PartName", "")
        normalized_part = part_name.lstrip("/")
        content_type = node.attrib.get("ContentType", "")
        is_hf_type = content_type.endswith((".header+xml", ".footer+xml"))
        if normalized_part.casefold() in imported_parts or (
            is_hf_type and not (target_extract_dir / normalized_part).is_file()
        ):
            root.remove(node)
    existing_overrides: Dict[str, ET.Element] = {}
    for node in root.findall(f"{{{CT_NS}}}Override"):
        part_uri = node.attrib.get("PartName", "")
        if part_uri:
            existing_overrides.setdefault(part_uri.casefold(), node)

    # OPC permits at most one Default per extension.  Keep an extension-keyed
    # map so imported media can never create a second Default with a different
    # MIME type; conflicts are represented by per-part Overrides instead.
    existing_defaults: Dict[str, str] = {}
    for node in root.findall(f"{{{CT_NS}}}Default"):
        extension = node.attrib.get("Extension", "").casefold()
        if extension:
            existing_defaults.setdefault(extension, node.attrib.get("ContentType", ""))

    for part_name, kind in sorted(part_to_type.items()):
        part_uri = f"/{part_name}"
        override_key = part_uri.casefold()
        if override_key in existing_overrides:
            continue
        content_type = (
            "application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"
            if kind == "header"
            else "application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"
        )
        node = ET.SubElement(
            root,
            f"{{{CT_NS}}}Override",
            {"PartName": part_uri, "ContentType": content_type},
        )
        existing_overrides[override_key] = node

    for media_part_name, content_type in sorted(result.media_content_types.items()):
        extension = Path(media_part_name).suffix.lower().lstrip(".") or "bin"
        default_content_type = existing_defaults.get(extension)
        if default_content_type is None:
            ET.SubElement(
                root,
                f"{{{CT_NS}}}Default",
                {"Extension": extension, "ContentType": content_type},
            )
            existing_defaults[extension] = content_type
            continue
        if default_content_type.casefold() == content_type.casefold():
            continue

        part_uri = f"/{media_part_name}"
        override_key = part_uri.casefold()
        existing = existing_overrides.get(override_key)
        if existing is None:
            existing = ET.SubElement(
                root,
                f"{{{CT_NS}}}Override",
                {"PartName": part_uri, "ContentType": content_type},
            )
            existing_overrides[override_key] = existing
        else:
            existing.set("ContentType", content_type)

    ct_path.write_bytes(serialize_content_types(root))
    log.append("Updated [Content_Types].xml for header/footer parts and media")


def import_headers_footers(target_extract_dir: Path, registry: Dict[str, Any], log: List[str]) -> HeaderFooterImportResult:
    result = HeaderFooterImportResult()
    entries = _iter_hf_entries(registry)
    if not entries:
        log.append("No architect headers/footers in registry; skipping import")
        return result

    entries = _select_reachable_entries(target_extract_dir, registry, entries, log)
    if not entries:
        log.append("No architect headers/footers are reachable from mapped sections; preserving target parts")
        return result

    removed_parts, removed_rels = _remove_existing_hf_files(target_extract_dir, log)
    result.removed_part_names.update(removed_parts)
    result.removed_rels_names.update(removed_rels)
    write_entries = _unique_part_entries(entries)
    part_to_type = _write_hf_parts(target_extract_dir, write_entries, log, result)
    part_to_rid = _rebuild_document_rels(target_extract_dir, part_to_type, log)
    _rewire_document_sectpr(target_extract_dir, registry, entries, part_to_rid, log)
    _ensure_content_types(target_extract_dir, part_to_type, result, log)
    return result


def _extract_numeric_from_section_id(value: str) -> str:
    m = re.search(r"SECTION\s+([\d\s]+)", value or "", flags=re.IGNORECASE)
    if m:
        return re.sub(r"\s+", " ", m.group(1)).strip()
    digits = re.findall(r"\d+", value or "")
    return " ".join(digits).strip()


def patch_header_footer_tokens(
    target_extract_dir: Path,
    source_tokens: Dict[str, str],
    target_tokens: Dict[str, str],
    log: List[str],
    part_names: List[str] | None = None,
) -> None:
    from .core.token_utils import apply_case_pattern, detect_case_pattern, smart_title_case

    word_dir = target_extract_dir / "word"
    if not word_dir.exists():
        return

    arch_title = source_tokens.get("SectionTitle", "")
    target_title_display = target_tokens.get("SectionTitle_display", "") or target_tokens.get("SectionTitle", "")
    target_title_raw = target_tokens.get("SectionTitle", "")
    arch_id_numeric = source_tokens.get("SectionID_numeric") or _extract_numeric_from_section_id(source_tokens.get("SectionID", ""))
    target_id_numeric = target_tokens.get("SectionID_numeric") or _extract_numeric_from_section_id(target_tokens.get("SectionID", ""))

    wt_pattern = re.compile(r"(<w:t\b[^>]*>)([\s\S]*?)(</w:t>)")

    def _replace_first_visible_text(paragraph_xml: str, old_text: str, new_text: str) -> tuple[str, bool]:
        nodes = list(wt_pattern.finditer(paragraph_xml))
        if not nodes:
            return paragraph_xml, False

        visible = "".join(html.unescape(n.group(2)) for n in nodes)
        start = visible.find(old_text)
        if start < 0:
            return paragraph_xml, False
        end = start + len(old_text)

        out: List[str] = []
        cursor = 0
        vis_offset = 0
        replacement_placed = False
        changed = False
        for node in nodes:
            out.append(paragraph_xml[cursor:node.start()])
            open_tag, escaped_text, close_tag = node.groups()
            node_text = html.unescape(escaped_text)
            node_start = vis_offset
            node_end = vis_offset + len(node_text)
            new_node_text = node_text

            if node_end > start and node_start < end:
                overlap_start = max(start, node_start) - node_start
                overlap_end = min(end, node_end) - node_start
                before = node_text[:overlap_start]
                after = node_text[overlap_end:]
                if not replacement_placed:
                    new_node_text = before + new_text + after
                    replacement_placed = True
                else:
                    new_node_text = before + after
                changed = True

            out.append(f"{open_tag}{html.escape(new_node_text)}{close_tag}")
            cursor = node.end()
            vis_offset = node_end

        out.append(paragraph_xml[cursor:])
        return "".join(out), changed

    def _replace_host_visible_text(
        paragraph_xml: str,
        old_text: str,
        new_text: str,
    ) -> tuple[str, bool]:
        changed = False

        def _edit(protected_xml: str) -> str:
            nonlocal changed
            updated, changed = _replace_first_visible_text(
                protected_xml,
                old_text,
                new_text,
            )
            return updated

        return edit_preserving_out_of_scope_subtrees(paragraph_xml, _edit), changed

    if part_names is None:
        part_paths = [
            *sorted(word_dir.glob("header*.xml")),
            *sorted(word_dir.glob("footer*.xml")),
        ]
    else:
        part_paths = []
        for part_name in sorted(set(part_names)):
            if not is_allowed_header_footer_part_name(part_name):
                raise ValueError(f"Unsafe header/footer part name: {part_name!r}")
            part_path = target_extract_dir / part_name
            if part_path.is_file():
                part_paths.append(part_path)
    for part_path in part_paths:
        part_xml = read_xml_text(part_path)
        modified = False
        paragraph_matches = list(iter_paragraph_xml_blocks(part_xml))
        updated_chunks: List[str] = []
        cursor = 0
        for start, end, paragraph_xml in paragraph_matches:
            updated_chunks.append(part_xml[cursor:start])

            analysis_xml = strip_out_of_scope_subtrees(paragraph_xml)
            wt_nodes = list(wt_pattern.finditer(analysis_xml))
            visible_norm = "".join(html.unescape(n.group(2)) for n in wt_nodes)
            new_paragraph = paragraph_xml

            if visible_norm and arch_title and target_title_display:
                match_forms = [smart_title_case(arch_title), arch_title.upper(), arch_title]
                seen_forms = set()
                for form in match_forms:
                    if not form or form in seen_forms:
                        continue
                    seen_forms.add(form)
                    if form in visible_norm:
                        pattern = detect_case_pattern(form)
                        replacement = apply_case_pattern(target_title_raw or target_title_display, pattern)
                        new_paragraph, changed = _replace_host_visible_text(
                            new_paragraph,
                            form,
                            replacement,
                        )
                        modified = modified or changed
                        break

            if arch_id_numeric and target_id_numeric:
                for src_variant, dst_variant in (
                    (arch_id_numeric, target_id_numeric),
                    (arch_id_numeric.replace(" ", ""), target_id_numeric.replace(" ", "")),
                ):
                    if not src_variant:
                        continue
                    if src_variant in visible_norm:
                        new_paragraph, changed = _replace_host_visible_text(
                            new_paragraph,
                            src_variant,
                            dst_variant,
                        )
                        modified = modified or changed
                        break

            updated_chunks.append(new_paragraph)
            cursor = end

        updated_chunks.append(part_xml[cursor:])
        if paragraph_matches:
            part_xml = "".join(updated_chunks)

        if modified:
            part_path.write_text(prepare_xml_text_for_utf8(part_xml), encoding="utf-8")
            log.append(f"Patched tokens in {part_path.name}")
        else:
            log.append(f"No token matches found in {part_path.name}")


# Backward-compatible import name.  Behavior now covers both part types.
patch_footer_tokens = patch_header_footer_tokens


def remap_header_footer_numids(
    target_extract_dir: Path,
    part_names: List[str],
    num_id_remap: Dict[int, int],
    log: List[str],
) -> None:
    """Remap direct header/footer numIds to imported collision-safe IDs."""
    for part_name in sorted(set(part_names)):
        if not is_allowed_header_footer_part_name(part_name):
            raise ValueError(f"Unsafe header/footer part name: {part_name!r}")
        path = target_extract_dir / part_name
        if not path.is_file():
            raise FileNotFoundError(f"Imported header/footer part is missing: {part_name}")
        xml = read_xml_text(path)
        replacements = 0

        def _replace(match: re.Match[str]) -> str:
            nonlocal replacements
            old_id = int(match.group(2))
            new_id = num_id_remap.get(old_id)
            if new_id is None:
                return match.group(0)
            replacements += 1
            return f'{match.group(1)}{new_id}"'

        updated = re.sub(r'(<w:numId\b[^>]*w:val=")(\d+)"', _replace, xml)
        if replacements:
            updated = prepare_xml_text_for_utf8(updated)
            ET.fromstring(updated.encode("utf-8"))
            path.write_text(updated, encoding="utf-8")
            log.append(f"Remapped {replacements} direct numbering reference(s) in {part_name}")
