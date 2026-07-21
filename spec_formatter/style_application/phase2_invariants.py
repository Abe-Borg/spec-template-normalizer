import re
import hashlib
import posixpath
import zipfile
from collections import Counter
from pathlib import Path
from typing import TYPE_CHECKING, List, Dict, Any, Optional
from urllib.parse import unquote, urlsplit
import xml.etree.ElementTree as ET

from .core.ooxml_namespaces import CT_NS, PKG_REL_NS, R_NS, W_NS
from .core.ooxml_text import decode_xml_bytes, prepare_xml_text_for_utf8
from .core.section_mapping import choose_section_sources
from .core.sectpr_tools import extract_all_sectpr_blocks, iter_document_sectpr_blocks
from .core.style_import import WORD_BUILTIN_STYLE_IDS
from .core.classification import (
    _build_numbering_catalog,
    _effective_numbering_semantics,
)
from .core.xml_helpers import (
    iter_element_xml_blocks,
    iter_paragraph_xml_blocks,
    paragraph_text_from_block,
)

if TYPE_CHECKING:
    from .core.application_policy import ApplicationPolicy


_OOXML_TRASH_ITEM_RE = re.compile(r"\[trash\]/[0-9A-Fa-f]{4}\.dat\Z")


def _is_ooxml_trash_item(name: str) -> bool:
    """Return whether a ZIP member uses the OOXML trash-item naming scheme."""
    return bool(_OOXML_TRASH_ITEM_RE.fullmatch(name))


def _sha256(b: bytes) -> str:
    return hashlib.sha256(b).hexdigest()

def _read_docx_part(docx: Path, internal_path: str) -> bytes:
    with zipfile.ZipFile(docx, "r") as z:
        return z.read(internal_path)


def _read_optional_docx_part(docx: Path, internal_path: str) -> bytes | None:
    with zipfile.ZipFile(docx, "r") as z:
        try:
            return z.read(internal_path)
        except KeyError:
            return None


def _resolve_application_policy(
    policy: Optional["ApplicationPolicy"],
    conversion_mode: Optional[str],
) -> "ApplicationPolicy":
    from .core.application_policy import application_policy_for_mode

    if policy is None:
        return application_policy_for_mode(conversion_mode or "format_only")
    if conversion_mode is not None and policy.conversion_mode != conversion_mode:
        raise ValueError(
            "application policy conversion_mode does not match explicit "
            f"conversion_mode ({policy.conversion_mode!r} != {conversion_mode!r})"
        )
    return policy


def _element_semantic_signature(element: ET.Element) -> tuple:
    """Namespace-prefix-independent signature for an OOXML element tree."""

    text = element.text or ""
    if not text.strip():
        text = ""
    return (
        element.tag,
        tuple(sorted(element.attrib.items())),
        text,
        tuple(_element_semantic_signature(child) for child in element),
    )


def _numbering_definition_signatures(numbering_xml: str) -> Counter:
    if not numbering_xml.strip():
        return Counter()
    root = ET.fromstring(numbering_xml)
    return Counter(_element_semantic_signature(child) for child in root)


def _verify_format_only_body_invariants(
    src_docx: Path,
    before_document_xml: str,
    after_document_xml: str,
    new_docx: Path | None,
) -> None:
    """Fail closed if Format-only changes target content or numbering semantics."""

    source_blocks = [block for _s, _e, block in iter_paragraph_xml_blocks(before_document_xml)]
    output_blocks = [block for _s, _e, block in iter_paragraph_xml_blocks(after_document_xml)]
    if len(source_blocks) != len(output_blocks):
        raise RuntimeError(
            "FORMAT_ONLY INVARIANT FAIL: target body paragraph count changed "
            f"({len(source_blocks)} -> {len(output_blocks)})"
        )

    source_text = [paragraph_text_from_block(block) for block in source_blocks]
    output_text = [paragraph_text_from_block(block) for block in output_blocks]
    if source_text != output_text:
        changed = next(
            idx
            for idx, (before, after) in enumerate(zip(source_text, output_text))
            if before != after
        )
        raise RuntimeError(
            "FORMAT_ONLY INVARIANT FAIL: target body text changed at paragraph "
            f"index {changed}"
        )

    source_styles_bytes = _read_optional_docx_part(src_docx, "word/styles.xml")
    source_numbering_bytes = _read_optional_docx_part(src_docx, "word/numbering.xml")
    source_styles = (
        decode_xml_bytes(source_styles_bytes, part_name="word/styles.xml (source)")
        if source_styles_bytes is not None
        else ""
    )
    source_numbering = (
        decode_xml_bytes(source_numbering_bytes, part_name="word/numbering.xml (source)")
        if source_numbering_bytes is not None
        else ""
    )

    if new_docx is not None:
        output_styles_bytes = _read_optional_docx_part(new_docx, "word/styles.xml")
        output_numbering_bytes = _read_optional_docx_part(new_docx, "word/numbering.xml")
        output_styles = (
            decode_xml_bytes(output_styles_bytes, part_name="word/styles.xml (output)")
            if output_styles_bytes is not None
            else ""
        )
        output_numbering = (
            decode_xml_bytes(output_numbering_bytes, part_name="word/numbering.xml (output)")
            if output_numbering_bytes is not None
            else ""
        )
    else:
        # Callers that validate an in-memory document before packaging cannot
        # have changed package-level styles/numbering yet.
        output_styles = source_styles
        output_numbering = source_numbering

    source_catalog = _build_numbering_catalog(source_numbering)
    output_catalog = _build_numbering_catalog(output_numbering)
    source_semantics = [
        _effective_numbering_semantics(block, source_styles, source_catalog)
        for block in source_blocks
    ]
    output_semantics = [
        _effective_numbering_semantics(block, output_styles, output_catalog)
        for block in output_blocks
    ]
    if source_semantics != output_semantics:
        changed = next(
            idx
            for idx, (before, after) in enumerate(
                zip(source_semantics, output_semantics)
            )
            if before != after
        )
        raise RuntimeError(
            "FORMAT_ONLY INVARIANT FAIL: effective target numbering changed at "
            f"paragraph index {changed}"
        )

    # Numbering import may append definitions needed by architect headers and
    # footers, but every pre-existing target definition must remain exact.
    source_definitions = _numbering_definition_signatures(source_numbering)
    output_definitions = _numbering_definition_signatures(output_numbering)
    missing_or_changed = source_definitions - output_definitions
    if missing_or_changed:
        raise RuntimeError(
            "FORMAT_ONLY INVARIANT FAIL: existing target numbering definitions changed"
        )

    # These structures are outside paragraph styling scope and must remain
    # byte-stable within document.xml even though the global shell may reflow.
    for qualified_name in (
        "w:tbl",
        "w:drawing",
        "w:pict",
        "w:object",
        "w:txbxContent",
    ):
        before_blocks = [
            block
            for _start, _end, block in iter_element_xml_blocks(
                before_document_xml,
                qualified_name,
            )
        ]
        after_blocks = [
            block
            for _start, _end, block in iter_element_xml_blocks(
                after_document_xml,
                qualified_name,
            )
        ]
        if before_blocks != after_blocks:
            raise RuntimeError(
                "FORMAT_ONLY INVARIANT FAIL: out-of-scope document structure "
                f"changed ({qualified_name})"
            )

def _extract_all_sectpr_blocks(document_xml: str) -> List[str]:
    return extract_all_sectpr_blocks(document_xml)


def _normalize_sectpr_for_comparison(sectpr: str) -> str:
    """Strip managed layout tags so only non-layout section semantics are compared."""
    out = sectpr
    for tag in (
        "pgSz", "pgMar", "cols", "docGrid",
        "headerReference", "footerReference", "titlePg",
    ):
        out = re.sub(rf'<w:{tag}\b[^>]*/>', '', out)
        out = re.sub(rf'<w:{tag}\b[^>]*>[\s\S]*?</w:{tag}>', '', out, flags=re.S)
    # reduce whitespace noise introduced by stripping
    out = re.sub(r'>\s+<', '><', out)
    out = out.strip()
    # Layout application expands legal self-closing sectPr blocks when it needs
    # to add children.  Canonicalize both empty spellings for comparison.
    out = re.sub(
        r"<w:sectPr\b([^>]*)>\s*</w:sectPr>",
        lambda match: f"<w:sectPr{match.group(1).rstrip()}/>",
        out,
        count=1,
        flags=re.S,
    )
    out = re.sub(r"\s+/>", "/>", out)
    return out


def _sectpr_records(document_xml: str) -> List[tuple[str, bool]]:
    """Return sectPr blocks and whether each is outside any paragraph."""
    paragraph_ranges = [
        (start, end)
        for start, end, _block in iter_element_xml_blocks(document_xml, "w:p")
    ]
    records: List[tuple[str, bool]] = []
    for start, _end, block in iter_document_sectpr_blocks(document_xml):
        inside_paragraph = any(p_start <= start < p_end for p_start, p_end in paragraph_ranges)
        records.append((block, not inside_paragraph))
    return records


def _registry_hf_rid_map(
    arch_headers: List[Dict[str, Any]],
    arch_footers: List[Dict[str, Any]],
) -> Dict[str, tuple[str, str]]:
    out: Dict[str, tuple[str, str]] = {}
    for kind, entries in (("header", arch_headers), ("footer", arch_footers)):
        for entry in entries:
            if not isinstance(entry, dict) or not isinstance(entry.get("part_name"), str):
                continue
            rid = next(
                (
                    entry.get(key)
                    for key in ("rel_id", "rid", "rId", "relationship_id")
                    if isinstance(entry.get(key), str) and entry.get(key)
                ),
                None,
            )
            if rid:
                out[rid] = (kind, entry["part_name"])
    return out


def _source_hf_refs(source: Dict[str, Any]) -> Dict[tuple[str, str], str]:
    refs: Dict[tuple[str, str], str] = {}
    for kind, keys in (
        ("header", ("header_refs", "headers")),
        ("footer", ("footer_refs", "footers")),
    ):
        values = next(
            (source.get(key) for key in keys if isinstance(source.get(key), dict)),
            {},
        )
        for ref_type, rid in values.items():
            if isinstance(ref_type, str) and isinstance(rid, str):
                refs[(kind, ref_type)] = rid
    return refs


def _sectpr_hf_refs(sectpr: str) -> List[tuple[str, str, str]]:
    refs: List[tuple[str, str, str]] = []
    for match in re.finditer(r"<w:(header|footer)Reference\b([^>]*)/?>", sectpr, flags=re.S):
        attrs = match.group(2)
        type_match = re.search(r"\bw:type\s*=\s*(['\"])(.*?)\1", attrs, flags=re.S)
        rid_match = re.search(r"\br:id\s*=\s*(['\"])(.*?)\1", attrs, flags=re.S)
        if type_match and rid_match:
            refs.append((match.group(1), type_match.group(2), rid_match.group(2)))
    return refs


def _extract_hf_relationship_subset(rels_xml: str) -> List[str]:
    rels = re.findall(r'<Relationship\b[^>]*/>', rels_xml)
    subset = [
        rel for rel in rels
        if 'relationships/header' in rel or 'relationships/footer' in rel
    ]
    return sorted(subset)


def _relationship_owner_part(rels_name: str) -> Optional[str]:
    if rels_name == "_rels/.rels":
        return ""
    marker = "/_rels/"
    if marker not in rels_name or not rels_name.endswith(".rels"):
        return None
    prefix, rel_name = rels_name.rsplit(marker, 1)
    return posixpath.join(prefix, rel_name[:-5])


def _relationship_part_for_owner(owner_part: str) -> str:
    directory, basename = posixpath.split(owner_part)
    if not basename:
        return "_rels/.rels"
    rels_name = f"{basename}.rels"
    return posixpath.join(directory, "_rels", rels_name) if directory else posixpath.join("_rels", rels_name)


def _resolve_relationship_target(owner_part: str, target: str) -> Optional[str]:
    parsed = urlsplit(target)
    if parsed.scheme or parsed.netloc:
        return None
    decoded = unquote(parsed.path).replace("\\", "/")
    if decoded.startswith("/"):
        resolved = posixpath.normpath(decoded.lstrip("/"))
    else:
        resolved = posixpath.normpath(
            posixpath.join(posixpath.dirname(owner_part), decoded)
        )
    if resolved in ("", ".", "..") or resolved.startswith("../"):
        return None
    return resolved


def validate_docx_package(docx_path: Path) -> None:
    """Fail closed when an emitted DOCX has broken OPC or Word references."""
    errors: List[str] = []
    parsed_xml: Dict[str, ET.Element] = {}

    try:
        with zipfile.ZipFile(docx_path, "r") as zf:
            names = zf.namelist()
            name_set = set(names)
            duplicates = sorted(name for name, count in Counter(names).items() if count > 1)
            if duplicates:
                errors.append(f"duplicate ZIP members: {duplicates}")
            names_by_casefold: Dict[str, List[str]] = {}
            for name in names:
                names_by_casefold.setdefault(name.casefold(), []).append(name)
            casefold_duplicates = [
                values
                for values in names_by_casefold.values()
                if len(values) > 1
            ]
            if casefold_duplicates:
                errors.append(
                    "case-insensitive duplicate ZIP members: "
                    f"{casefold_duplicates}"
                )
            invalid_names = sorted(
                name for name in name_set
                if name.startswith("/")
                or "\\" in name
                or posixpath.normpath(name).startswith("../")
            )
            if invalid_names:
                errors.append(f"unsafe package part names: {invalid_names}")
            bad_crc = zf.testzip()
            if bad_crc:
                errors.append(f"CRC failure in package part: {bad_crc}")

            for required in (
                "[Content_Types].xml",
                "_rels/.rels",
                "word/document.xml",
                "word/_rels/document.xml.rels",
                "word/styles.xml",
            ):
                if required not in name_set:
                    errors.append(f"missing required package part: {required}")

            for name in sorted(name_set):
                if not (name.endswith(".xml") or name.endswith(".rels") or name == "[Content_Types].xml"):
                    continue
                try:
                    parsed_xml[name] = ET.fromstring(zf.read(name))
                except ET.ParseError as exc:
                    errors.append(f"{name}: XML parse error: {exc}")

            ct_root = parsed_xml.get("[Content_Types].xml")
            if ct_root is not None:
                defaults: Dict[str, str] = {}
                overrides: Dict[str, str] = {}
                for node in ct_root.findall(f"{{{CT_NS}}}Default"):
                    ext = node.attrib.get("Extension", "").lower()
                    content_type = node.attrib.get("ContentType", "")
                    if not ext or not content_type:
                        errors.append("[Content_Types].xml contains malformed Default")
                    elif ext in defaults:
                        errors.append(f"duplicate content-type Default for extension: {ext}")
                    else:
                        defaults[ext] = content_type
                for node in ct_root.findall(f"{{{CT_NS}}}Override"):
                    raw_part = node.attrib.get("PartName", "")
                    content_type = node.attrib.get("ContentType", "")
                    part = posixpath.normpath(unquote(raw_part).lstrip("/"))
                    override_key = part.casefold()
                    if not raw_part.startswith("/") or not content_type or part.startswith("../"):
                        errors.append(f"malformed content-type Override: {raw_part!r}")
                    elif override_key in overrides:
                        errors.append(f"duplicate content-type Override for part: {part}")
                    else:
                        overrides[override_key] = content_type
                        if part not in name_set:
                            errors.append(f"content-type Override targets missing part: {part}")

                for name in sorted(name_set):
                    # OOXML trash items are discarded physical ZIP items, not
                    # package parts. ISO/IEC 29500 requires the exact
                    # ``[trash]/HHHH.dat`` form to have no content type.
                    if (
                        name.endswith("/")
                        or name == "[Content_Types].xml"
                        or _is_ooxml_trash_item(name)
                    ):
                        continue
                    ext = name.rsplit(".", 1)[-1].lower() if "." in name else ""
                    if name.casefold() not in overrides and ext not in defaults:
                        errors.append(f"package part has no content type: {name}")

            for rels_name, root in sorted(parsed_xml.items()):
                if not rels_name.endswith(".rels"):
                    continue
                owner = _relationship_owner_part(rels_name)
                if owner is None:
                    errors.append(f"relationship part has invalid location: {rels_name}")
                    continue
                if owner and owner not in name_set:
                    errors.append(
                        f"relationship part {rels_name} has missing owner part {owner}"
                    )
                if root.tag != f"{{{PKG_REL_NS}}}Relationships":
                    errors.append(f"{rels_name}: invalid Relationships root")
                    continue
                seen_ids: set[str] = set()
                for rel in list(root):
                    if rel.tag != f"{{{PKG_REL_NS}}}Relationship":
                        errors.append(f"{rels_name}: unsupported relationship element")
                        continue
                    rid = rel.attrib.get("Id", "")
                    rel_type = rel.attrib.get("Type", "")
                    target = rel.attrib.get("Target", "")
                    target_mode = rel.attrib.get("TargetMode")
                    if not rid or not rel_type or not target:
                        errors.append(
                            f"{rels_name}: relationship missing Id, Type, or Target"
                        )
                        continue
                    if rid in seen_ids:
                        errors.append(f"{rels_name}: duplicate relationship Id {rid}")
                    seen_ids.add(rid)
                    if target_mode not in {None, "Internal", "External"}:
                        errors.append(
                            f"{rels_name}: relationship {rid} has invalid TargetMode "
                            f"{target_mode!r}"
                        )
                        continue
                    if target_mode == "External":
                        continue
                    resolved = _resolve_relationship_target(owner, target)
                    if not resolved:
                        errors.append(f"{rels_name}: unsafe relationship target {target!r}")
                    elif resolved not in name_set:
                        errors.append(
                            f"{rels_name}: relationship {rid} targets missing part {resolved}"
                        )

            # Every office-document relationship attribute must resolve in the
            # relationship part owned by the XML part that carries it.  This
            # catches dangling header images, hyperlinks, charts, and embedded
            # objects even when the relationship target validator itself sees
            # a well-formed (but unrelated) .rels file.
            relationship_attributes = {
                f"{{{R_NS}}}id",
                f"{{{R_NS}}}embed",
                f"{{{R_NS}}}link",
            }
            for part_name, root in sorted(parsed_xml.items()):
                if part_name.endswith(".rels") or part_name == "[Content_Types].xml":
                    continue
                referenced_ids: set[str] = set()
                for node in root.iter():
                    for attr_name, value in node.attrib.items():
                        if attr_name not in relationship_attributes:
                            continue
                        local_name = attr_name.rsplit("}", 1)[-1]
                        if not value.strip():
                            errors.append(
                                f"{part_name}: empty r:{local_name} relationship reference"
                            )
                        else:
                            referenced_ids.add(value)
                if not referenced_ids:
                    continue
                rels_name = _relationship_part_for_owner(part_name)
                rels_root = parsed_xml.get(rels_name)
                if rels_root is None:
                    errors.append(
                        f"{part_name}: relationship references exist but {rels_name} is missing or malformed"
                    )
                    continue
                declared_ids = {
                    rel.attrib.get("Id", "")
                    for rel in rels_root.findall(f"{{{PKG_REL_NS}}}Relationship")
                    if rel.attrib.get("Id")
                }
                for rid in sorted(referenced_ids - declared_ids):
                    errors.append(
                        f"{part_name}: unresolved relationship reference {rid!r} in {rels_name}"
                    )

            styles_root = parsed_xml.get("word/styles.xml")
            style_ids: set[str] = set()
            if styles_root is not None:
                raw_style_ids = [
                    node.attrib.get(f"{{{W_NS}}}styleId", "")
                    for node in styles_root.findall(f"{{{W_NS}}}style")
                ]
                style_ids = {sid for sid in raw_style_ids if sid}
                duplicate_styles = sorted(
                    sid for sid, count in Counter(raw_style_ids).items() if sid and count > 1
                )
                if duplicate_styles:
                    errors.append(f"duplicate style IDs: {duplicate_styles}")

            style_ref_tags = {"pStyle", "rStyle", "tblStyle", "basedOn", "link", "next"}
            for part_name, root in sorted(parsed_xml.items()):
                if not part_name.startswith("word/"):
                    continue
                for node in root.iter():
                    local = node.tag.rsplit("}", 1)[-1]
                    if local not in style_ref_tags:
                        continue
                    sid = node.attrib.get(f"{{{W_NS}}}val", "")
                    if sid and sid not in style_ids and sid not in WORD_BUILTIN_STYLE_IDS:
                        errors.append(f"{part_name}: unresolved style reference {sid!r}")

            numbering_root = parsed_xml.get("word/numbering.xml")
            abstract_ids: set[int] = set()
            num_ids: set[int] = set()
            if numbering_root is not None:
                raw_abstract_ids: List[int] = []
                raw_num_ids: List[int] = []
                for node in numbering_root.findall(f"{{{W_NS}}}abstractNum"):
                    raw = node.attrib.get(f"{{{W_NS}}}abstractNumId", "")
                    try:
                        raw_abstract_ids.append(int(raw))
                    except ValueError:
                        errors.append(f"numbering.xml has invalid abstractNumId {raw!r}")
                abstract_ids = set(raw_abstract_ids)
                for node in numbering_root.findall(f"{{{W_NS}}}num"):
                    raw = node.attrib.get(f"{{{W_NS}}}numId", "")
                    try:
                        num_id = int(raw)
                        raw_num_ids.append(num_id)
                    except ValueError:
                        errors.append(f"numbering.xml has invalid numId {raw!r}")
                        continue
                    ref = node.find(f"{{{W_NS}}}abstractNumId")
                    raw_ref = ref.attrib.get(f"{{{W_NS}}}val", "") if ref is not None else ""
                    try:
                        abstract_ref = int(raw_ref)
                    except ValueError:
                        errors.append(f"numbering.xml numId {num_id} has invalid abstractNumId {raw_ref!r}")
                    else:
                        if abstract_ref not in abstract_ids:
                            errors.append(
                                f"numbering.xml numId {num_id} references missing abstractNumId {abstract_ref}"
                            )
                num_ids = set(raw_num_ids)
                duplicate_abstract = sorted(
                    value for value, count in Counter(raw_abstract_ids).items() if count > 1
                )
                duplicate_nums = sorted(
                    value for value, count in Counter(raw_num_ids).items() if count > 1
                )
                if duplicate_abstract:
                    errors.append(f"duplicate abstractNum IDs: {duplicate_abstract}")
                if duplicate_nums:
                    errors.append(f"duplicate num IDs: {duplicate_nums}")

            for part_name, root in sorted(parsed_xml.items()):
                if not part_name.startswith("word/") or part_name == "word/numbering.xml":
                    continue
                for node in root.iter(f"{{{W_NS}}}numId"):
                    raw = node.attrib.get(f"{{{W_NS}}}val", "")
                    try:
                        num_id = int(raw)
                    except ValueError:
                        errors.append(f"{part_name}: invalid numId reference {raw!r}")
                        continue
                    if num_id != 0 and num_id not in num_ids:
                        errors.append(f"{part_name}: unresolved numId reference {num_id}")
    except (OSError, zipfile.BadZipFile) as exc:
        raise RuntimeError(f"FINAL PACKAGE VALIDATION FAIL: {exc}") from exc

    if errors:
        raise RuntimeError(
            "FINAL PACKAGE VALIDATION FAIL:\n" + "\n".join(f"  - {e}" for e in errors)
        )


def _normalize_rpr_for_comparison(rpr_block: str) -> str:
    """
    Normalize an rPr block for comparison by stripping font-related elements.
    
    We allow changes to:
    - <w:rFonts .../> (font family)
    - <w:sz .../> (font size)  
    - <w:szCs .../> (complex script font size)
    
    Everything else (bold, italic, color, etc.) must remain unchanged.
    """
    result = rpr_block
    # Strip rFonts
    result = re.sub(r'<w:rFonts\b[^>]*/>', '', result)
    result = re.sub(r'<w:rFonts\b[^>]*>[\s\S]*?</w:rFonts>', '', result, flags=re.S)
    # Strip sz
    result = re.sub(r'<w:sz\b[^>]*/>', '', result)
    # Strip szCs
    result = re.sub(r'<w:szCs\b[^>]*/>', '', result)
    return result


def _extract_and_normalize_rpr_blocks(document_xml: str) -> List[str]:
    """
    Extract all rPr blocks from document.xml and normalize them.
    This allows us to check that non-font formatting is preserved.
    """
    rpr_blocks = re.findall(r"<w:rPr\b[\s\S]*?</w:rPr>", document_xml)
    return [_normalize_rpr_for_comparison(b) for b in rpr_blocks]


def _verify_target_header_footer_preserved(src_docx: Path, new_docx: Path) -> None:
    before_rels = decode_xml_bytes(
        _read_docx_part(src_docx, "word/_rels/document.xml.rels"),
        part_name="word/_rels/document.xml.rels (source)",
    )
    after_rels = decode_xml_bytes(
        _read_docx_part(new_docx, "word/_rels/document.xml.rels"),
        part_name="word/_rels/document.xml.rels (output)",
    )
    if _extract_hf_relationship_subset(before_rels) != _extract_hf_relationship_subset(after_rels):
        raise RuntimeError("INVARIANT FAIL: relationship subset changed")

    def _targets(rels_xml: str) -> set[str]:
        root = ET.fromstring(prepare_xml_text_for_utf8(rels_xml).encode("utf-8"))
        targets: set[str] = set()
        for rel in root.findall(f"{{{PKG_REL_NS}}}Relationship"):
            rel_type = rel.attrib.get("Type", "")
            if not rel_type.endswith(("/header", "/footer")):
                continue
            if rel.attrib.get("TargetMode") == "External":
                continue
            target = _resolve_relationship_target(
                "word/document.xml",
                rel.attrib.get("Target", ""),
            )
            if target is None:
                raise RuntimeError(
                    "INVARIANT FAIL: unsafe target header/footer relationship"
                )
            targets.add(target)
        return targets

    before_parts = _targets(before_rels)
    after_parts = _targets(after_rels)
    if before_parts != after_parts:
        raise RuntimeError(
            "INVARIANT FAIL: target header/footer relationship targets changed "
            "without architect replacements"
        )

    with zipfile.ZipFile(src_docx, "r") as z_before, zipfile.ZipFile(new_docx, "r") as z_after:
        before_names = set(z_before.namelist())
        after_names = set(z_after.namelist())
        for part_name in sorted(before_parts):
            if part_name not in before_names or part_name not in after_names:
                raise RuntimeError(
                    f"INVARIANT FAIL: target header/footer part missing: {part_name}"
                )
            if z_before.read(part_name) != z_after.read(part_name):
                raise RuntimeError(
                    f"INVARIANT FAIL: target header/footer part changed: {part_name}"
                )
            rels_name = _relationship_part_for_owner(part_name)
            before_has_rels = rels_name in before_names
            after_has_rels = rels_name in after_names
            if before_has_rels != after_has_rels:
                raise RuntimeError(
                    f"INVARIANT FAIL: target header/footer relationships changed: {rels_name}"
                )
            if before_has_rels and z_before.read(rels_name) != z_after.read(rels_name):
                raise RuntimeError(
                    f"INVARIANT FAIL: target header/footer relationships changed: {rels_name}"
                )


def verify_phase2_invariants(
    src_docx: Path,
    new_document_xml: bytes,
    new_docx: Path | None = None,
    arch_template_registry: Dict[str, Any] | None = None,
    policy: Optional["ApplicationPolicy"] = None,
    conversion_mode: Optional[str] = None,
) -> None:
    """
    Verify Phase 2 invariants:
    1. sectPr non-layout semantics unchanged (managed layout tags may change)
    2. If architect header/footer data is present, output header/footer parts match architect set
       and sectPr references resolve to valid document rels IDs
    3. Run properties unchanged EXCEPT for font-related elements (rFonts, sz, szCs)
    
    The font exception allows us to strip hardcoded fonts from MasterSpec docs
    so that style-level fonts take effect.
    """
    # 1) sectPr non-layout semantics unchanged
    before_doc = decode_xml_bytes(
        _read_docx_part(src_docx, "word/document.xml"),
        part_name="word/document.xml (source)",
    )
    after_doc = decode_xml_bytes(new_document_xml, part_name="word/document.xml (output)")

    application_policy = _resolve_application_policy(policy, conversion_mode)
    if application_policy.preserve_target_numbering:
        _verify_format_only_body_invariants(
            src_docx,
            before_doc,
            after_doc,
            new_docx,
        )

    before_records = _sectpr_records(before_doc)
    after_records = _sectpr_records(after_doc)
    before_sectprs = [block for block, _is_body in before_records]
    after_sectprs = [block for block, _is_body in after_records]
    before_body_count = sum(1 for _block, is_body in before_records if is_body)
    after_body_indices = [i for i, (_block, is_body) in enumerate(after_records) if is_body]
    created_body_sectpr = (
        before_body_count == 0
        and len(after_body_indices) == 1
        and len(after_records) == len(before_records) + 1
    )
    if len(before_records) != len(after_records) and not created_body_sectpr:
        raise RuntimeError("INVARIANT FAIL: sectPr block count changed")

    before_norm = [_normalize_sectpr_for_comparison(s) for s in before_sectprs]
    after_norm = [_normalize_sectpr_for_comparison(s) for s in after_sectprs]
    if created_body_sectpr:
        created_index = after_body_indices[0]
        if after_norm[created_index] != "<w:sectPr/>":
            raise RuntimeError(
                "INVARIANT FAIL: created body sectPr contains unmanaged section semantics"
            )
        existing_after_norm = after_norm[:created_index] + after_norm[created_index + 1:]
        existing_after_contexts = [
            is_body
            for i, (_block, is_body) in enumerate(after_records)
            if i != created_index
        ]
        if before_norm != existing_after_norm or [is_body for _block, is_body in before_records] != existing_after_contexts:
            raise RuntimeError("INVARIANT FAIL: non-layout sectPr semantics changed")
    elif (
        before_norm != after_norm
        or [is_body for _block, is_body in before_records]
        != [is_body for _block, is_body in after_records]
    ):
        raise RuntimeError("INVARIANT FAIL: non-layout sectPr semantics changed")

    # 2) header/footer invariants (against architect registry, when provided)
    hf_data = (arch_template_registry or {}).get("headers_footers", {}) if isinstance(arch_template_registry, dict) else {}
    raw_arch_headers = hf_data.get("headers", []) if isinstance(hf_data, dict) else []
    raw_arch_footers = hf_data.get("footers", []) if isinstance(hf_data, dict) else []
    arch_headers = raw_arch_headers if isinstance(raw_arch_headers, list) else []
    arch_footers = raw_arch_footers if isinstance(raw_arch_footers, list) else []

    if new_docx is not None and not (arch_headers or arch_footers):
        _verify_target_header_footer_preserved(src_docx, new_docx)

    if new_docx is not None and (arch_headers or arch_footers):
        if not after_sectprs:
            raise RuntimeError(
                "INVARIANT FAIL: architect header/footer data exists but output has no sectPr"
            )

        rid_to_arch_part = _registry_hf_rid_map(arch_headers, arch_footers)
        page_layout = (
            arch_template_registry.get("page_layout", {})
            if isinstance(arch_template_registry, dict)
            else {}
        )
        try:
            section_sources = choose_section_sources(
                len(after_sectprs),
                page_layout,
                require_default=True,
                log=[],
            )
        except ValueError as exc:
            raise RuntimeError(f"INVARIANT FAIL: {exc}") from exc

        expected_by_section: List[Dict[tuple[str, str], str]] = []
        expected_parts: set[str] = set()
        for source in section_sources:
            expected: Dict[tuple[str, str], str] = {}
            for key, arch_rid in _source_hf_refs(source).items():
                mapped = rid_to_arch_part.get(arch_rid)
                if mapped is None:
                    raise RuntimeError(
                        "INVARIANT FAIL: mapped architect section references unknown "
                        f"header/footer relationship {arch_rid!r}"
                    )
                mapped_kind, part_name = mapped
                if mapped_kind != key[0]:
                    raise RuntimeError(
                        f"INVARIANT FAIL: architect relationship {arch_rid!r} has wrong kind"
                    )
                expected[key] = part_name
                expected_parts.add(part_name)
            expected_by_section.append(expected)

        if not expected_parts:
            # Phase 1 may capture orphan architect parts.  With no mapped
            # section references, the importer deliberately performs no
            # replacement and the target package must remain intact.
            _verify_target_header_footer_preserved(src_docx, new_docx)

        if expected_parts:
            with zipfile.ZipFile(new_docx, "r") as z_after:
                rels_xml = z_after.read("word/_rels/document.xml.rels")
                rels_root = ET.fromstring(rels_xml)
                relationships = {
                    rel.attrib.get("Id"): rel
                    for rel in rels_root.findall('.//{*}Relationship')
                    if rel.attrib.get("Id")
                }
                actual_parts = set()
                for rel in relationships.values():
                    rel_type = rel.attrib.get("Type", "")
                    if not rel_type.endswith(("/header", "/footer")):
                        continue
                    target = _resolve_relationship_target(
                        "word/document.xml",
                        rel.attrib.get("Target", ""),
                    )
                    if target:
                        actual_parts.add(target)
                if expected_parts != actual_parts:
                    raise RuntimeError(
                        "INVARIANT FAIL: output header/footer relationship targets do not "
                        "match architect registry\n"
                        f"Expected: {sorted(expected_parts)}\nActual: {sorted(actual_parts)}"
                    )

                out_doc_xml = decode_xml_bytes(
                    z_after.read("word/document.xml"),
                    part_name="word/document.xml (packaged output)",
                )
                output_sectprs = _extract_all_sectpr_blocks(out_doc_xml)
                if len(output_sectprs) != len(expected_by_section):
                    raise RuntimeError("INVARIANT FAIL: output section mapping count changed")

                all_refs = [ref for sectpr in output_sectprs for ref in _sectpr_hf_refs(sectpr)]
                unresolved = sorted({rid for _kind, _ref_type, rid in all_refs if rid not in relationships})
                if unresolved:
                    raise RuntimeError(
                        "INVARIANT FAIL: document.xml contains header/footer refs with missing rel IDs: "
                        + ", ".join(unresolved)
                    )

                for index, (sectpr, expected) in enumerate(zip(output_sectprs, expected_by_section)):
                    refs = _sectpr_hf_refs(sectpr)
                    actual_keys = [(kind, ref_type) for kind, ref_type, _rid in refs]
                    if len(actual_keys) != len(set(actual_keys)):
                        raise RuntimeError(
                            f"INVARIANT FAIL: duplicate header/footer reference type in section {index}"
                        )
                    if set(actual_keys) != set(expected):
                        raise RuntimeError(
                            f"INVARIANT FAIL: section {index} header/footer references do not match "
                            f"architect mapping; expected={sorted(expected)}, actual={sorted(actual_keys)}"
                        )

                    for kind, ref_type, rid in refs:
                        rel = relationships[rid]
                        rel_type = rel.attrib.get("Type", "")
                        if not rel_type.endswith(f"/{kind}"):
                            raise RuntimeError(
                                f"INVARIANT FAIL: section {index} {kind}/{ref_type} relationship "
                                f"{rid} has wrong type"
                            )
                        target = _resolve_relationship_target(
                            "word/document.xml",
                            rel.attrib.get("Target", ""),
                        )
                        if target != expected[(kind, ref_type)]:
                            raise RuntimeError(
                                f"INVARIANT FAIL: section {index} {kind}/{ref_type} resolves to "
                                f"{target!r}, expected {expected[(kind, ref_type)]!r}"
                            )

    # 3) no run-level formatting edits EXCEPT font-related (rFonts, sz, szCs)
    # We normalize rPr blocks by stripping font elements, then compare
    before_rpr_normalized = _extract_and_normalize_rpr_blocks(before_doc)
    after_rpr_normalized = _extract_and_normalize_rpr_blocks(after_doc)
    
    # Note: The number of rPr blocks might change if we remove empty ones,
    # so we compare the non-empty normalized blocks
    before_rpr_filtered = [b for b in before_rpr_normalized if b.strip() and b.strip() != '<w:rPr></w:rPr>']
    after_rpr_filtered = [b for b in after_rpr_normalized if b.strip() and b.strip() != '<w:rPr></w:rPr>']
    
    # Instead of strict equality (which fails if rPr blocks are removed),
    # we check that no NON-FONT formatting was changed.
    # This is a relaxed check - we're mainly guarding against accidental changes.
    
    # Verify that no non-font formatting was lost.
    # We can't do a strict count comparison because stripping fonts can remove
    # entire rPr blocks (when rFonts/sz/szCs were the only children).
    # Instead, check that every non-empty normalized "before" block still appears
    # somewhere in the "after" set. This catches accidental bold/italic/color changes.
    before_set = {}
    for b in before_rpr_filtered:
        before_set[b] = before_set.get(b, 0) + 1

    after_set = {}
    for a in after_rpr_filtered:
        after_set[a] = after_set.get(a, 0) + 1

    for block, count in before_set.items():
        after_count = after_set.get(block, 0)
        if after_count < count:
            raise RuntimeError(
                f"INVARIANT FAIL: non-font run formatting was lost. "
                f"A normalized rPr block appeared {count}x before but {after_count}x after.\n"
                f"Block: {block[:200]}"
            )
