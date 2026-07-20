"""
Stability snapshot and verification for DOCX processing.

Ensures headers, footers, sectPr blocks, and document.xml.rels
remain unchanged during Phase 2 processing.
"""

import hashlib
import xml.etree.ElementTree as ET
from pathlib import Path
from dataclasses import dataclass
from typing import Dict

from .ooxml_text import read_xml_text
from .ooxml_namespaces import PKG_REL_NS
from .opc_paths import (
    relationship_part_name_for_owner,
    resolve_internal_relationship_target,
)
from .sectpr_tools import extract_all_sectpr_blocks


def sha256_bytes(b: bytes) -> str:
    return hashlib.sha256(b).hexdigest()


def sha256_text(s: str) -> str:
    return sha256_bytes(s.encode("utf-8"))


@dataclass
class StabilitySnapshot:
    header_footer_hashes: Dict[str, str]
    sectpr_hash: str
    doc_rels_hash: str


def snapshot_headers_footers(extract_dir: Path) -> Dict[str, str]:
    """Hash relationship-reachable header/footer owners and their rels."""
    hashes: Dict[str, str] = {}
    document_rels = extract_dir / "word" / "_rels" / "document.xml.rels"
    if not document_rels.is_file():
        return hashes
    root = ET.fromstring(document_rels.read_bytes())
    for rel in root.findall(f"{{{PKG_REL_NS}}}Relationship"):
        rel_type = rel.attrib.get("Type", "")
        if not rel_type.endswith(("/header", "/footer")):
            continue
        if rel.attrib.get("TargetMode") == "External":
            continue
        part_name = resolve_internal_relationship_target(
            "word/document.xml",
            rel.attrib.get("Target", ""),
        )
        part_path = extract_dir / part_name
        if not part_path.is_file():
            raise ValueError(f"Referenced header/footer part is missing: {part_name}")
        hashes[part_name] = sha256_bytes(part_path.read_bytes())
        rels_name = relationship_part_name_for_owner(part_name)
        rels_path = extract_dir / rels_name
        if rels_path.is_file():
            hashes[rels_name] = sha256_bytes(rels_path.read_bytes())
    return hashes


def snapshot_doc_rels_hash(extract_dir: Path) -> str:
    rels_path = extract_dir / "word" / "_rels" / "document.xml.rels"
    if not rels_path.exists():
        return ""
    return sha256_bytes(rels_path.read_bytes())


def extract_sectpr_block(document_xml: str) -> str:
    """Return exact in-scope paired and self-closing section-property blocks."""
    return "\n".join(extract_all_sectpr_blocks(document_xml))


def snapshot_stability(extract_dir: Path) -> StabilitySnapshot:
    doc_path = extract_dir / "word" / "document.xml"
    doc_text = read_xml_text(doc_path)
    sectpr = extract_sectpr_block(doc_text)
    return StabilitySnapshot(
        header_footer_hashes=snapshot_headers_footers(extract_dir),
        sectpr_hash=sha256_text(sectpr),
        doc_rels_hash=snapshot_doc_rels_hash(extract_dir),
    )


def verify_stability(extract_dir: Path, snap: StabilitySnapshot) -> None:
    current_hf = snapshot_headers_footers(extract_dir)
    if current_hf != snap.header_footer_hashes:
        changed = []
        all_keys = set(current_hf.keys()) | set(snap.header_footer_hashes.keys())
        for k in sorted(all_keys):
            if current_hf.get(k) != snap.header_footer_hashes.get(k):
                changed.append(k)
        raise ValueError(f"Header/footer stability check FAILED. Changed: {changed}")

    doc_text = read_xml_text(extract_dir / "word" / "document.xml")
    current_sectpr = extract_sectpr_block(doc_text)
    if sha256_text(current_sectpr) != snap.sectpr_hash:
        raise ValueError("Section properties (w:sectPr) stability check FAILED.")

    # relationships must be stable too (header/footer binding lives here)
    current_rels = snapshot_doc_rels_hash(extract_dir)
    if current_rels != snap.doc_rels_hash:
        raise ValueError("document.xml.rels stability check FAILED (can break header/footer).")
