from __future__ import annotations

import copy
import json
from typing import Any, Dict, List


_REFERENCE_FIELDS = ("header_refs", "footer_refs")
_REFERENCE_TYPES = ("default", "even", "first")
_LAYOUT_FIELDS = ("page_size", "page_margins", "columns", "doc_grid")


def resolve_effective_section_chain(
    page_layout: Dict[str, Any],
) -> List[Dict[str, Any]]:
    """Return section copies with Word header/footer inheritance resolved.

    A missing reference for one header/footer type inherits only that same type
    from the preceding section.  The registry deliberately records the raw
    ``sectPr`` references, so resolve them here without mutating a cached
    contract-v2 profile.
    """

    chain_raw = (
        page_layout.get("section_chain", [])
        if isinstance(page_layout, dict)
        else []
    )
    chain = [item for item in chain_raw if isinstance(item, dict)]
    inherited = {
        field: {reference_type: None for reference_type in _REFERENCE_TYPES}
        for field in _REFERENCE_FIELDS
    }
    resolved: List[Dict[str, Any]] = []

    for raw_section in chain:
        section = copy.deepcopy(raw_section)
        for field in _REFERENCE_FIELDS:
            raw_refs = raw_section.get(field)
            refs = copy.deepcopy(raw_refs) if isinstance(raw_refs, dict) else {}
            for reference_type in _REFERENCE_TYPES:
                explicit = (
                    raw_refs.get(reference_type)
                    if isinstance(raw_refs, dict)
                    else None
                )
                if explicit is not None:
                    inherited[field][reference_type] = copy.deepcopy(explicit)
                refs[reference_type] = copy.deepcopy(
                    inherited[field][reference_type]
                )
            section[field] = refs
        resolved.append(section)

    return resolved


def _resolved_default_section(
    default_section: Dict[str, Any],
    raw_chain: List[Dict[str, Any]],
    resolved_chain: List[Dict[str, Any]],
) -> Dict[str, Any]:
    """Map a raw default section to its effective chain entry when possible."""

    def apply_explicit_default_values(
        effective_section: Dict[str, Any],
    ) -> Dict[str, Any]:
        resolved_default = copy.deepcopy(effective_section)
        for field in _LAYOUT_FIELDS:
            if field in default_section:
                resolved_default[field] = copy.deepcopy(default_section[field])
        for field in _REFERENCE_FIELDS:
            default_refs = default_section.get(field)
            if not isinstance(default_refs, dict):
                continue
            resolved_refs = resolved_default.setdefault(field, {})
            for reference_type in _REFERENCE_TYPES:
                explicit = default_refs.get(reference_type)
                if explicit is not None:
                    resolved_refs[reference_type] = copy.deepcopy(explicit)
        return resolved_default

    section_index = default_section.get("section_index")
    if section_index is not None:
        matching_positions = [
            position
            for position, section in enumerate(raw_chain)
            if section.get("section_index") == section_index
        ]
        if len(matching_positions) > 1:
            raise ValueError(
                "Architect template section chain has duplicate section_index values."
            )
        if matching_positions:
            return apply_explicit_default_values(
                resolved_chain[matching_positions[0]]
            )

    # Older or hand-built registries may omit section_index.  The extractor's
    # default section is the last chain entry, so prefer the last structural
    # match when several raw entries are identical.
    for position in range(len(raw_chain) - 1, -1, -1):
        if raw_chain[position] == default_section:
            return apply_explicit_default_values(resolved_chain[position])

    return copy.deepcopy(default_section)


def _canonical_shell_signature(section: Dict[str, Any]) -> str:
    """Return only the architect-owned shell semantics for conflict checks."""

    managed = {
        key: section.get(key)
        for key in _LAYOUT_FIELDS
    }
    for field in _REFERENCE_FIELDS:
        refs = section.get(field)
        managed[field] = {
            reference_type: (
                refs.get(reference_type) if isinstance(refs, dict) else None
            )
            for reference_type in _REFERENCE_TYPES
        }
    return json.dumps(managed, sort_keys=True, separators=(",", ":"))


def choose_section_sources(
    target_count: int,
    page_layout: Dict[str, Any],
    *,
    require_default: bool,
    log: List[str],
) -> List[Dict[str, Any]]:
    chain_raw = (
        page_layout.get("section_chain", [])
        if isinstance(page_layout, dict)
        else []
    )
    chain = [item for item in chain_raw if isinstance(item, dict)]
    effective_chain = resolve_effective_section_chain(page_layout)
    default_raw = page_layout.get("default_section") if isinstance(page_layout, dict) else None
    default_section = (
        _resolved_default_section(default_raw, chain, effective_chain)
        if isinstance(default_raw, dict)
        else None
    )

    if effective_chain:
        signatures = {
            _canonical_shell_signature(section) for section in effective_chain
        }
        if len(signatures) != 1:
            raise ValueError(
                "Architect template has conflicting section shells; use one canonical "
                "page layout and default/even/first header-footer mapping."
            )

    if default_section is not None:
        if (
            effective_chain
            and _canonical_shell_signature(default_section) not in signatures
        ):
            raise ValueError(
                "Architect template default section conflicts with its section chain."
            )
        # The architect shell is canonical and applies to every target section;
        # target section-break placement and unmanaged semantics remain owned
        # by the target document.
        mapped: List[Dict[str, Any]] = [default_section for _ in range(target_count)]
        if target_count != len(chain):
            log.append(
                f"target sections={target_count}, architect sections={len(chain)}; "
                "using page_layout.default_section for every target section"
            )
        return mapped

    if require_default:
        raise ValueError(
            "Template registry missing usable page_layout.default_section"
        )

    return [effective_chain[0] for _ in range(target_count)] if effective_chain else []

