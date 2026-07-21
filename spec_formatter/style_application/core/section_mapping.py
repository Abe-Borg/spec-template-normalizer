from __future__ import annotations

import json
from typing import Any, Dict, List


def _canonical_shell_signature(section: Dict[str, Any]) -> str:
    """Return only the architect-owned shell semantics for conflict checks."""

    managed = {
        key: section.get(key)
        for key in (
            "page_size",
            "page_margins",
            "columns",
            "doc_grid",
            "header_refs",
            "footer_refs",
        )
    }
    return json.dumps(managed, sort_keys=True, separators=(",", ":"))


def choose_section_sources(
    target_count: int,
    page_layout: Dict[str, Any],
    *,
    require_default: bool,
    log: List[str],
) -> List[Dict[str, Any]]:
    chain_raw = page_layout.get("section_chain", []) if isinstance(page_layout, dict) else []
    chain = [item for item in chain_raw if isinstance(item, dict)]
    default_raw = page_layout.get("default_section") if isinstance(page_layout, dict) else None
    default_section = default_raw if isinstance(default_raw, dict) else None

    if chain:
        signatures = {_canonical_shell_signature(section) for section in chain}
        if len(signatures) != 1:
            raise ValueError(
                "Architect template has conflicting section shells; use one canonical "
                "page layout and default/even/first header-footer mapping."
            )

    if default_section is not None:
        if chain and _canonical_shell_signature(default_section) not in {
            _canonical_shell_signature(section) for section in chain
        }:
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

    return [chain[0] for _ in range(target_count)] if chain else []

