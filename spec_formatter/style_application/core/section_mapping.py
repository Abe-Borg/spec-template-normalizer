from __future__ import annotations

from typing import Any, Dict, List


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

    if target_count == len(chain) and target_count > 0:
        return chain

    if default_section is not None:
        # A section-chain entry is positional only when both documents have the
        # same section count.  On a mismatch, reusing the leading architect
        # entries can apply a cover/title-page section to an ordinary target
        # section.  The registry's declared default is the only safe source.
        mapped: List[Dict[str, Any]] = [default_section for _ in range(target_count)]
        if target_count != len(chain):
            log.append(
                f"target sections={target_count}, architect sections={len(chain)}; "
                "using page_layout.default_section for every target section"
            )
        return mapped

    if require_default and target_count != len(chain):
        raise ValueError(
            "Template registry missing usable page_layout.default_section for section-count mismatch"
        )

    return chain[:target_count]

