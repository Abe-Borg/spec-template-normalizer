"""
Core package for Phase 2 MEP Specification Styling Engine.

Re-exports the public interface from submodules.
"""

from .xml_helpers import (
    iter_paragraph_xml_blocks,
    paragraph_text_from_block,
    paragraph_contains_sectpr,
    paragraph_pstyle_from_block,
    paragraph_numpr_from_block,
    paragraph_ppr_hints_from_block,
    apply_pstyle_to_paragraph_block,
    strip_run_font_formatting,
)

from .stability import (
    StabilitySnapshot,
    sha256_bytes,
    sha256_text,
    snapshot_headers_footers,
    snapshot_doc_rels_hash,
    extract_sectpr_block,
    snapshot_stability,
    verify_stability,
)

from .style_import import (
    materialize_arch_style_block,
    extract_style_block_raw,
    import_arch_styles_into_target,
    insert_styles_into_styles_xml,
)

from .classification import (
    PHASE2_MASTER_PROMPT,
    PHASE2_RUN_INSTRUCTION,
    BOILERPLATE_PATTERNS,
    strip_boilerplate_with_report,
    build_phase2_slim_bundle,
    apply_phase2_classifications,
    validate_phase2_classification_contract,
)

from .registry import (
    resolve_arch_extract_root,
    load_available_roles_from_registry,
    load_arch_style_registry,
    write_phase2_preflight,
    build_arch_styles_xml_from_registry,
    preflight_validate_registries,
)
