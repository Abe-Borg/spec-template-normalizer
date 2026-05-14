# CLAUDE.md — AI Assistant Guide for DOCX CSI Normalizer

## Project Overview

This is **Phase 1** of a two-phase DOCX specification automation pipeline. It takes an architect's Word specification template (.docx) and produces two formal contract artifacts:

1. **`arch_style_registry.json`** — Maps CSI (Construction Specifications Institute) structural roles to Word paragraph styles
2. **`arch_template_registry.json`** — Captures the complete formatting environment ("rendering VM snapshot")

Phase 2 (separate codebase) uses these artifacts to apply architect formatting to MEP (Mechanical/Electrical/Plumbing) consultant specs.

**The architect's template is sacred.** The output document must be pixel-identical to the input — only `<w:pStyle>` tags are inserted.

## Repository Structure

```
.
├── docx_decomposer.py          # Library module — extraction, slim bundle, style application
├── llm_classifier.py           # LLM automation — calls Anthropic API, chunking, repair chain
├── paragraph_rules.py          # CSI role detection, paragraph classification, skip rules
├── gui.py                      # Tkinter GUI wrapper (thin — no business logic)
├── arch_env_extractor.py       # Environment capture — produces arch_template_registry.json (library)
├── phase1_validator.py         # Contract validation — validates both registries before writing
├── phase1_smoke_test.py        # Validation test suite
├── master_prompt.txt           # System prompt for LLM CSI classification
├── run_instruction_prompt.txt  # Task prompt for LLM
├── instructions.json           # Example LLM output (style instructions)
├── schemas/
│   ├── arch_style_registry.v1.schema.json      # Formal JSON Schema for style registry (v1)
│   ├── arch_style_registry.v2.schema.json      # Formal JSON Schema for style registry (v2, current)
│   ├── phase1_instructions.schema.json         # Formal JSON Schema for LLM instruction output
│   └── arch_template_registry.example.json     # Example/reference for environment registry
├── tests/
│   ├── test_arch_env_extractor.py              # Regression tests for XML extraction
│   ├── test_arch_template_registry_validation.py  # Template registry validation tests
│   ├── test_phase1_validator.py                # Contract validation tests
│   ├── test_phase1_hardening.py                # Post-classification repair chain tests
│   ├── test_paragraph_rules.py                 # Paragraph rule detection tests
│   ├── test_phase1_contracts.py                # End-to-end contract shape tests
│   ├── test_apply_reserved_styles.py           # Reserved style application tests
│   ├── test_document_structure.py              # Document structure tests
│   ├── test_docx_decomposer_rpr_hints.py       # rPr hint extraction tests
│   ├── test_registry_numbering.py              # Numbering registry tests
│   ├── test_semantic_validation.py             # Semantic validation tests
│   └── test_style_catalog.py                   # Style catalog tests
├── requirements.txt            # Runtime dependencies (anthropic)
├── requirements-build.txt      # PyInstaller build dependencies
├── README.md
└── .gitignore
```

## Technology Stack

- **Language:** Python 3.8+
- **External API:** Anthropic (Claude) — for semantic CSI structure classification
- **Key stdlib modules:** `zipfile`, `re`, `json`, `xml.etree.ElementTree`, `hashlib`, `pathlib`, `tkinter`
- **Runtime dependency:** `anthropic` (for API calls)

## Architecture and Data Flow

### GUI Pipeline (gui.py)

```
DOCX (.docx file)
  │
  └─ [Run Phase 1] ──► extract ZIP
                          │
                          ├── build_slim_bundle() ──► slim_bundle.json
                          │                                │
                          │                   classify_document() (Anthropic API)
                          │                                │
                          │                                ▼
                          │                       instructions.json (saved for audit)
                          │                                │
                          ├── validate_instructions()      │
                          ├── apply_instructions()  ◄──────┘
                          │     ├── derive styles from exemplar paragraphs
                          │     ├── insert <w:pStyle> tags only
                          │     └── verify_stability() (hash checks)
                          │
                          ├──► arch_style_registry.json  ──► (copied to output folder)
                          ├──► arch_template_registry.json ──► (copied to output folder)
                          └──► coverage metric (% paragraphs classified)
```

## Critical Design Invariants

**These are hard rules. Violating them will break the pipeline or corrupt documents.**

1. **Never full-XML-parse `document.xml`** — Uses regex (`iter_paragraph_xml_blocks()`) to preserve paragraph indices and raw XML structure. ElementTree is only used for `styles.xml` name lookups and catalog building.

2. **Surgical XML insertion only** — The only modification to `document.xml` is inserting/replacing `<w:pStyle>` elements. Nothing else may change.

3. **Exemplar-based formatting** — New CSI styles are derived from actual paragraphs in the template (`derive_from_paragraph_index`). The LLM is forbidden from specifying any formatting (pPr, rPr, fonts, spacing, alignment, etc.).

4. **Stability snapshots** — `StabilitySnapshot` (dataclass) records SHA-256 hashes of headers, footers, section properties, and document.xml.rels before any modifications. `verify_stability()` enforces these haven't changed after processing.

5. **No sectPr paragraphs** — Paragraphs containing `<w:sectPr>` are never styled and never used as exemplars.

6. **No DOCX reconstruction** — Phase 1 intentionally does NOT produce a .docx output file. It works on the extracted folder only.

## Key Source Files

### `docx_decomposer.py` — Library Module

| Function | Purpose |
|---|---|
| `extract_docx()` | Unzips .docx into workspace directory (with OneDrive lock retry) |
| `build_slim_bundle()` | Creates minimal JSON (text + numbering hints) for LLM input |
| `build_style_catalog()` | Builds style name/type catalog from `styles.xml` |
| `build_numbering_catalog()` | Builds numbering definition catalog from `numbering.xml` |
| `iter_paragraph_xml_blocks()` | Regex iterator over `<w:p>` blocks — preserves indices |
| `paragraph_text_from_block()` | Extracts visible text from paragraph XML |
| `paragraph_contains_sectpr()` | Checks if paragraph contains `<w:sectPr>` |
| `paragraph_pstyle_from_block()` | Extracts existing `<w:pStyle>` value from paragraph |
| `paragraph_numpr_from_block()` | Extracts numbering properties (numId, ilvl) |
| `validate_instructions()` | Strict validation of LLM output before application |
| `apply_instructions()` | Main apply logic: create styles, insert pStyle, verify stability |
| `apply_pstyle_to_paragraph_block()` | Surgically inserts `<w:pStyle>` into a single paragraph |
| `derive_style_def_from_paragraph()` | Extracts pPr/rPr from exemplar paragraph to build style definition |
| `build_style_xml_block()` | Generates `<w:style>` XML for insertion into `styles.xml` |
| `insert_styles_into_styles_xml()` | Inserts style blocks into `styles.xml` |
| `emit_arch_style_registry()` | Writes the final `arch_style_registry.json` contract |
| `snapshot_stability()` / `verify_stability()` | Hash-based invariant enforcement |

### `llm_classifier.py` — LLM Automation

Pure module (no CLI) — called by `gui.py`.

| Function | Purpose |
|---|---|
| `classify_document()` | Main entry: calls Anthropic API with slim bundle, returns instructions dict. Default model: `claude-opus-4-6` |
| `compute_coverage()` | Computes % of classifiable paragraphs that received a style |
| `estimate_tokens()` | Rough token count for chunking decisions |
| `_repair_missing_roles()` | Auto-adds roles where strong text signals prove they exist but LLM omitted them |
| `_repair_role_exemplar_mismatches()` | Moves exemplar to role-consistent paragraph when exemplar text contradicts declared role |
| `_repair_strong_signal_mismatches()` | Overwrites apply_pStyle entries where regex signals contradict the assigned styleId |
| `_repair_coverage_gaps()` | Nearest-neighbor fallback fill for paragraphs still unclassified after patch retries |
| `_build_patch_prompt()` | Builds a targeted prompt for coverage patch API calls |

**Design constraints:** No CLI of its own. Retry logic (up to 2 retries) for transient API failures. Chunking activates automatically when token estimate > 80K or paragraphs > 300. After initial classification, a deterministic repair chain runs before any API patch retries.

### `paragraph_rules.py` — CSI Role Detection and Skip Rules

Imported by `docx_decomposer.py` and `llm_classifier.py`. Contains all regex patterns and classification logic for CSI structural roles.

| Function | Purpose |
|---|---|
| `compute_skip_reason()` | Returns skip label (empty, sectPr, in_table, end_of_section, editor_note, copyright_notice, specifier_note) or None |
| `is_classifiable_paragraph()` | Returns True if a paragraph should receive a CSI style |
| `is_editor_note()` | Detects bracketed editor instructions |
| `is_copyright_notice()` | Detects copyright/distribution boilerplate |
| `is_specifier_note()` | Detects specifier editing instructions (Retain/Revise/etc.) |
| `detect_role_signal()` | Returns the unambiguous CSI role for a paragraph's text based on regex patterns |
| `infer_expected_roles()` | Returns the set of roles expected in a document based on strong text signals |

### `gui.py` — Tkinter GUI (Primary Entry Point)

Thin wrapper over the pipeline functions — no business logic.

| Class | Purpose |
|---|---|
| `App` | Main window: template picker, API key field, output folder picker, Run button, log area, status |
| `PipelineThread` | Background thread that runs the full pipeline |
| `LogRedirector` | Thread-safe stdout redirector for log display |

### `arch_env_extractor.py` — Environment Capture (library module)

Imported as a library by gui.py and phase1_smoke_test.py. No CLI entry point.

| Function | Purpose |
|---|---|
| `extract_arch_template_registry()` | Main orchestrator — builds complete registry |
| `extract_doc_defaults()` | Extracts `<w:docDefaults>` (baseline rPr/pPr) |
| `extract_style_defs()` | All style definitions with raw XML blocks |
| `extract_latent_styles()` | Extracts latent style exception settings |
| `extract_table_styles()` | Extracts table-specific style definitions |
| `extract_styles_section()` | Composite: doc_defaults + style_defs + latent + table styles |
| `extract_theme()` | Theme fonts and colors from `theme1.xml` |
| `extract_settings()` | Compatibility flags from `settings.xml` |
| `extract_page_layout()` | Section properties, margins, columns |
| `extract_headers_footers()` | Complete header/footer XML |
| `extract_numbering()` | Numbering definitions from `numbering.xml` |
| `extract_fonts()` | Font table declarations |
| `extract_relationships()` | Document relationship entries |
| `extract_package_inventory()` | Inventories which OOXML parts are present |

### `phase1_validator.py` — Contract Validation

Validates both registries before they are written to disk. Imported by `phase1_smoke_test.py`.

| Function | Purpose |
|---|---|
| `validate_template_registry()` | Validates `arch_template_registry.json` shape and XML fragment well-formedness |
| `validate_style_registry()` | Validates `arch_style_registry.json` shape and required CSI roles |
| `validate_cross_registry()` | Cross-checks that style IDs in the style registry exist in the template registry |
| `validate_phase1_contracts()` | Runs all three validations above in sequence |

### `phase1_smoke_test.py` — Validation

Calls `extract_docx()`, `build_slim_bundle()`, `apply_instructions()`, `build_style_registry_dict()`, and `extract_arch_template_registry()` directly. Delegates contract validation to `phase1_validator.validate_phase1_contracts()`, which checks required CSI roles, template registry structure, XML fragment well-formedness, and cross-registry consistency (style IDs referenced by the style registry must exist in the template registry). `SectionID` is optional. The test fails if either registry is missing, malformed, or contains truncated XML fragments.

## Commands

### GUI (primary entry point)
```bash
python gui.py
```

The GUI provides:
- **Template** picker — select the architect .docx specification
- **API Key** field — Anthropic API key (pre-populated from `ANTHROPIC_API_KEY` env var)
- **Output Folder** picker — where `arch_style_registry.json` and `arch_template_registry.json` are written (defaults to same directory as the input .docx)
- **Run Phase 1** button — runs the full pipeline
- Post-completion buttons to open the output folder or view the style registry

### Unit Tests
```bash
python -m pytest tests/
```

### Smoke Test (developer tool)
```bash
python phase1_smoke_test.py TEMPLATE.docx instructions.json
```


## CSI Role Hierarchy and Allowed Style IDs

The pipeline recognizes these CSI structural roles (from schema):

| Role | Style ID | Required? |
|---|---|---|
| `SectionID` | `CSI_SectionID__ARCH` | Optional |
| `SectionTitle` | `CSI_SectionTitle__ARCH` or `CSI_SectionTitle__ARCH` | Required |
| `PART` | `CSI_Part__ARCH` | Required |
| `ARTICLE` | `CSI_Article__ARCH` | Required |
| `PARAGRAPH` | `CSI_Paragraph__ARCH` | Required |
| `SUBPARAGRAPH` | `CSI_Subparagraph__ARCH` | Required |
| `SUBSUBPARAGRAPH` | `CSI_Subsubparagraph__ARCH` | Required |

All created style IDs must match the pattern `CSI_*__ARCH`.

## Output Artifacts

### `arch_style_registry.json`
```json
{
  "version": 2,
  "source_docx": "TEMPLATE.docx",
  "source_tokens": {
    "SectionID": "SECTION 23 05 00",
    "SectionTitle": "COMMON WORK RESULTS FOR HVAC"
  },
  "roles": {
    "PART": { "style_id": "CSI_Part__ARCH", "exemplar_paragraph_index": 4, "style_name": "..." },
    ...
  }
}
```
Current schema: `schemas/arch_style_registry.v2.schema.json`. `source_tokens` captures the literal text of the exemplar paragraph for each role (SectionID and SectionTitle only), used by Phase 2 for section identification.

### `arch_template_registry.json`
Complete formatting environment with sections: `meta`, `package_inventory`, `doc_defaults`, `styles`, `theme`, `settings`, `page_layout`, `headers_footers`, `numbering`, `fonts`, `custom_xml`, `capture_policy`.

### `arch_styles_raw.xml` and `arch_settings_raw.xml`
Byte-exact copies of the architect template's `word/styles.xml` and `word/settings.xml`. Phase 2 can import these directly for exact style and settings fidelity. Written alongside the registries in the output directory.

### Coverage Metric
After classification, the pipeline reports what percentage of non-empty, non-sectPr, non-editor-note paragraphs received a style. Coverage must be 100% for classifiable paragraphs; otherwise Phase 1 fails validation.

## Development Conventions

### Code Style
- Python 3.8+ compatible (uses `from __future__ import annotations`)
- Type hints throughout (`Dict`, `List`, `Optional`, `Tuple`, `Set`, `Any` from `typing`)
- Frozen dataclasses for immutable state (`StabilitySnapshot`)
- Functions are well-documented with inline comments explaining "why"

### XML Handling
- **Regex-first for `document.xml`** — preserves byte-level structure and paragraph indices
- **ElementTree only for read-only lookups** on `styles.xml`, `numbering.xml`
- Raw XML blocks are stored as strings in JSON registries (not parsed/re-serialized)
- `_canonicalize()` strips rsids and proofing marks for cleaner output

### Error Handling
- Hard `ValueError` raises for all invariant violations
- No silent failures — every validation check is explicit
- Descriptive error messages with context (paragraph index, style ID, etc.)

### Testing
- `tests/` directory contains pytest-compatible regression tests for XML extraction and contract validation
- `phase1_smoke_test.py` provides end-to-end validation with direct function calls (requires a .docx and instructions.json)
- Stability verification is built into the apply pipeline itself
- Smoke test creates timestamped extraction directories to avoid collisions

## Common Pitfalls When Modifying This Code

1. **Do not switch `document.xml` parsing to ElementTree** — it will reformat XML and break paragraph index mapping.

2. **Do not add formatting fields to the LLM instruction schema** — the LLM must never specify pPr/rPr. Only `derive_from_paragraph_index` is allowed.

3. **Do not modify paragraphs containing `<w:sectPr>`** — these are section break containers and styling them can corrupt the document.

4. **Do not remove stability checks** — they are the primary safety mechanism ensuring the template isn't corrupted.

5. **`requirements.txt` is for runtime dependencies** (`anthropic`). Build/packaging dependencies are in `requirements-build.txt`.

6. **`.docx` files and `*_extracted/` directories are local test data** — they are generated during development and testing but are not committed to the repo.

7. **`llm_classifier.py` must remain a pure module** — no CLI of its own. It is imported only by `gui.py`.

8. **`gui.py` must remain a thin wrapper** — no pipeline logic. It imports and calls library functions from `docx_decomposer.py`.

9. **`docx_decomposer.py` is a library module only** — no CLI entry point. All user interaction goes through `gui.py`.

## Environment Setup

```bash
pip install -r requirements.txt
export ANTHROPIC_API_KEY='your-key-here'
```

Runtime: Python 3.8+ on Windows or Linux.

## License

Copyright 2025 Abraham Borg. All Rights Reserved. Proprietary software — no license to use, copy, modify, or distribute without written permission.
