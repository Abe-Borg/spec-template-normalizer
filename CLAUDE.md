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
├── docx_decomposer.py          # Main orchestrator — extraction, slim bundle, style application, CLI
├── llm_classifier.py           # LLM automation — calls Anthropic API, chunking, coverage check
├── gui.py                      # Tkinter GUI wrapper (thin — no business logic)
├── arch_env_extractor.py       # Environment capture — produces arch_template_registry.json
├── phase1_smoke_test.py        # Validation test suite
├── master_prompt.txt           # System prompt for LLM CSI classification
├── run_instruction_prompt.txt  # Task prompt for LLM
├── instructions.json           # Example LLM output (style instructions)
├── schemas/
│   ├── arch_style_registry.v1.schema.json   # Formal JSON Schema for style registry
│   └── arch_template_registry.json          # Example/template for environment registry
├── requirements.txt            # Runtime dependencies (anthropic)
├── requirements-build.txt      # PyInstaller build dependencies
├── *.docx                      # Sample architect specification templates
├── *_extracted/                 # DOCX extraction working directories (generated)
├── README.md
└── .gitignore
```

## Technology Stack

- **Language:** Python 3.8+
- **External API:** Anthropic (Claude) — for semantic CSI structure classification
- **Key stdlib modules:** `zipfile`, `re`, `json`, `xml.etree.ElementTree`, `hashlib`, `argparse`, `pathlib`, `tkinter`
- **Runtime dependency:** `anthropic` (for API calls)

## Architecture and Data Flow

### Automated path (`--classify` flag — recommended)

```
DOCX (.docx file)
  │
  └─ [--classify] ──► extract ZIP
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
                        ├──► arch_style_registry.json
                        ├──► arch_template_registry.json
                        └──► coverage metric (% paragraphs classified)
```

### Manual path (fallback/debugging)

```
DOCX (.docx file)
  │
  ├─ [--normalize-slim] ──► extract ZIP ──► build_slim_bundle() ──► slim_bundle.json
  │                                                                       │
  │                                              (manual: feed to Claude LLM)
  │                                                                       │
  │                                                                       ▼
  │                                                              instructions.json
  │                                                                       │
  └─ [--apply-instructions] ──► validate_instructions()                   │
                                    │                                     │
                                    ├── derive styles from exemplar paragraphs
                                    ├── insert <w:pStyle> tags only
                                    ├── verify_stability() (hash checks)
                                    │
                                    ├──► arch_style_registry.json
                                    └──► arch_template_registry.json
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

### `docx_decomposer.py` — Main Orchestrator

| Function | Purpose |
|---|---|
| `extract_docx()` | Unzips .docx into workspace directory |
| `build_slim_bundle()` | Creates minimal JSON (text + numbering hints) for LLM input |
| `iter_paragraph_xml_blocks()` | Regex iterator over `<w:p>` blocks — preserves indices |
| `paragraph_text_from_block()` | Extracts visible text from paragraph XML |
| `validate_instructions()` | Strict validation of LLM output before application |
| `apply_instructions()` | Main apply logic: create styles, insert pStyle, verify stability |
| `apply_pstyle_to_paragraph_block()` | Surgically inserts `<w:pStyle>` into a single paragraph |
| `derive_style_def_from_paragraph()` | Extracts pPr/rPr from exemplar paragraph to build style definition |
| `build_style_xml_block()` | Generates `<w:style>` XML for insertion into `styles.xml` |
| `emit_arch_style_registry()` | Writes the final `arch_style_registry.json` contract |
| `snapshot_stability()` / `verify_stability()` | Hash-based invariant enforcement |

### `llm_classifier.py` — LLM Automation

Pure module (no CLI) — called by `docx_decomposer.py` and `gui.py`.

| Function | Purpose |
|---|---|
| `classify_document()` | Main entry: calls Anthropic API with slim bundle, returns instructions dict |
| `compute_coverage()` | Computes % of classifiable paragraphs that received a style |
| `estimate_tokens()` | Rough token count for chunking decisions |

**Design constraints:** Under 200 lines. No CLI of its own. Retry logic (up to 2 retries) for transient API failures. Chunking activates automatically when token estimate > 80K.

### `gui.py` — Tkinter GUI

Thin wrapper over the pipeline functions — no business logic.

| Class | Purpose |
|---|---|
| `App` | Main window: file picker, API key field, Run button, log area, status |
| `PipelineThread` | Background thread that runs the full pipeline |
| `LogRedirector` | Thread-safe stdout redirector for log display |

### `arch_env_extractor.py` — Environment Capture

| Function | Purpose |
|---|---|
| `extract_arch_template_registry()` | Main orchestrator — builds complete registry |
| `extract_doc_defaults()` | Extracts `<w:docDefaults>` (baseline rPr/pPr) |
| `extract_style_defs()` | All style definitions with raw XML blocks |
| `extract_theme()` | Theme fonts and colors from `theme1.xml` |
| `extract_settings()` | Compatibility flags from `settings.xml` |
| `extract_page_layout()` | Section properties, margins, columns |
| `extract_headers_footers()` | Complete header/footer XML |
| `extract_numbering()` | Numbering definitions from `numbering.xml` |
| `extract_fonts()` | Font table declarations |

### `phase1_smoke_test.py` — Validation

Runs both `--normalize-slim` and `--apply-instructions` in sequence, then validates `arch_style_registry.json` against the schema and checks all required CSI roles are present. `SectionID` is optional.

## Commands

### Automated Workflow (recommended)
```bash
# One command does everything: extract → classify → apply → emit registries
python docx_decomposer.py TEMPLATE.docx --classify
```

### GUI
```bash
python gui.py
```

### Manual Workflow (fallback/debugging)
```bash
# Step 1: Extract and prepare slim bundle for LLM
python docx_decomposer.py TEMPLATE.docx --normalize-slim

# Step 2: (Manual) Send master_prompt.txt + run_instruction_prompt.txt + slim_bundle.json to Claude
#         Save LLM JSON output as instructions.json

# Step 3: Apply instructions and generate both registries
python docx_decomposer.py TEMPLATE.docx --apply-instructions instructions.json
```

### Standalone Environment Extraction
```bash
python arch_env_extractor.py TEMPLATE.docx
python arch_env_extractor.py --extract-dir TEMPLATE_extracted
```

### Smoke Test
```bash
python phase1_smoke_test.py TEMPLATE.docx instructions.json
```

### CLI Flags (`docx_decomposer.py`)
- `--classify` — Full automated pipeline (extract → LLM classify → apply → emit registries)
- `--normalize-slim` — Generate `slim_bundle.json` for manual LLM analysis
- `--apply-instructions <json>` — Apply LLM instructions, produce both registries
- `--api-key <key>` — Anthropic API key (default: `ANTHROPIC_API_KEY` env var)
- `--model <id>` — Model ID for classification (default: `claude-sonnet-4-20250514`)
- `--extract-dir <dir>` — Custom extraction directory
- `--use-extract-dir <dir>` — Reuse existing extracted folder
- `--registry-out <path>` — Copy `arch_style_registry.json` to a specific location
- `--skip-env-extract` — Skip `arch_template_registry.json` generation
- `--master-prompt <file>` — Custom master prompt (default: `master_prompt.txt`)
- `--run-instruction <file>` — Custom run instruction (default: `run_instruction_prompt.txt`)

## CSI Role Hierarchy and Allowed Style IDs

The pipeline recognizes these CSI structural roles (from schema):

| Role | Style ID | Required? |
|---|---|---|
| `SectionID` | `CSI_SectionID__ARCH` | Optional |
| `SectionTitle` | `CSI_SectionTitle__ARCH` or `CSI_SectionName__ARCH` | Required |
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
  "version": 1,
  "source_docx": "TEMPLATE.docx",
  "roles": {
    "PART": { "style_id": "CSI_Part__ARCH", "exemplar_paragraph_index": 4, "style_name": "..." },
    ...
  }
}
```
Validated against `schemas/arch_style_registry.v1.schema.json`.

### `arch_template_registry.json`
Complete formatting environment with sections: `meta`, `package_inventory`, `doc_defaults`, `styles`, `theme`, `settings`, `page_layout`, `headers_footers`, `numbering`, `fonts`, `custom_xml`, `capture_policy`.

### Coverage Metric
After classification, the pipeline reports what percentage of non-empty, non-sectPr, non-editor-note paragraphs received a style. Coverage below 90% triggers a warning.

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
- No formal test framework (unittest/pytest) — uses `phase1_smoke_test.py` with subprocess calls
- Stability verification is built into the apply pipeline itself
- Test creates timestamped extraction directories to avoid collisions

## Common Pitfalls When Modifying This Code

1. **Do not switch `document.xml` parsing to ElementTree** — it will reformat XML and break paragraph index mapping.

2. **Do not add formatting fields to the LLM instruction schema** — the LLM must never specify pPr/rPr. Only `derive_from_paragraph_index` is allowed.

3. **Do not modify paragraphs containing `<w:sectPr>`** — these are section break containers and styling them can corrupt the document.

4. **Do not remove stability checks** — they are the primary safety mechanism ensuring the template isn't corrupted.

5. **`requirements.txt` is for runtime dependencies** (`anthropic`). Build/packaging dependencies are in `requirements-build.txt`.

6. **The `.docx` files and `*_extracted/` directories in the repo are test data** — they are architect specification templates used for development and testing.

7. **`llm_classifier.py` must remain a pure module** — no CLI of its own. It is called by `docx_decomposer.py` (via `--classify`) and by `gui.py`.

8. **`gui.py` must remain a thin wrapper** — no pipeline logic. It imports and calls the same functions as the CLI.

## Environment Setup

```bash
pip install -r requirements.txt
export ANTHROPIC_API_KEY='your-key-here'
```

Runtime: Python 3.8+ on Windows or Linux.

## License

Copyright 2025 Andrew Gossman. All Rights Reserved. Proprietary software — no license to use, copy, modify, or distribute without written permission.
