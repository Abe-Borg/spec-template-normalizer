# DOCX CSI Normalizer (Phase 1)

A tool that adds CSI structural tagging to Word specification documents without changing how they look, and captures the complete formatting environment for downstream reuse.

## Project Status

**Phase 1 Complete** - This project handles architect template normalization and environment capture. It produces formal contract artifacts (`arch_style_registry.json` and `arch_template_registry.json`) for use in downstream specification automation pipelines.

Phase 2 (applying architect formatting to MEP specs) is implemented separately.

## What it does

Takes an architect's Word spec template and produces two formal artifacts:

### 1. Structure Normalization (`arch_style_registry.json`)
- Identifies CSI structural elements (PART, Article, Paragraph, etc.) using LLM classification
- Creates Word paragraph styles based on actual formatting from exemplar paragraphs
- Applies those styles to matching paragraphs
- Guarantees zero visual change to the document
- Produces a semantic contract mapping CSI roles to Word styleIds

### 2. Environment Capture (`arch_template_registry.json`)
- Captures the complete formatting environment (the "rendering VM snapshot"):
  - Document defaults (docDefaults)
  - All style definitions with raw XML formatting blocks
  - Theme (fonts, colors)
  - Settings + compatibility flags
  - Page layout (sectPr, margins, columns)
  - Headers and footers
  - Numbering definitions
  - Font table

The output document looks identical but now has proper paragraph styles you can work with programmatically. The two JSON registries provide everything needed to recreate the architect's formatting environment in other documents.

## Architecture

### Two-Phase Design

**Phase 1 (this project):** Template normalization
- Input: Architect's Word template
- Output: Normalized template + two registries
- Purpose: Establish semantic structure and capture formatting DNA

**Phase 2 (separate):** Format application
- Input: MEP spec + arch registries
- Output: MEP spec with architect's formatting applied
- Purpose: Apply architect's formatting to consultant specs

### Key Innovation: LLM-Driven Classification

Instead of brittle pattern-matching rules:
1. Extract minimal "slim bundle" (text + numbering hints, no formatting)
2. Classify via Anthropic API (automated through GUI)
3. LLM classifies CSI structure and selects exemplar paragraphs
4. Script derives formatting locally from exemplars
5. Applies styles with surgical XML insertion

**Critical:** LLM never specifies formatting directly - only classification and exemplar selection.

## Installation
```bash
pip install -r requirements.txt
```

Set your Anthropic API key:
```bash
export ANTHROPIC_API_KEY='your-key-here'
```

For PyInstaller packaging, use the separate build requirements:
```bash
pip install -r requirements-build.txt
```

## Usage

### GUI (recommended)
```bash
python gui.py
```

The GUI is the primary entry point and runs the full automated pipeline:
- Select a `.docx` file
- Enter your API key (pre-populated from `ANTHROPIC_API_KEY` env var if set)
- Optionally choose an output folder (defaults to same directory as the input .docx)
- Click "Run Phase 1"
- View real-time progress and coverage metric
- Open the output folder or view the registry when done

The pipeline performs: extract → build slim bundle → classify via Anthropic API (default model: `claude-opus-4-6`) → apply styles → emit both registries.

After completion, you'll have:
- `ARCH_TEMPLATE_extracted/arch_style_registry.json` - CSI role mappings
- `ARCH_TEMPLATE_extracted/arch_template_registry.json` - Complete environment
- Modified `ARCH_TEMPLATE_extracted/` folder with styles applied

### Standalone environment extraction
```bash
# Extract environment registry from a .docx file
python arch_env_extractor.py ARCH_TEMPLATE.docx

# From already-extracted folder
python arch_env_extractor.py --extract-dir ARCH_TEMPLATE_extracted

# Custom output path
python arch_env_extractor.py ARCH_TEMPLATE.docx --output /path/to/output.json
```

## What gets created

### arch_style_registry.json
Maps CSI roles to Word styleIds:
```json
{
  "version": 1,
  "source_docx": "ARCH_TEMPLATE.docx",
  "roles": {
    "PART": {
      "style_id": "CSI_Part__ARCH",
      "exemplar_paragraph_index": 4,
      "style_name": "CSI Part (Architect Template)"
    },
    "ARTICLE": { "style_id": "CSI_Article__ARCH", ... },
    ...
  }
}
```

### arch_template_registry.json
Complete formatting environment with raw XML blocks (see `schemas/arch_template_registry.example.json` for full structure).

### Coverage metric
After classification, the pipeline reports what percentage of content paragraphs were classified. Paragraphs that are empty, contain section breaks, say "END OF SECTION", or are editor notes in brackets are excluded from the count. Coverage must be 100% for classifiable paragraphs; otherwise Phase 1 fails validation.

### Paragraph styles in DOCX
- `CSI_SectionID__ARCH` (optional)
- `CSI_SectionTitle__ARCH`
- `CSI_Part__ARCH`
- `CSI_Article__ARCH`
- `CSI_Paragraph__ARCH`
- `CSI_Subparagraph__ARCH`
- `CSI_Subsubparagraph__ARCH`

Each style captures exact formatting from exemplar paragraphs.

## How it works

1. **Extract**: Unzips DOCX, records hashes of headers/footers/section properties
2. **Slim bundle**: Creates minimal JSON (text + numbering hints) for LLM
3. **LLM classify**: Claude analyzes structure and returns JSON with role assignments + exemplar selections
4. **Derive locally**: Script extracts formatting from chosen exemplar paragraphs
5. **Apply surgically**: Inserts `<w:pStyle>` tags into paragraphs by index
6. **Capture environment**: Extracts complete formatting environment into arch_template_registry.json
7. **Verify**: Fails if anything changed except `<w:pStyle>` additions

## What it doesn't do

- Change visual appearance
- Modify numbering definitions
- Touch headers, footers, or section breaks
- Normalize spacing, indents, or alignment
- Generate new content
- Apply architect formatting to other documents (that's Phase 2)

These are intentional safeguards. The architect's template is sacred.

## Safety features

- Hard fails if headers/footers change (hash verification)
- Hard fails if section properties change (hash verification)
- Hard fails if relationships change (hash verification)
- Hard fails if paragraph properties drift beyond `<w:pStyle>` insertion
- LLM forbidden from specifying formatting (only structure classification)
- Comprehensive validation of LLM output before application

## Testing

### Unit tests
```bash
python -m pytest tests/
```

Regression tests cover XML extraction (`test_arch_env_extractor.py`), template registry validation (`test_arch_template_registry_validation.py`), and contract validation (`test_phase1_validator.py`).

### Smoke test
```bash
python phase1_smoke_test.py ARCH_TEMPLATE.docx instructions.json
```

The smoke test runs the full pipeline end-to-end and delegates contract validation to `phase1_validator`. It checks:
- Both registries are created
- arch_style_registry.json matches schema
- All required CSI roles are present (SectionID is optional)
- XML fragments are well-formed
- Cross-registry consistency (style IDs exist in template registry)
- No stability invariants violated

## Schemas

Formal JSON schemas are provided in `schemas/`:
- `arch_style_registry.v1.schema.json` - Style registry contract
- `schemas/arch_template_registry.example.json` - Environment registry example file (not runtime output)

## Requirements

- Python 3.8+
- Anthropic API key (default model: `claude-opus-4-6`)
- Windows or Linux (tested on both)

## Troubleshooting

**"Paragraph drift detected"**
The script changed something it shouldn't have. This is a bug, not expected behavior.

**"derive_from_paragraph_index out of range"**
LLM referenced a paragraph that doesn't exist. Try re-running or check your custom prompt.

**"LLM formatting fields are forbidden"**
LLM tried to specify formatting directly instead of referencing an exemplar. This violates the contract.

**"roles['PART'] exemplar_paragraph_index must equal derive_from_paragraph_index"**
The LLM's role mapping doesn't match its create_styles entries. The exemplar paragraph used for a role must be the same one used to derive that role's style.

**"No API key provided"**
Set the `ANTHROPIC_API_KEY` environment variable or enter the key in the GUI's API Key field.

**Coverage below 100%**
The LLM may not have classified all content paragraphs. Try re-running the pipeline via the GUI.


## Copyright Notice

**Copyright 2025 Abraham Borg. All Rights Reserved.**

This software and associated documentation files (the "Software") are the proprietary property of Abraham Borg.

**Unauthorized copying, modification, distribution, or use of this Software, via any medium, is strictly prohibited without express written permission from the copyright holder.**

This Software is provided for review and reference purposes only. No license or right to use, copy, modify, or distribute this Software for any purpose, commercial or non-commercial, is granted.
