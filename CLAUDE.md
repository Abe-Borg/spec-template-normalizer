# CLAUDE.md - Engineering Guide for Spec Template Normalizer

## Purpose

This repository is Phase 1 of a two-step DOCX formatting pipeline. It analyzes an architect's specification template and publishes a validated `.phase1` bundle. A separate Phase 2 application validates that complete bundle and applies its style system to another specification.

The supported Phase 1 pipeline is observational with respect to the input document:

- Never modify, retag, repackage, or replace the selected DOCX.
- Never mutate the extracted working snapshot.
- Never publish a normalized DOCX or extracted package.
- Derive generated CSI styles into a separate `portable_styles.xml` document.
- Publish only a complete, validated, checksummed `.phase1` directory.

The source snapshot is the authority for all artifacts created in a run.

## Canonical architecture

`gui.py` is a thin client of `phase1_pipeline.run_phase1()`. The headless pipeline owns the business flow:

```text
source DOCX
  -> stable private snapshot
  -> contained, bounded package extraction
  -> slim paragraph/style/numbering bundle
  -> LLM role + styled/ignored classification
  -> instruction and 100% disposition validation
  -> portable stylesheet derivation (separate output; no retagging)
  -> bounded environment capture
  -> registry + audit validation
  -> private bundle staging
  -> source-part, cross-artifact, size, and checksum validation
  -> atomic rename to unique *.phase1 directory
```

If any step fails, no final bundle is published. Temporary work and unpublished staging data are cleaned up.

## Repository map

```text
phase1_pipeline.py       Canonical headless orchestration and immutable source snapshot
phase1_bundle.py         Bundle models, audit creation, staging, validation, atomic publication
docx_decomposer.py       Safe package extraction, slim bundle, semantic checks, portable style derivation
llm_classifier.py        Anthropic request, bounded retries, targeted coverage patches
paragraph_rules.py       CSI signals, numbering-aware role inference, classifiable-universe rules
arch_env_extractor.py    Bounded formatting-environment capture
phase1_validator.py      Instruction, style registry, template registry, and cross-registry validation
gui.py                   CustomTkinter client; no pipeline business logic
master_prompt.txt         Classifier system contract
run_instruction_prompt.txt
instructions.json        Example classifier instruction shape
schemas/
  phase1_bundle_manifest.v1.schema.json
  classification_audit.v1.schema.json
  phase1_instructions.schema.json
  arch_style_registry.v2.schema.json
  arch_style_registry.v1.schema.json       Legacy registry schema
  arch_template_registry.example.json      Reference shape, not a generated artifact
tests/                    Unit, adversarial, contract, bundle, and pipeline regression tests
```

`apply_instructions()` is a legacy developer surface and must not be wired into `gui.py`, `run_phase1()`, or Phase 2. `phase1_smoke_test.py` is the supported offline harness: it injects precomputed instructions into the same immutable `run_phase1()` -> `build_portable_styles_xml()` -> bundle path.

## Non-negotiable invariants

### 1. One immutable source identity

`run_phase1()` copies the selected DOCX to a private snapshot, checks size and modification time around the copy, and hashes the snapshot. Extraction, environment capture, registry metadata, audit metadata, and bundle staging all use that snapshot. Do not hash or reopen the user's live path later in the run as an artifact authority.

### 2. No source paragraph mutation

The production flow reads `word/document.xml`; it does not insert or replace `<w:pStyle>`. New role styles are built in memory and written to `portable_styles.xml`. `source_styles.xml` remains byte-identical to the source `word/styles.xml`.

### 3. Formatting stays local and source-derived

The LLM may select roles, explicit ignored paragraphs, and exemplar indices. It may not specify formatting XML or formatting attributes. `build_portable_styles_xml()` derives paragraph/run properties from the selected source exemplar and preserves source-style inheritance. A direct property seen in one marked run must not be promoted to an entire style unless it is common to the visible text-bearing runs.

### 4. Explicit classification universe

Structurally out-of-scope paragraphs are limited to empty paragraphs, table paragraphs, and empty structural section-break paragraphs. A visible paragraph remains in the classification universe when its paragraph properties contain `w:sectPr`; its section properties must be preserved exactly. Every other candidate index must occur exactly once in one of:

- `apply_pStyle`: styled CSI content
- `ignored_paragraphs`: non-CSI/editorial content with a non-empty reason

The sets must be disjoint and their union must equal the expected indices. Editor notes and copyright/specifier notices are auditable ignored content, not invisible skips. There is no nearest-neighbor style fallback.

### 5. Bundle is the API boundary

Phase 2 consumes the entire `.phase1` directory, starting with strict manifest validation. Two loose registries are not a valid Phase 1 handoff. The bundle validator checks required artifacts, exact paths, source identity, sizes, hashes, JSON/XML validity, audit consistency, registry relationships, and unlisted files.

### 6. Atomic publication

Artifacts are copied into a private sibling staging directory, revalidated there, and published by same-filesystem directory rename. New runs use unique names. Existing bundles are not overwritten by default. Never write generated artifacts directly into a final `.phase1` directory.

## Bundle contract

A successful run publishes:

| Artifact ID | File | Required | Kind |
|---|---|---:|---|
| `style_registry` | `arch_style_registry.json` | yes | generated |
| `template_registry` | `arch_template_registry.json` | yes | generated |
| `classification_audit` | `classification_audit.json` | yes | generated |
| `source_styles` | `source_styles.xml` | yes | exact source bytes |
| `portable_styles` | `portable_styles.xml` | yes | generated |
| `source_settings` | `source_settings.xml` | only when the source has `word/settings.xml` | exact source bytes |

`phase1_bundle_manifest.json` identifies format `spec-template-normalizer.phase1`, manifest version 1, bundle ID, UTC creation time, producer/run/classifier identity, prompt hashes, source filename/hash/size, required artifact IDs, and each artifact's path/media type/hash/size/source kind.

The normal directory name is:

```text
<safe-source-stem>--<source-sha12>--<run-token12>.phase1
```

The classification audit embeds the validated instruction object and records every slim paragraph with its text fingerprint, truncation flag, skip reason, and `styled`, `ignored`, or `out_of_scope` disposition. It also hashes the instruction and paragraph collections.

## CSI role contract

Allowed roles and reserved generated style IDs are:

| Role | Style ID |
|---|---|
| `SectionID` | `CSI_SectionID__ARCH` |
| `SectionTitle` | `CSI_SectionTitle__ARCH` |
| `PART` | `CSI_Part__ARCH` |
| `ARTICLE` | `CSI_Article__ARCH` |
| `PARAGRAPH` | `CSI_Paragraph__ARCH` |
| `SUBPARAGRAPH` | `CSI_Subparagraph__ARCH` |
| `SUBSUBPARAGRAPH` | `CSI_Subsubparagraph__ARCH` |
| `END_OF_SECTION` | `CSI_EndOfSection__ARCH` |

Role expectations come from text signals and effective Word numbering, including numbering inherited through paragraph styles. Do not treat arbitrary `A.`, `1.`, or `a.` text globally as proof of CSI hierarchy. Validate exemplars, role/style coherence, numbering family/level coverage, style inheritance, and style references against the source catalogs.

## Module responsibilities

### `phase1_pipeline.py`

- `run_phase1()` is the sole supported end-to-end entry point.
- `_snapshot_source()` establishes the stable input used for the whole run.
- Reads the two prompt files, invokes classification, requires complete coverage, writes generated artifacts to private work storage, and delegates publication to `phase1_bundle`.
- Returns `Phase1Result` only after atomic publication.

Keep filesystem/UI concerns outside of classifiers and decomposers. Keep GUI progress reporting behind the optional callback.

### `phase1_bundle.py`

- `build_classification_audit()` and `validate_classification_audit()` freeze and verify paragraph dispositions.
- `stage_phase1_bundle()` copies artifacts to private staging, verifies exact source parts and all contracts, writes the manifest, and validates the staged directory.
- `publish_staged_bundle()` performs the atomic rename.
- `validate_bundle_directory()` is the consumer-facing integrity gate and should also be used by Phase 2.

Do not weaken `reject_unlisted`, required artifact, filename, hash, or source identity checks for convenience.

### `docx_decomposer.py`

- `extract_docx()` performs safe, bounded OPC/ZIP extraction.
- `build_slim_bundle()` reads visible paragraph text, table/section status, direct and effective numbering, source style IDs, and catalogs.
- `validate_instructions()` performs document-aware shape, semantic, role, exemplar, style, and coverage checks.
- `build_portable_styles_xml()` derives and inserts generated styles into a separate stylesheet string without writing the extracted source package.
- `build_style_registry_dict()` emits role metadata, source identity, resolved formatting, and numbering provenance for Phase 2.

Paragraph indices are tied to the source `document.xml` paragraph sequence. Preserve the visible-text and index semantics when changing XML parsing.

### `llm_classifier.py`

- `classify_document()` performs one primary classifier request and bounded targeted patch requests.
- API retry is limited to transient connection, rate-limit, and server failures.
- Invalid credentials, invalid requests, malformed JSON, semantic contradictions, and unresolved coverage fail closed.
- Current single-pass input limit is an estimated 150,000 tokens.

Do not restore positional or nearest-neighbor gap filling. A missing paragraph must receive an explicit supported style or an explicit ignore reason.

### `arch_env_extractor.py`

Captures the supported formatting environment: package inventory, document defaults, styles, latent/table styles, theme, compatibility settings, section/page layout, header/footer XML and contained media, numbering, font table, and relationships.

This registry is deliberately bounded and normalized; it is not a complete rendering VM or a byte-exact mirror of every package part. `capture_policy` describes normalization. Exact source styles/settings are separate bundle artifacts.

### `phase1_validator.py`

Owns the formal instruction and registry contracts. Keep its allowed role/style sets synchronized with prompt files and JSON schemas. Cross-registry checks must use the portable stylesheet/environment that Phase 2 will actually consume.

### `gui.py`

Owns only input collection, background-thread execution, progress/log display, and final status. It calls `run_phase1()` and displays `Phase1Result`. It must not recreate extraction, classification, artifact copying, or validation logic.

## Untrusted input and limits

DOCX input and relationship metadata are untrusted.

- Reject absolute, traversal, duplicate, and symbolic-link package members.
- Limits: 10,000 package entries; 512 MiB total uncompressed; 128 MiB per part; compression ratio at most 1,000.
- Parse and validate relationship parts. Resolve internal targets only inside the package root.
- Never dereference external relationship targets, local paths, UNC paths, URLs, or encoded traversal.
- Reject malformed relationship XML and broken required relationship metadata.
- Header/footer media limits: 16 MiB per asset and 64 MiB total.
- Preserve content types from `[Content_Types].xml` where available.

Any new extractor must have adversarial tests for containment, external targets, malformed XML, symlinks/reparse behavior, and size bounds.

## Environment capture semantics

The template registry stores normalized source-derived XML fragments. Current capture policy does not canonicalize whitespace but does strip volatile rsid attributes and proofing markers. Do not call these fields “raw XML.”

Use these terms consistently:

- `source_styles.xml`: byte-exact original styles part
- `source_settings.xml`: byte-exact original settings part, when present
- `portable_styles.xml`: generated stylesheet for Phase 2
- `arch_template_registry.json`: bounded normalized environment capture

The retired names `arch_styles_raw.xml` and `arch_settings_raw.xml` are not bundle artifacts.

## Development commands

```bash
pip install -r requirements-dev.txt
python -m pytest tests
python gui.py
```

Headless usage is through Python:

```python
from pathlib import Path
from phase1_pipeline import run_phase1

result = run_phase1(
    source_docx=Path("template.docx"),
    output_root=Path("output"),
    api_key="...",
)
```

## Change checklist

Before considering a Phase 1 change complete:

1. Confirm the source path and extracted snapshot remain unchanged.
2. Confirm every classifiable paragraph is explicitly styled or ignored once.
3. Confirm generated style inheritance and numbering provenance match the exemplar.
4. Confirm exact source style/settings artifacts still match the source ZIP members.
5. Validate the staged bundle, manifest, hashes, source identity, audit, and registries.
6. Run focused tests plus the complete test suite.
7. Test a real architect template and the Phase 2 bundle consumer when the contract changes.
8. Update prompts, schemas, validators, README, and Phase 2 together for wire-contract changes.

## Common mistakes

- Calling `apply_instructions()` from the production pipeline.
- Treating the extraction directory as an output deliverable.
- Writing loose artifacts into the selected output root.
- Describing the two registries as the complete handoff.
- Calling normalized registry fragments “raw” or “complete VM state.”
- Omitting ignored paragraphs from coverage.
- Filling missing classifications from adjacent paragraphs.
- Dropping style-only numbering definitions from the slim catalog or registry.
- Trusting an LLM-provided `basedOn` instead of the exemplar's source style.
- Following an external relationship or a path that escapes the package.
- Publishing before all files have been copied and revalidated.

## Platform and license

Runtime is Python 3.9+ with Windows as the primary GUI platform. Core package processing is intended to remain portable.

Copyright 2025 Abraham Borg. All Rights Reserved. Proprietary software; no license is granted without written permission.
