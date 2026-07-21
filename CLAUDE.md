# CLAUDE.md - Engineering Guide for Spec Template Normalizer

## Purpose

This repository is the unified Specification Formatter. It analyzes and caches
an architect template, classifies one or more target specifications, applies an
explicit formatting policy, and publishes validated DOCX files plus complete
run provenance. The `.phase1` bundle remains an internal integrity boundary;
it is not a separate user workflow or a separate application.

The canonical public entry point is
`spec_formatter.format_specifications()`. `gui.py` is a thin client of that API.
Architect and target inputs are immutable; only private snapshots and new
output files may be changed.

## Canonical architecture

`spec_formatter.pipeline.format_specifications()` owns the business flow:

```text
architect DOCX
  -> immutable snapshot and bounded package extraction
  -> styled/ignored role classification and source-derived portable styles
  -> bounded shell capture (styles, theme/defaults, settings, layout, headers/footers)
  -> strict checksummed *.phase1 profile
  -> versioned cache namespace and exact compatibility validation

target DOCX files
  -> immutable per-target snapshot in a short generated temp path
  -> bounded extraction and deterministic/AI styled-or-ignored dispositions
  -> one immutable ApplicationPolicy
       format_only: target-owned text and numbering
       csi_to_canadian: fail-closed hierarchy conversion
  -> collision-safe style import and full architect shell application
  -> mode-specific content, numbering, structure, and package invariants
  -> atomic DOCX publication in one timestamped run directory

run
  -> per-target audit JSON
  -> redacted run.log
  -> run.json written after target results and audits
```

Targets are independent: one target can fail without discarding other validated
outputs. A failed target never publishes a partial DOCX. A run directory and
manifest still record partial or total failure.

## Repository map

```text
spec_formatter/pipeline.py
    canonical public orchestration, profile cache, isolated runs, manifests
spec_formatter/template_analysis.py
    namespaced facade over architect analysis and bundle validation
spec_formatter/style_application/
    target extraction, shell application, numbering/header import, invariants
spec_formatter/style_application/core/application_policy.py
    immutable mode-dependent mutation contract
spec_formatter/style_application/core/classification.py
    numbering-aware target dispositions and paragraph application
spec_formatter/style_application/core/style_import.py
    effective style materialization and collision-safe import
phase1_pipeline.py, phase1_bundle.py, docx_decomposer.py
    architect snapshot, analysis, profile construction, and validation
gui.py
    input collection, immutable active-run display, progress, and results
schemas/
    formal architect-profile contracts
tests/
    unit, adversarial, contract, integration, GUI, and round-trip regressions
output/spec_test_corpus_smoke.py
    offline realistic-corpus checks against this repository's namespaced engine
```

`phase1_pipeline.run_phase1()` remains a compatibility and internal profile
builder surface. New integrations must call the unified public API.
`apply_instructions()` is a legacy developer surface and must not be wired into
the GUI or target application.

## Non-negotiable invariants

### 1. Immutable source identities

Every architect and target is copied to a private snapshot while size,
modification time, and SHA-256 are checked. All processing uses that snapshot;
the live input is checked again before target publication. Never mutate an input
or treat a later live-file read as artifact authority.

### 2. One explicit application policy

Resolve `conversion_mode` once with `application_policy_for_mode()` and pass the
same policy through numbering import, style import, paragraph application, and
validation. Do not recreate mode checks independently in downstream modules.

- `format_only` preserves target body text and target numbering semantics and
  does not import architect body numbering.
- `csi_to_canadian` performs only the existing fail-closed supported hierarchy
  conversion and may import architect numbering for classified roles.
- Both modes apply the architect's complete shell.

### 3. Explicit disposition coverage

Every visible classifiable target paragraph must occur exactly once in one of:

- `classifications`: styled CSI content
- `ignored_paragraphs`: non-CSI/editorial content with a non-empty reason

The sets are disjoint and cover the complete classifiable universe. Empty
structural paragraphs, table paragraphs, drawings, and text boxes are recorded
out of scope. Ignored paragraphs receive no paragraph or run edits. Missing,
duplicate, overlapping, or unknown dispositions fail closed; never restore
nearest-neighbor fallback.

### 4. Format-only numbering is target-owned

Snapshot effective target numbering before shell/style changes, including
numbering inherited through `basedOn`. Materialize target `numPr` when changing
its style would otherwise lose that inheritance. Detach imported body styles
from architect numbering and suppress architect numbering on originally
unnumbered paragraphs. Before publication, prove unchanged body text, unchanged
effective numbering semantics, and preservation of all original target
numbering definitions.

Automatic numbering evidence precedes text-only heuristics. An automatically
numbered paragraph whose stored text is `GENERAL` can be a PART, while an
automatically numbered requirement beginning `Section ...` is not a SectionID.

### 5. Architect formatting is source-derived and collision-safe

The LLM selects roles/dispositions, not XML formatting. Role styles come from
validated architect exemplars. Resolve paragraph formatting through the full
architect `basedOn` chain, and remove a target direct property only when the
effective architect style supplies that property. Never remove numbering,
`sectPr`, tracked changes, or protected subtrees as generic formatting cleanup.

Never replace an existing target style ID, including `Normal`. Clone a
conflicting architect style and its dependencies under deterministic private
IDs, rewrite `basedOn`/`next`/`link` and imported header/footer references, and
reject a deterministic namespace collision with different content.

### 6. Profile bundle and cache are strict boundaries

Target application consumes the complete `.phase1` directory after strict
manifest validation; loose registries are not a valid handoff. Cache profiles
under a versioned contract namespace and require exact source hash, producer,
classifier, model, and prompt compatibility. Bump the profile contract when a
consumer-visible bundle assumption changes.

Architect analysis remains observational: derive generated styles in
`portable_styles.xml`; preserve byte-exact `source_styles.xml` and optional
`source_settings.xml`; never retag or publish a normalized architect DOCX.

### 7. Isolated, atomic, auditable publication

Create one `<UTC timestamp>_<mode>_<run-id>` directory per validated invocation,
before template analysis begins, so template/profile initialization failures
still publish a failed manifest, run log, and per-target not-started audits.
Stage each target under a short generated system-temp path, validate the complete
DOCX, and atomically publish it into that run directory. Then atomically write
per-target audits, `run.log`, and finally `run.json`. Never put secrets or
document text in metadata. Existing run directories and flat legacy outputs
are immutable history.

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
| `SUBPARAGRAPH_LEVEL_5` | `CSI_SubparagraphLevel5__ARCH` |
| `SUBPARAGRAPH_LEVEL_6` | `CSI_SubparagraphLevel6__ARCH` |
| `SUBPARAGRAPH_LEVEL_7` | `CSI_SubparagraphLevel7__ARCH` |
| `SUBPARAGRAPH_LEVEL_8` | `CSI_SubparagraphLevel8__ARCH` |
| `END_OF_SECTION` | `CSI_EndOfSection__ARCH` |

Role expectations come from text signals and effective Word numbering, including numbering inherited through paragraph styles. Do not treat arbitrary `A.`, `1.`, or `a.` text globally as proof of CSI hierarchy. Validate exemplars, role/style coherence, numbering family/level coverage, style inheritance, and style references against the source catalogs.

## Module responsibilities

### `spec_formatter/pipeline.py`

- `format_specifications()` is the canonical public entry point.
- Validates inputs and modes, prepares a compatible cached profile, allocates
  the isolated run directory, snapshots/dispatches targets, atomically publishes
  metadata, and returns `FormatRunResult`.
- `output_dir` is an output root; the returned `output_dir` is a backward-
  compatible alias of the concrete `run_dir`.
- Folder expansion excludes the architect only when discovered through a
  folder. An explicitly supplied architect target must reach validation and
  fail.

Keep UI concerns out of this module and content/classification decisions in the
target engine. Redact the API key from every error/log/manifest path.

### `spec_formatter/style_application/core/application_policy.py`

Owns all mode-dependent decisions. Add a policy field rather than scattering a
new conversion-mode conditional across the pipeline. Its contract version must
be recorded in `run.json` and changed when policy semantics change.

### `spec_formatter/style_application/batch_runner.py`

- Loads one validated profile and prepares/classifies targets.
- `_apply_classified_target()` is the shared application path for single and
  batch flows; do not duplicate environment/numbering/style sequencing.
- Captures target styles/numbering before shell mutation, applies the selected
  policy, produces audit/numbering checks, validates, and packages the result.

### `spec_formatter/style_application/core/classification.py`

- Builds the target slim bundle and deterministic CSI/ignore dispositions.
- Gives effective Word numbering stronger precedence than text-only signals.
- Validates exact disjoint coverage and rejects deterministic overrides.
- Applies only styled entries, leaves ignored entries exact, resolves effective
  architect paragraph properties through `basedOn`, and enforces Format-only
  text/numbering invariants.

Paragraph indices are tied to the `word/document.xml` paragraph sequence.
Preserve that index and visible-text contract when changing XML parsing.

### `spec_formatter/style_application/core/style_import.py`

Imports only the requested architect style closure. Materialize effective
formatting, detach Format-only body styles from architect numbering, namespace
every target-ID collision deterministically, and rewrite dependency references.
Return the source-to-final style-ID map to every body/header/footer consumer.

### Target shell, packaging, and invariants

- `arch_env_applier.py` applies document defaults, theme/settings,
  compatibility, and canonical section/page layout.
- `header_footer_importer.py` and `numbering_importer.py` import bounded
  dependency sets and remap relationships/IDs.
- `phase2_invariants.py` verifies body, numbering, protected structure,
  section, header/footer, relationship, and package contracts.
- `docx_decomposer.py` extracts targets safely; `docx_patch.py` assembles and
  validates replacements before publication.

### Architect profile modules

`phase1_pipeline.py` snapshots and analyzes the architect. `phase1_bundle.py`
creates and validates the complete bundle. Root `docx_decomposer.py` builds the
architect slim bundle, derives portable styles without changing the source,
and emits role metadata. `llm_classifier.py`, `paragraph_rules.py`,
`arch_env_extractor.py`, and `phase1_validator.py` own architect classification,
signals, shell capture, and cross-contract validation respectively.

### `gui.py`

Owns input collection, background execution, immutable active-run display,
progress/log rendering, and final status. It calls `format_specifications()` and
displays `FormatRunResult`. Lock every run-affecting control while work is
active, display all target processor log lines and audit counts, and open the
actual `run_dir`. Never recreate pipeline business logic in the GUI.

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

The template registry stores normalized source-derived XML fragments. Current
capture policy does not canonicalize whitespace but does strip volatile rsid
attributes and proofing markers. Do not call these fields "raw XML."

Use these terms consistently:

- `source_styles.xml`: byte-exact original styles part
- `source_settings.xml`: byte-exact original settings part, when present
- `portable_styles.xml`: generated stylesheet for Phase 2
- `arch_template_registry.json`: bounded normalized environment capture

The retired names `arch_styles_raw.xml` and `arch_settings_raw.xml` are not bundle artifacts.

## Run artifacts and public results

Each invocation creates:

```text
<UTC timestamp>_<mode>_<run-id>/
  *_FORMATTED.docx or *_CANADIAN_FORMATTED.docx
  target-<sequence>-<source-hash>.audit.json
  run.log
  run.json
```

`run.json` records the run/mode/status/timestamps; application, policy, and
profile contract versions; output paths; architect path/hash; cache bundle
identity; model/prompt fingerprints; target/output hashes; audit paths;
disposition counts; numbering checks; durations; and redacted errors. It must
not contain API keys or document text.

`FormatRunResult` retains `success`, `succeeded`, `failed`, `output_paths`, and
the historical `output_dir`, and adds `run_id`, `conversion_mode`,
`output_root`, `run_dir`, and `manifest_path`. `TargetFormatResult` retains its
historical fields and adds source/output SHA-256, `audit_path`,
`audit_summary`, application audit details, and numbering checks. Additive
fields must keep safe defaults so existing test doubles and callers continue
to work.

## Development commands

```bash
pip install -r requirements-dev.txt
python -m pytest -q
python gui.py
python output/spec_test_corpus_smoke.py
```

Headless usage is through Python:

```python
from pathlib import Path
from spec_formatter import format_specifications

result = format_specifications(
    architect_template=Path("template.docx"),
    target_specs=[Path("targets")],
    output_dir=Path("output"),
    api_key="...",
    conversion_mode="format_only",
)

print(result.run_dir, result.manifest_path, result.output_paths)
```

## Change checklist

Before considering a formatter change complete:

1. Confirm architect and target sources remain unchanged.
2. Confirm every classifiable target paragraph is styled or ignored exactly once.
3. Exercise both application policies; prove Format-only text and numbering are
   unchanged and Canadian conversion still fails closed.
4. Confirm generated style inheritance, collision remapping, and header/footer
   style references resolve correctly without replacing target style IDs.
5. Validate the complete architect bundle and versioned cache compatibility.
6. Validate each output package, audit, `run.log`, `run.json`, hashes, and
   redaction behavior for success, partial failure, and total failure.
7. Test deep Windows paths and folder discovery containing the architect.
8. Run focused tests, the complete suite, and the local realistic-corpus smoke.
9. Render and inspect every page of representative original and output DOCX
   files when formatting or shell behavior changes.
10. Update prompts, schemas, validators, README, and this guide together for
    contract changes.

## Common mistakes

- Calling `apply_instructions()` from the production pipeline.
- Treating the extraction directory as an output deliverable.
- Treating the selected output root as the concrete run directory.
- Writing loose formatted files or logs directly into the output root.
- Reimplementing mode checks outside `ApplicationPolicy`.
- Importing architect body numbering in Format-only.
- Treating visible text as stronger evidence than effective Word numbering.
- Restyling an ignored paragraph or silently dropping it from coverage.
- Replacing a target style merely because an architect style uses the same ID.
- Inspecting only a direct style's `pPr` and ignoring its `basedOn` chain.
- Describing the two registries as the complete handoff.
- Calling normalized registry fragments "raw" or "complete VM state."
- Filling missing classifications from adjacent paragraphs.
- Dropping style-only numbering definitions from the slim catalog or registry.
- Trusting an LLM-provided `basedOn` instead of the exemplar's source style.
- Following an external relationship or a path that escapes the package.
- Publishing before the DOCX is fully copied and revalidated.
- Recording secrets or paragraph text in run metadata.
- Pointing the corpus smoke at a sibling checkout instead of the namespaced
  implementation in this repository.

## Platform and license

Runtime is Python 3.9+ with Windows as the primary GUI platform. Core package processing is intended to remain portable.

Copyright 2025 Abraham Borg. All Rights Reserved. Proprietary software; no license is granted without written permission.
