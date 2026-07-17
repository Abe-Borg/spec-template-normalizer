# Design System - Phase 1 Desktop UI

This document defines the user-interface language for the Spec Template Normalizer Phase 1 desktop application. The GUI is a thin CustomTkinter client of the headless `phase1_pipeline.run_phase1()` workflow.

The product promise shown in the interface must match the implementation: Phase 1 analyzes an immutable snapshot, never changes the selected DOCX, and publishes one complete, validated `.phase1` bundle for Phase 2.

## Product principles

- **Source safety is visible.** Copy must say “analyze,” “snapshot,” and “publish bundle,” not “normalize document,” “apply styles,” or “modify template.”
- **One run, one handoff.** Success points to the unique `.phase1` directory, never to loose registry files or an extracted working directory.
- **Progress describes real stages.** Status messages mirror the headless pipeline and do not imply the source package is being edited.
- **Failure is explicit.** Ambiguous classification, package limits, malformed input, and integrity failures are errors, not warnings that can be ignored.
- **Technical detail is available without clutter.** The main window stays compact; the activity log and help dialogs explain provenance, checksums, limits, and bundle contents.

## Supported user flow

```text
Select source DOCX + API key + output root
  -> Run Phase 1
  -> snapshot source
  -> safely extract and inspect
  -> classify styled/ignored universe
  -> derive portable styles separately
  -> capture bounded environment
  -> validate and atomically publish
  -> show coverage and *.phase1 bundle path
```

At no point should the UI suggest that a modified DOCX will be returned. The output-folder control selects a parent directory; the pipeline creates a unique child directory.

## Visual language

The application uses a dark, card-based desktop layout with a restrained blue accent. Segoe UI carries interface text and Consolas carries paths, identifiers, and log output.

### Colors

| Token | Value | Use |
|---|---|---|
| `bg_dark` | `#0D0D0D` | Window canvas |
| `bg_card` | `#1A1A1A` | Input and activity cards |
| `bg_input` | `#252525` | Fields, log surface, secondary controls |
| `border` | `#333333` | Field and secondary-button borders |
| `text_primary` | `#FFFFFF` | Title, selected values, important output |
| `text_secondary` | `#B0B0B0` | Labels, descriptions, normal status |
| `text_muted` | `#707070` | Section labels, hints, disabled metadata |
| `accent` | `#3B82F6` | Primary action, progress, links |
| `accent_hover` | `#2563EB` | Primary-action hover |
| `accent_glow` | `#60A5FA` | Reserved focus/completion accent |
| `success` | `#22C55E` | Validated publication and full coverage |
| `success_glow` | `#4ADE80` | Reserved success emphasis |
| `warning` | `#F59E0B` | Nonfatal attention states only |
| `error` | `#EF4444` | Rejected input or failed run |

Color always carries state. Never use success green until the bundle has passed validation and atomic publication.

### Typography

| Element | Font | Size | Weight/color |
|---|---|---:|---|
| App title | Segoe UI | 24 | Bold, `text_primary` |
| Subtitle | Segoe UI | 12 | Regular, `text_secondary` |
| Card label | Segoe UI | 11 | Bold, `text_muted`, uppercase |
| Input label | Segoe UI | 12 | Regular, `text_secondary` |
| Path/API field | Consolas | 12 | Regular, `text_primary` |
| Primary button | Segoe UI | 14 | Bold |
| Secondary/help button | Segoe UI | 12 | Regular, `text_secondary` |
| Activity log | Consolas | 12 | Regular, `text_secondary` |
| Status bar | Segoe UI | 11 | State color |
| Help dialog title | Segoe UI | 18 | Bold, `text_primary` |

Do not use monospace for explanatory prose. Use it for machine identities such as paths, hashes, artifact names, paragraph indices, and errors.

## Window and layout

### Main window

- Default: `900 x 750`
- Minimum: `750 x 600`
- Canvas: `bg_dark`
- Outer padding: 24 px
- Card radius: 8 px
- Vertical card gap: 12 px

Content order:

1. Header: product title, subtitle, **How to Use**, and **How It Works**.
2. Collapsible **INPUTS** card.
3. Full-width **Run Phase 1** button.
4. Four-pixel indeterminate progress bar while running.
5. Expandable **ACTIVITY LOG** card, which consumes remaining height.
6. Compact status line.

The log is the only expanding region. Do not add result dashboards or file grids that compete with the core one-file workflow.

### Inputs card

The card uses a two-column grid:

- Label column: 100 px, left aligned.
- Control column: flexible width with 8 px leading gap.
- Row spacing: 8 px vertically.
- Card padding: 16 px horizontal, 12 px header, 16 px below content.

Required controls:

| Label | Control | Behavior |
|---|---|---|
| Template | Path entry + Browse | Accept one `.docx`; choosing it initializes the output root to its parent |
| API Key | Masked entry + Show/Hide | Pre-fill from `ANTHROPIC_API_KEY` when available |
| Output Folder | Path entry + Browse | Select the parent in which a unique `.phase1` directory will be created |

Entries and browse controls are 36 px high. Password masking remains on by default.

### Activity log

The log uses a `bg_input` surface, 4 px radius, word wrapping, and a Consolas 12 font. It is read-only except for programmatic appends. The card header includes a low-emphasis **Clear** action.

Messages should be short, factual, and stage-oriented. Canonical progress language includes:

- `Snapshotting <file>...`
- `Safely unpacking the DOCX package...`
- `Reading paragraph, style, and numbering structure...`
- `Classifying <n> paragraphs...`
- `Deriving portable styles without changing the source package...`
- `Capturing the source formatting environment...`
- `Validating checksums and publishing the bundle...`
- `Published validated bundle: <name>.phase1`
- `Coverage: 100.0% (<handled>/<classifiable>)`

Do not log API keys, complete prompt bodies, or base64 media. Errors may include paragraph indices and artifact names needed to diagnose the failure.

## Component states

### Primary action

| State | Label | Color | Behavior |
|---|---|---|---|
| Ready | `Run Phase 1` | `accent` | Enabled |
| Running | `Processing...` | `accent` | Disabled; progress bar active |
| Published | `✓ Complete` | `success` | Disabled briefly, then reset |
| Failed | `✕ Failed` | `error` | Disabled briefly, then reset |

“Complete” means the atomic rename succeeded. Validation that only reached a staging directory is not completion.

### Status line

- Initial: `Ready` in `text_secondary`.
- Client validation error: concise reason in `error`.
- Running: `Running...` in `text_secondary`.
- Success: `Success - Coverage: ...` in `success`.
- Failure: `Failed - see log for details` in `error`.

Coverage describes explicit styled-or-ignored handling. Avoid phrasing it as “percent of paragraphs styled,” because ignored candidate paragraphs legitimately count as handled.

### Collapsible cards

Card headers are fully clickable, use a hand cursor, and include a muted Consolas arrow:

- `▼` expanded
- `▶` collapsed

Keep the section label visible in either state. Collapsing the activity log must not stop or pause a run.

### Secondary buttons

- Background: `bg_input`
- Hover: `border`
- One-pixel `border` outline
- Text: `text_secondary`
- Corner radius: framework default or 6 px
- Heights: 36 px for browse/key controls, 32 px for help/dialog actions, 24 px for log utility actions

## Output presentation contract

Help text and any future result view must identify the entire bundle as the deliverable. List these exact names:

| File | User-facing description |
|---|---|
| `phase1_bundle_manifest.json` | Version, source identity, artifact sizes, and SHA-256 checksums |
| `arch_style_registry.json` | CSI role/style and numbering metadata |
| `arch_template_registry.json` | Supported formatting-environment capture |
| `classification_audit.json` | Paragraph-level styled/ignored/out-of-scope record |
| `source_styles.xml` | Exact source stylesheet |
| `portable_styles.xml` | Source styles plus safely derived CSI styles |
| `source_settings.xml` | Optional exact source settings part |

Never display the retired names `arch_styles_raw.xml` or `arch_settings_raw.xml`. Never tell users to keep “the two registries” together; the manifest, audit, and XML artifacts are part of the contract. The source DOCX is not copied into the published bundle.

Use this handoff sentence consistently:

> Give the complete `.phase1` folder to Phase 2. Do not edit, rename, add, or remove files inside it.

## Help content requirements

### How to Use

Explain prerequisites, the three inputs, the run action, expected duration only as a nonbinding estimate, the final coverage status, and the complete bundle contents. State that rerunning is safe because the selected DOCX is never modified.

### How It Works

Explain the two-phase boundary and these real stages: stable snapshot, safe extraction, slim structural bundle, AI classification, local style derivation, separate portable stylesheet, bounded environment capture, validation, and atomic publication.

Avoid these obsolete claims:

- Phase 1 “applies styles” or “inserts pStyle tags” into the template.
- The extracted working directory is an output.
- The output is a normalized DOCX or two loose JSON files.
- The template registry is a complete rendering VM snapshot.
- Header/footer and relationship bytes are modified and then checked for drift.

## Error communication

Error messages should answer three questions: what was rejected, which source index/artifact/limit is involved, and whether the user can retry safely.

Important classes:

- Source changed during snapshot: finish saving, then rerun.
- Package traversal, symlink, duplicate, compression, or size violation: reject the source as unsafe.
- External relationship: record but never fetch; a broken required internal relationship is an error.
- Header/footer media over 16 MiB per item or 64 MiB total: reduce embedded source assets.
- Estimated classifier input over 150,000 tokens: reduce template scope.
- Incomplete classification: show unresolved paragraph indices; do not imply an automatic neighbor-style fallback.
- Bundle checksum/contract failure: no final bundle was published; rerun from source.

Warnings must not be used for a condition that invalidates the Phase 1 contract.

## Accessibility and interaction

- Maintain text labels in addition to color for every state.
- Preserve strong contrast between text hierarchy levels and surfaces.
- Keep keyboard focus visible on entries and buttons.
- Do not place critical instructions only in the activity log; repeat the final bundle handoff in help/result copy.
- Keep API keys masked and out of clipboard/log helpers unless the user explicitly chooses **Show**.
- Help dialogs are modal, scrollable, at least `640 x 480`, and provide an explicit **Close** button.

## Architecture boundary for UI changes

The GUI may collect values, call `run_phase1()`, receive progress strings, and display `Phase1Result`. It must not:

- extract the DOCX itself;
- call the classifier directly;
- derive or apply styles;
- construct registries or manifests;
- copy artifacts into output paths;
- decide classification coverage;
- publish or overwrite bundle directories.

Those operations belong to the headless pipeline so GUI and automated callers share the same safety and validation guarantees.
