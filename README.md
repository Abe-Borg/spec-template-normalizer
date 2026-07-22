# Specification Formatter

One application takes an architect's Word specification template and one or
more target specifications, then produces formatted target DOCX files. Users do
not switch programs or manually transfer intermediate files.

The original template and target files are never modified.

## What the app does

1. Analyzes the architect's `.docx` template and captures its CSI paragraph
   styles, numbering, fonts, headers, footers, settings, and page layout.
2. Validates and caches that template analysis by the template's SHA-256 hash,
   engine version, prompt hashes, classifier model, and profile-contract
   version.
3. Classifies each visible target paragraph as a CSI role or an explicit,
   auditable ignore disposition; empty structural paragraphs and table content
   are recorded out of scope.
4. Applies one explicit mode policy:
   - **Format only** keeps the target's text, list ownership, markers, levels,
     counters, and restart semantics.
   - **Canadian CSC PageFormat** converts supported CSI hierarchy to the
     architect's demonstrated automatic numbering.
5. Applies the architect's complete document shell: source-derived CSI
   formatting, theme/defaults, compatibility settings, page layout, and
   default/even/first headers and footers.
6. Validates content, numbering, protected structure, and the complete DOCX
   package before publishing it.

The internal template profile remains a checksummed integrity boundary, but it
is not part of the user workflow. If the architect template has not changed,
the app safely reuses its validated analysis from the operating system's
per-user application cache. A profile is reused only from the current cache
contract namespace; changing that contract invalidates older cached profiles
without deleting them.

## Download for Windows

Most users do not need Python. Download the latest **SpecificationFormatterSetup.exe**
from the [Releases page](https://github.com/abe-borg/spec-template-normalizer/releases/latest)
and run it. It installs per-user (no admin prompt) with a Start-menu shortcut.

The app is **not code-signed**, so the first time you run the installer Windows
SmartScreen shows *"Windows protected your PC."* Choose **More info → Run
anyway**. (The download is still integrity-checked: the in-app updater verifies
every update's SHA-256 before running it.)

### Updates

The app checks for updates once a day on launch and via the **Check for Updates**
button in the footer. When a newer version exists it shows what's new and lets
you **Download & Install** (the app closes so the installer can replace it),
**Skip this Version**, or decide **Later**. Update checks can be turned off with
the `SPEC_FORMATTER_DISABLE_UPDATE_CHECK` environment variable.

## Install from source (developers)

Python 3.10 or newer is required.

```powershell
python -m venv venv
.\venv\Scripts\python.exe -m pip install -r requirements.txt
```

An Anthropic API key is required when a new architect template must be analyzed
or when a target contains paragraphs that cannot be classified locally. The key
is used for the active run. Tick **Remember** next to the key field to save it to
the Windows Credential Manager (via `keyring`) so it is pre-filled on the next
launch; leave it unticked and the key is held only in memory. The field is also
pre-filled from the `ANTHROPIC_API_KEY` environment variable when set.

## Run the app

```powershell
.\venv\Scripts\python.exe gui.py
```

In the single window:

1. Choose the architect's template DOCX.
2. Add target DOCX files, or add a folder containing target specs.
3. Choose **Format only** or **Convert CSI hierarchy to Canadian CSC PageFormat**.
4. Choose the output folder.
5. Enter the Anthropic API key when needed.
6. Click **Format Specs**.

Each run creates an isolated directory below the selected output root:

```text
<UTC timestamp>_<mode>_<run-id>/
  <target>_FORMATTED.docx              # Format only
  <target>_CANADIAN_FORMATTED.docx     # Canadian mode
  target-<sequence>-<source-hash>.audit.json
  run.log
  diagnostics.jsonl
  run.json
```

When selected targets from different folders share a filename, the app adds a
stable source suffix so neither output can overwrite the other. Failed reruns
therefore cannot make an older output appear current. `run.json` records the
mode, application/profile contract versions, template and target hashes, model
and prompt fingerprints, cache identity, output hashes, disposition counts,
numbering checks, durations, a `diagnostics` rollup, and errors. API keys and
document text are never written to run metadata.

`diagnostics.jsonl` is the detailed, structured diagnostics stream: one JSON
object per phase event (`seq`, `ts`, `level`, `component`, `event`, and a
`fields` object of counts/timings such as per-phase `duration_ms`, styles
imported, numbering remaps, and paragraphs modified). It complements the
human-readable `run.log`. Diagnostics carry only numbers and short structural
identifiers -- never document text or secrets -- and the verbosity is set with
`diagnostics_level` (`debug`/`info`/`warning`/`error`, default `info`) or the
`SPEC_FORMATTER_DIAGNOSTICS_LEVEL` environment variable, which overrides it.

Folder discovery ignores Word lock files, current `_FORMATTED.docx` outputs,
and legacy `_PHASE2_FORMATTED.docx` outputs. If the architect is present in a
selected folder it is excluded from discovery; explicitly selecting the
architect as a target remains an error.

## Format-only mode

Format-only treats target content and numbering as authoritative. Before any
architect parts are applied, the engine snapshots each paragraph's effective
numbering, including numbering inherited through a paragraph style. Imported
body-role styles are detached from architect numbering, inherited target
`numPr` is materialized where needed, and existing target numbering definitions
remain intact. Publication fails if body text or effective list semantics
change.

Only paragraphs with validated CSI roles receive paragraph/run formatting.
Non-CSI and editorial content is returned as `ignored_paragraphs` with a reason
and its paragraph XML is left untouched. Tables, drawings, and text boxes are
outside body restyling. The architect shell is document-global, so ignored or
out-of-scope content can still reflow under the architect's page geometry,
theme, and defaults.

Architect styles are always imported into deterministic private `SF_*`
namespaces. Existing target style IDs, including built-ins such as `Normal`,
are never overwritten, and every imported dependency reference is rewritten to
the private namespace. Direct target paragraph and run properties are removed
only when the effective architect style supplies the same property through its
full `basedOn` chain.

## Canadian CSC PageFormat mode

Canadian mode is one option inside the same GUI and pipeline. It recognizes
typical CSI article and list levels, removes manually typed CSI markers without
altering the requirement text, and retargets both typed and automatic source
numbering to the architect template's Canadian automatic Word numbering. The
recognized hierarchy continues through the deeper `1)`, `a)`, `(1)`, and
`(a)` subparagraph levels supported by Word's nine-level list model.

For a fail-closed conversion, the architect template must demonstrate true
automatic Canadian numbering for every numbered role used by the targets:
articles such as `1.1` and subordinate levels such as `.1`. A template that
still uses `A. / 1. / a.`, literal typed numbering, or no numbering for a used
role is rejected in Canadian mode instead of producing a misleading result.
Mixed paragraphs that contain both a typed marker and Word automatic numbering
are also rejected as ambiguous.

The first implementation deliberately supports the sequence it can prove:
source counters must start at 1, be contiguous, and appear under their expected
parent level. Counter gaps, custom starts/restarts, or automatic source
list-instance changes stop that target without publishing a partial conversion.
Every paragraph sharing a converted automatic source list must
participate in the conversion; a filtered table or boilerplate list item is
rejected because omitting it would shift later counters. A typed Canadian
decimal such as `.1` is treated as a marker only when Word stores a structural
list tab after it; this prevents a leading value such as `.125 mm` from being
deleted. A two-component article such as `1.1 SUMMARY` may use a normal space
when its heading-like text, active PART, and contiguous counter prove the
hierarchy. A markerless paragraph that the classifier assigns to a numbered
role is instead preserved unchanged, omitted from numbered style application,
and reported as a warning; the converter does not invent a list item. For
targets containing ARTICLE headings,
the target must contain a preceding classified PART heading, and the current
engine requires the architect's PART, ARTICLE, and used subordinate roles to be
levels of one coherent automatic multilevel Word list. That is an application
safety limitation, not a claim that CSC PageFormat requires one particular
Word/OOXML implementation.

Conversion counts and diagnostics are included in the application's saved
activity log, including when a later validation or publication step fails.

This mode converts numbering hierarchy and presentation only. It does **not**
reorder articles, replace US codes or standards, convert units, change spelling
or terminology, revise technical requirements, or certify NMS compliance.

## Build the Windows installer

The distributable Windows app is a PyInstaller **one-folder** build wrapped by an
Inno Setup installer. On Windows, from the repo root:

```powershell
.\venv\Scripts\python.exe -m pip install -r requirements.txt -r requirements-build.txt
.\venv\Scripts\pyinstaller.exe packaging\windows\specification-formatter.spec --noconfirm --clean
& "${env:ProgramFiles(x86)}\Inno Setup 6\ISCC.exe" /DMyAppVersion=1.0.0 packaging\windows\installer.iss
```

The app folder is written to `dist\SpecificationFormatter\` and the installer to
`dist\installer\SpecificationFormatterSetup.exe`. Releases are normally built and
published automatically by `.github/workflows/release.yml` on a `vX.Y.Z` tag — see
[docs/RELEASE_WINDOWS.md](docs/RELEASE_WINDOWS.md) for the full runbook. The
legacy `build_app.ps1` one-file script is retained for quick local smoke builds.

## Headless API

`spec_formatter.format_specifications()` is the canonical programmatic entry
point:

```python
from pathlib import Path
from spec_formatter import format_specifications

result = format_specifications(
    architect_template=Path("Architect Template.docx"),
    target_specs=[Path("Mechanical.docx"), Path("Electrical Specs")],
    output_dir=Path("Formatted Specs"),
    api_key="...",
    conversion_mode="csi_to_canadian",  # omit for format-only behavior
    diagnostics_level="debug",          # optional; default "info"
)

for target in result.targets:
    print(target.source_path, target.success, target.output_path or target.error)

print(result.run_id, result.conversion_mode)
print(result.output_root, result.run_dir, result.manifest_path)
print(result.diagnostics_path)  # diagnostics.jsonl for this run
```

Target inputs may be individual DOCX paths or folders. Folder expansion is
non-recursive. Multi-file runs are independent: one corrupt target is reported
as a failure without discarding valid outputs from other targets.
`output_dir` is the output **root**; `result.output_dir` remains a
backward-compatible alias of the concrete `result.run_dir`.

`FormatRunResult` exposes `run_id`, `conversion_mode`, `output_root`, `run_dir`,
`manifest_path`, `diagnostics_path`, `targets`, `success`, `succeeded`,
`failed`, and `output_paths`. Each `TargetFormatResult` includes source/output
hashes, `audit_path`, disposition counts, numbering checks, structured
`diagnostics` events, processor log lines, duration, conversion report, and any
error in addition to the historical source/success/output fields.

## Internal architecture

The app keeps the strongest safety properties of both original programs:

```text
architect template DOCX
    -> stable snapshot and bounded extraction
    -> complete classification and source-derived styles
    -> checksummed, atomically published internal template profile
    -> strict profile validation

target specification DOCX files
    -> bounded extraction
    -> deterministic/AI styled-or-ignored paragraph dispositions
    -> immutable application policy selected at the orchestration boundary
    -> target-owned numbering preservation OR fail-closed Canadian conversion
    -> collision-safe styles and complete architect shell application
    -> mode-specific content, numbering, stability, and package validation
    -> atomic publication into an isolated, manifested run directory
```

The architect-template engine remains available through
`phase1_pipeline.run_phase1()` for compatibility. The target application engine
is namespaced under `spec_formatter.style_application` to prevent module-name
collisions. New integrations should call only the unified public API.

### Internal profile and classification contract

The architect profile is an implementation detail, not a second user-facing
program. Each atomically published profile has a unique directory and a
manifest that records its format and producer versions, source identity,
classifier identity, required artifacts, byte sizes, and SHA-256 checksums.
Those artifacts include the role/style and template-environment registries,
the complete classification audit, and byte-exact source styles and settings
alongside portable source-derived styles. Strict validation rejects missing,
altered, or unlisted profile files.

Profiles live under a versioned cache namespace (`contract-v<version>`). A
cache hit also requires the exact source hash, producer version, classifier
identity, and prompt hashes, so a wire-contract change cannot silently reuse an
older profile.

Formal JSON contracts remain in `schemas/`. The template-environment registry
is a bounded, normalized representation; use its exact `source_styles.xml` and
optional `source_settings.xml` companion parts when byte identity matters.

The supported CSI roles are `SectionID`, `SectionTitle`, `PART`, `ARTICLE`,
`PARAGRAPH`, `SUBPARAGRAPH`, `SUBSUBPARAGRAPH`, the deeper semantic roles
`SUBPARAGRAPH_LEVEL_5` through `SUBPARAGRAPH_LEVEL_8`, and `END_OF_SECTION`. Every
in-scope paragraph must be explicitly styled or ignored; unresolved coverage
fails closed and is never filled from a neighboring paragraph. The classifier
selects roles, dispositions, and exemplars only. Formatting properties come
from the architect's OOXML, including style inheritance and numbering. Empty
structural paragraphs and table content are out of scope, while visible
section-break paragraphs remain classifiable and retain their `sectPr`.

## Safety guarantees

- Input DOCX packages are treated as untrusted ZIP containers. Unsafe paths,
  duplicate members, symbolic links, suspicious compression, oversized parts,
  and missing required parts fail closed.
- Extraction is capped at 10,000 entries, 512 MiB total, 128 MiB per part, and
  a 1,000:1 compression ratio. Header/footer media is capped at 16 MiB per
  asset and 64 MiB total.
- Internal relationship targets must resolve inside the package. External
  targets are recorded but never fetched, and classifier input is capped at an
  estimated 150,000 tokens.
- Architect and target inputs are read-only; outputs are separate files.
- Template profiles are reused only after manifest, size, checksum, source hash,
  producer-version, prompt/model fingerprint, and cache-contract validation.
- Tables and drawing/text-box content are excluded from paragraph restyling.
- Explicit ignored paragraphs receive no paragraph/run edits; unresolved or
  overlapping dispositions fail closed.
- Imported architect styles never replace an existing target style ID.
- Format-only verifies unchanged body text and effective target numbering and
  preserves every pre-existing target numbering definition.
- Canadian conversion edits only recognized leading numbering markers in
  classified paragraphs and verifies that substantive text and protected OOXML
  remain unchanged.
- Short generated temporary paths avoid carrying user-controlled deep paths
  into Windows staging.
- Every output is fully validated before atomic publication into its run folder.

## Tests

```powershell
.\venv\Scripts\python.exe -m pip install -r requirements-dev.txt
.\venv\Scripts\python.exe -m pytest -q
```

The suite includes the original template-analysis coverage, the vendored
style-application regression suite under its new namespace, unified workflow
and cache tests, and offline end-to-end DOCX round trips that verify styles,
numbering, settings, headers/footers, page layout, tables, text boxes, source
immutability, ignored dispositions, mode separation, partial-failure isolation,
run provenance, long paths, and package validity.

`tests/test_sanitized_format_only_corpus.py` builds a tracked, non-proprietary
154-paragraph reproduction of the supplied acceptance case and runs it through
the public unified application. It verifies all 121 inherited automatic list
items, the 19 critical markers, `GENERAL` as `PART`, byte-stable ignored XML,
and the architect shell without any sibling-repository dependency.

## Copyright Notice

**Copyright 2025 Abraham Borg. All Rights Reserved.**

This software and associated documentation files are proprietary. Unauthorized
copying, modification, distribution, or use is prohibited without the copyright
holder's express written permission.
