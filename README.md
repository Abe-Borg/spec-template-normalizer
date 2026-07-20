# Specification Formatter

One application takes an architect's Word specification template and one or
more target specifications, then produces formatted target DOCX files. Users do
not switch programs or manually transfer intermediate files.

The original template and target files are never modified.

## What the app does

1. Analyzes the architect's `.docx` template and captures its CSI paragraph
   styles, numbering, fonts, headers, footers, settings, and page layout.
2. Validates and caches that template analysis by the template's SHA-256 hash,
   engine version, prompt hashes, and classifier model.
3. Classifies each target spec's in-scope paragraphs.
4. Optionally converts conventional CSI numbering (`1.01 / A. / 1. / a. /
   1) / a) / (1) / (a)`)
   to the Canadian numeric hierarchy demonstrated by the architect template.
5. Applies the architect's formatting while preserving technical wording,
   tables, drawings, text boxes, and unmanaged document structure.
6. Validates each complete DOCX package before publishing it.

The internal template profile remains a checksummed integrity boundary, but it
is not part of the user workflow. If the architect template has not changed,
the app safely reuses its validated analysis from the operating system's
per-user application cache; the selected output folder contains only formatted
documents and run logs.

## Install

Python 3.10 or newer is required.

```powershell
python -m venv venv
.\venv\Scripts\python.exe -m pip install -r requirements.txt
```

An Anthropic API key is required when a new architect template must be analyzed
or when a target contains paragraphs that cannot be classified locally. The key
is used only for the active run and is not saved by the application.

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

Format-only outputs are named `<target>_FORMATTED.docx`; Canadian outputs are
named `<target>_CANADIAN_FORMATTED.docx`, so the two modes cannot silently
replace each other. When selected targets from different folders share the
same filename, the app adds a stable source suffix so neither output can
overwrite the other. A timestamped activity log is saved beside the formatted
documents.

Folder discovery ignores Word lock files, current `_FORMATTED.docx` outputs,
and legacy `_PHASE2_FORMATTED.docx` outputs.

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
parent level. Gaps, unnumbered list items, custom starts/restarts, or automatic
source list-instance changes stop that target without publishing a partial
conversion. Every paragraph sharing a converted automatic source list must
participate in the conversion; a filtered table or boilerplate list item is
rejected because omitting it would shift later counters. A typed Canadian
decimal such as `.1` is treated as a marker only when Word stores a structural
list tab after it; this prevents a leading value such as `.125 mm` from being
deleted. For targets containing ARTICLE headings,
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

## Build one Windows executable

Install the build dependencies, then run the included build script:

```powershell
.\venv\Scripts\python.exe -m pip install -r requirements-build.txt
.\build_app.ps1
```

The single-file application is written to
`dist\SpecificationFormatter.exe`.

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
)

for target in result.targets:
    print(target.source_path, target.success, target.output_path or target.error)
```

Target inputs may be individual DOCX paths or folders. Folder expansion is
non-recursive. Multi-file runs are independent: one corrupt target is reported
as a failure without discarding valid outputs from other targets.

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
    -> deterministic/AI paragraph classification
    -> optional fail-closed CSI-to-Canadian hierarchy conversion
    -> environment, numbering, style, header/footer, and layout application
    -> stability and complete-package validation
    -> atomic publication as *_FORMATTED.docx or *_CANADIAN_FORMATTED.docx
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
  and producer-version validation.
- Tables and drawing/text-box content are excluded from paragraph restyling.
- Formatting is applied through validated architect styles and numbering
  definitions; unknown roles are not guessed.
- Canadian conversion edits only recognized leading numbering markers in
  classified paragraphs and verifies that substantive text and protected OOXML
  remain unchanged.
- Each output is fully validated before it replaces an earlier formatted output.

## Tests

```powershell
.\venv\Scripts\python.exe -m pip install -r requirements-dev.txt
.\venv\Scripts\python.exe -m pytest -q
```

The suite includes the original template-analysis coverage, the vendored
style-application regression suite under its new namespace, unified workflow
and cache tests, and offline end-to-end DOCX round trips that verify styles,
numbering, settings, headers/footers, page layout, tables, text boxes, source
immutability, partial-failure isolation, and package validity.

## Copyright Notice

**Copyright 2025 Abraham Borg. All Rights Reserved.**

This software and associated documentation files are proprietary. Unauthorized
copying, modification, distribution, or use is prohibited without the copyright
holder's express written permission.
