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
architect DOCX  (only for format_only and csi_to_canadian)
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
       csi_to_canadian_standalone: the same conversion onto the built-in
         CSC list, with no architect and no shell application
       canadian_to_csi: the inverse, writing typed CSI markers
  -> collision-safe style import and full architect shell application
  -> mode-specific content, numbering, structure, and package invariants
  -> atomic DOCX publication in one timestamped run directory

run
  -> per-target audit JSON
  -> redacted run.log
  -> structured diagnostics.jsonl (phase timings and counts)
  -> run.json written after target results and audits
```

Targets are independent: one target can fail without discarding other validated
outputs. A failed target never publishes a partial DOCX. A run directory and
manifest still record partial or total failure.

## Repository map

```text
spec_formatter/pipeline.py
    canonical public orchestration, profile cache, isolated runs, manifests
spec_formatter/diagnostics.py
    thread-safe, redaction-safe structured diagnostics recorder and rollup
spec_formatter/llm_usage.py
    the one observed-usage contract for both classifiers, including failures
spec_formatter/resources.py
    one root for shipped prompts and notices (sys._MEIPASS when frozen)
spec_formatter/builtin_scheme.py
    the committed CSC PageFormat numbering, styles, and role contracts used
    when a run has no architect template
spec_formatter/template_analysis.py
    namespaced facade over architect analysis and bundle validation
spec_formatter/style_application/
    target extraction, shell application, numbering/header import, invariants
spec_formatter/style_application/core/conversion_modes.py
    the closed set of conversion-mode strings and their validator
spec_formatter/style_application/core/application_policy.py
    immutable mode-dependent mutation contract
spec_formatter/style_application/core/marker_tools.py
    leading-marker machinery shared by both hierarchy converters
spec_formatter/style_application/core/canadian_to_csi.py
    Canadian -> typed CSI marker conversion
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
tests/fixtures/
    example classifier instructions and the sanitized format-only corpus
tests/test_sanitized_format_only_corpus.py
    offline realistic-corpus regression against this repository's engine
tests/test_builtin_scheme.py
    proves the built-in scheme passes the unmodified architect validators
tests/test_canadian_to_csi.py, tests/test_architect_free_modes.py
    the reverse converter and both architect-free modes end to end
```

`phase1_pipeline.run_phase1()` remains a compatibility and internal profile
builder surface. New integrations must call the unified public API. The
architect analysis is observational only: there is no surface that writes
classifications back into the architect package, and none may be added.

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
- `csi_to_canadian_standalone` runs that same conversion against the built-in
  CSC PageFormat scheme instead of an architect template.
- `canadian_to_csi` is the inverse: it resolves each paragraph's current
  Canadian number and writes it into the text as a literal CSI marker.
- The two architect modes apply the architect's complete shell. The two
  architect-free modes apply **no** shell at all.

`requires_architect_template` and `numbering_scheme` on the policy are the one
place that says which of those a mode is. The pipeline must not accept an
architect template for a mode that does not use one: it fails with
`input_architect_not_accepted` rather than ignoring the selection, because a
user who chose a template and watched a run succeed would reasonably believe
it had been applied. `accepted_output_suffixes` is the matching concession in
the other direction -- `canadian_to_csi` legitimately consumes this
application's own `_CANADIAN.docx` and `_CANADIAN_FORMATTED.docx` output, so
those must not be refused as "already formatted".

What Format-only preserves is **semantic, not byte-level**, and the guarantee
is scoped to the **body**. `_verify_format_only_body_invariants` compares
paragraph blocks from `word/document.xml` before and after; that is where
"unchanged text and numbering" is proven and where the claim stops. The DOCX
package is not byte-identical and is not meant to be -- styles are imported,
the shell is applied, and parts are re-serialized. Header and footer wording
is deliberately outside the promise: `import_headers_footers` removes the
target's parts and writes the architect's, so target-authored header/footer
text is expected to change in both modes. Do not describe or test Format-only
as package byte identity, and do not describe it as preserving every word in
the file.

The same distinction applies to ignored paragraphs. Leaving a paragraph's XML
unedited proves the engine did not touch it; it does not prove the paragraph
still *renders* the same. The architect shell is document-global, so document
defaults, theme, and page geometry can reflow untouched content. That is
expected behaviour, not a preservation failure -- and it is why the change
checklist asks for visual inspection of representative output when shell or
formatting behaviour changes, rather than treating XML invariants as
sufficient on their own.

### 3. Explicit disposition coverage

Every visible classifiable target paragraph must occur exactly once in one of:

- `classifications`: styled CSI content
- `ignored_paragraphs`: non-CSI/editorial content with a non-empty reason

The sets are disjoint and cover the complete classifiable universe. Empty
structural paragraphs, table paragraphs, drawings, and text boxes are recorded
out of scope. Ignored paragraphs receive no paragraph or run edits. Missing,
duplicate, overlapping, or unknown dispositions fail closed; never restore
nearest-neighbor fallback.

The shared application path (`_apply_classified_target_impl`) re-verifies this
coverage itself, as its first stage (`disposition_verification`), against the
same role list the bundle was built with. It does not trust that a caller
coerced its payload, and the audit's `unresolved` count is never clamped: a
negative value exposes an over-full payload instead of hiding it as zero, and
any non-zero value fails the target before application begins.

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

Text-only PART/ARTICLE hits are deterministic only when the rest of the line
is shaped like a heading: empty, or starting with a capital letter or digit,
not ending like a sentence, and at most about twelve words. `1.01 SUMMARY`,
`PART 1 - GENERAL`, and `1.1 General requirements` qualify; `1.5 times the
pipe diameter shall be maintained.` and `PART 1 of the Contract Documents
shall govern.` are left for the model. A section header whose remainder
names a second `SECTION <number>` is a cross-reference, not a SectionID. A
typed single-letter marker `i.`/`v.`/`x.` (any case, any of the `x.`, `x)`,
`(x)` styles) may be a roman numeral, so it is deterministic only when the
previous paragraph carries the preceding letter in the same style. A
deterministic text-only classification cannot be overridden by the model, so
when in doubt the heuristics return nothing.

**Resolving a paragraph locally is not the same as resolving it correctly.**
A higher deterministic rate lowers cost and latency and removes a source of
variance; it says nothing on its own about precision. Because a deterministic
result cannot be overridden by the model, a wrong rule is *more* damaging than
a wrong model answer, not less. Every rule therefore needs adversarial
coverage of its close negatives -- cross-references, heading-shaped
requirement sentences, Roman/alpha ambiguity, template-specific numbering
conflicts -- and the metric for adding one is correct resolution with no
known precision regression, never a lower unresolved percentage.

### 5. Architect formatting is source-derived and collision-safe

The LLM selects roles/dispositions, not XML formatting. Role styles come from
validated architect exemplars. Resolve paragraph formatting through the full
architect `basedOn` chain. In Format-only, remove a target direct property only
when the effective architect style supplies that property. Canadian conversion
is broader by design: it replaces `jc`, `ind`, `spacing`, and `numPr` on every
converted paragraph because it retargets each one to the architect's list-level
indents, and a surviving target indent would fight the imported numbering. In
both modes, never remove `sectPr`, tracked changes, or protected subtrees as
generic formatting cleanup, and never remove numbering in Format-only.

Never replace an existing target style ID, including `Normal`. Clone a
conflicting architect style and its dependencies under deterministic private
IDs, rewrite `basedOn`/`next`/`link` and imported header/footer references, and
reject a deterministic namespace collision with different content.

### 6. Profile bundle and cache are strict boundaries

Target application consumes the complete `.phase1` directory after strict
manifest validation; loose registries are not a valid handoff. Cache profiles
under a versioned contract namespace and require exact source hash, producer,
engine fingerprint, classifier, model, and prompt compatibility. Bump the
profile contract when a consumer-visible bundle assumption changes.

The engine fingerprint (`engine_identity.ENGINE_SOURCE_DIGEST`) is a committed
digest over the files that shape a profile (`arch_env_extractor.py`,
`docx_decomposer.py`, `llm_classifier.py`, `paragraph_rules.py`,
`phase1_validator.py`). It is recorded in the manifest as
`producer.engine_fingerprint` and compared on every cache lookup, so a change
to repair logic, text-signal rules, or shell capture invalidates cached
profiles without anyone remembering to bump `PIPELINE_VERSION`. Runtime
hashing cannot work in the frozen build, so `tests/test_engine_identity.py`
recomputes the digest from the checkout and fails until the constant is
updated (`python engine_identity.py` prints the new value).

After a fresh profile is published, older profiles of the same template beyond
the newest two are removed from the cache namespace; the selected profile is
never removed and only counts are logged.

Architect analysis remains observational: derive generated styles in
`portable_styles.xml`; preserve byte-exact `source_styles.xml` and optional
`source_settings.xml`; never retag or publish a normalized architect DOCX.

### 7. Isolated, atomic, auditable publication

Create one `<UTC timestamp>_<mode>_<run-id>` directory per validated invocation,
before template analysis begins, so template/profile initialization failures
still publish a failed manifest, run log, diagnostics stream, and per-target
not-started audits. Stage each target under a short generated system-temp path,
validate the complete DOCX, and atomically publish it into that run directory.
Then atomically write per-target audits, `run.log`, `diagnostics.jsonl`, and
finally `run.json`. Never put secrets or document text in any of these,
including diagnostics fields. Existing run directories and flat legacy outputs
are immutable history.

Partial files for those atomic publications live in `<run_dir>/.staging/`
(same filesystem, so `os.replace` stays atomic) and the run removes that
directory when it ends, on success or failure; a hard kill can leave it
behind, but nothing ever sweeps other run directories. If writing the run
artifacts themselves fails after the DOCX files are published, the run still
writes a failed `run.json` with `failure_phase: "publication"` and truthful
per-target outcomes, and the raised error carries `run_dir` and
`manifest_path`. Profile provenance for `run.json` is captured on
`TemplateProfile.provenance` when the profile is selected, not by
re-validating the bundle after the outputs are already published.

## Runs without an architect template

`csi_to_canadian_standalone` and `canadian_to_csi` take a target and nothing
else. Two things make that safe rather than merely convenient.

**The built-in scheme is validated, not trusted.** `spec_formatter/builtin_scheme.py`
generates a nine-level CSC PageFormat list (`PART %1`, `%1.%2`, `.%3` ... `.%9`,
every level `decimal`, starting at 1, no `lvlRestart`) plus the twelve
`CSI_*__ARCH` role styles and their role contracts, all from constants in that
file. `tests/test_builtin_scheme.py` then runs those through the *unmodified*
`_validate_canadian_role_contract`, `_validate_complete_article_hierarchy`, and
`_validate_architect_numbering` from `core/csi_to_canadian.py` -- the same
gate a real architect template must pass. Do not add a parallel, more
forgiving check for the built-in scheme's benefit; if the scheme cannot pass
the architect's contract, the scheme is wrong.

**The numbering is applied directly, never through a style.** The built-in
role contracts declare `direct_numpr`, so `apply_phase2_classifications` gives
a classified paragraph `w:numPr` in its own `w:pPr` and leaves its `pStyle`
alone. `ApplicationPolicy.applies_role_styles` is the switch, and it is False
for both architect-free modes.

That is not a detail. Swapping a paragraph's style for a generated one that
carries no `rPr` silently flattens whatever its own style supplied -- a heading
that was Cambria bold 14pt falls back to the document defaults -- which
contradicts the mode's whole promise that only numbering and hierarchy change.
Per-level indents therefore live in `numbering.xml`, not in the styles, so the
hierarchy still reads correctly. The generated stylesheet remains as internal
scaffolding for the role contract, the registry cross-checks and the numbering
import plan; nothing from it is imported into a target.

**It is not dressed up as a `.phase1` bundle.** A bundle manifest exists to
prove an analyzed artifact on disk was not altered between analysis and use. A
scheme generated in-process from committed constants has no such gap -- no API
call, no cache, no disk round-trip -- and giving it a manifest would mean
inventing a `producer.classifier`, prompt hashes, and an engine fingerprint for
work that never ran, which makes every manifest mean less. Instead
`batch_runner.builtin_shared_config()` is a second **constructor** for the
existing `SharedConfig`. That is not a second application path:
`process_single_file()` is still the one entry point a target reaches and
`_apply_classified_target()` the one shared path beneath it. `SharedConfig.arch_root`
is `None` for these runs and `builtin_scheme` is True.

`preflight_validate_registries(..., applies_shell=False)` skips the two
shell-only checks (page layout and the header/footer contract) because there is
no shell to check. That narrows the check for one caller; it does not weaken it
for anyone else, and a test asserts the page-layout check still fires by
default.

`run.json` records `numbering_scheme` (`architect`, `builtin_csc`, or
`typed_csi`) alongside `architect_template` and `builtin_scheme`. All three keys
are always present, with nulls rather than omissions, so an absent template can
never be mistaken for one that simply was not recorded. Only a mode that
actually renders from the built-in list quotes its digest: `canadian_to_csi`
writes literal markers and imports no numbering, so its `builtin_scheme` is
null.

### Canadian to CSI

`core/canadian_to_csi.py` is the inverse of `core/csi_to_canadian.py`, and the
asymmetry matters: **writing a number is a stronger claim than removing one.**
Removing a typed marker needs only the marker; writing one needs the counter
Word would have rendered, and a wrong number becomes literal text in a document
a reader will trust.

The counter walk is therefore only performed inside the same fence the forward
converter already builds: every converted paragraph on one list instance, that
instance starting at 1 with no `lvlRestart` and no level override, and every
paragraph on it converted. Under those conditions a counter is a plain
per-level tally -- increment this level, delete the deeper ones -- and anything
outside them fails closed. A numbered role with no number in the source is
preserved unchanged and reported as a warning; no marker is invented for it.

**The classified role must agree with the level Word is rendering.** The walk
is keyed on `ROLE_LEVEL[role]`, so a paragraph classified `ARTICLE` while
sitting at `ilvl` 0 would be written `1.1` when the document actually shows
`PART 2` -- a number it never displayed, committed as permanent text. Both
converters therefore share `_validate_automatic_source` (in `marker_tools`,
with the error code parameterised), and the reverse converter additionally
requires `ilvl == ROLE_LEVEL[role]`. Every document this application's Canadian
conversion produces satisfies that, because
`_validate_complete_article_hierarchy` already requires it of the architect.

**A paragraph whose own mark is an unresolved revision has no single number.**
Reject an inserted paragraph mark and the paragraph disappears; accept a
deleted one and it merges into the next. Automatic numbering renumbers itself
either way, which is exactly the safety net a literal marker gives up, so every
marker after such a paragraph would silently become wrong the moment somebody
resolved the revision. `_paragraph_mark_revision` fails those closed with
`canadian_to_csi_tracked_hierarchy`. It is scoped to `w:pPr/w:rPr`: a `w:ins`
anywhere else marks inserted *text*, which is the ordinary state of a spec
under review and changes no paragraph's position.

Two placement rules the automatic branch must keep. The marker may not be
written inside a field result or a tracked insertion: the numbering
suppression sits on `w:pPr`, outside any such subtree, so updating the field or
rejecting the revision would delete the marker and leave the paragraph with no
number at all. And `w:numPr` goes after `w:pStyle`, because `CT_PPr` is a
sequence -- the reverse order is invalid OOXML even where Word tolerates it.
`_PPR_CHILD_ORDER` is the complete 36-element sequence from the ISO schema and
`_ppr_insertion_point` is the one way anything is added to a `w:pPr`. Do not
re-derive a shorter table for a particular element: a list abbreviated for
`w:numPr` (#7) places `w:ind` (#23) immediately after `w:pStyle`, which is
invalid, and only an XSD check notices.

**Cancelling numbering cancels the level's geometry with it.** A numbering
level's `w:pPr` -- its `w:ind` and its `w:tabs` num stop -- applies only while
the paragraph is a list member, so `numId=0` takes the indentation too whenever
it lived in `numbering.xml` rather than in the style. That is the ordinary
shape of a Canadian PageFormat stylesheet: list styles carry `w:numPr` and no
`w:ind`, so *every* indent in the document comes from the level and the whole
outline flattens into one column while text, numbers and run structure stay
provably intact. `_suppress_automatic_numbering` therefore materializes the
level's geometry onto the paragraph in the same edit, which is what Word writes
when a user turns numbering off by hand.

Restoration is narrow on purpose. Precedence between a style's `w:ind` and its
numbering level's is genuinely unsettled, so `_restorable_level_geometry`
restores only what nothing else could have supplied -- no direct `w:ind` on the
paragraph and none anywhere in its effective style chain. Guessing which source
Word preferred would risk changing a rendering in order to protect it. The
differential geometry invariant catches the residue.

**Markers are tracked when the source is.** A working spec normally has
`<w:trackRevisions/>` on and carries the author's own pending edits. Writing
numbers into it as plain accepted text puts the application's work beyond the
review every other change in the file is subject to. Where `settings.xml` says
edits are tracked, each marker is written as `w:ins` authored
`MARKER_REVISION_AUTHOR` (`"Specification Formatter"`) -- deliberately not the
document author, because the invariant projects the application's revisions out
by author, Word's markup pane should separate a machine conversion from a
person's edits, and a check that nothing altered the author's content outside a
revision must not pass trivially. `source_tracks_revisions`, `markers_tracked`
and `marker_author` are always recorded, so "these markers are plain text" is a
visible decision rather than an absent field.

**The conversion is asserted against a prediction, not described afterwards.**
The counter walk runs to completion before any paragraph is touched, and
`_verify_prediction` then checks the assembled document against that list --
read back out of the XML, so it cannot pass by agreeing with the code that
produced it. It tests what each paragraph *leads with* rather than whether it
changed, because a typed Canadian `PART 1` converts to a CSI `PART 1` and
correctly changes nothing; and it separately proves no unpredicted paragraph
changed text, which the per-paragraph check inside the edit loop structurally
cannot see. Fails closed with `conversion_prediction_mismatch`.

Markers are `PART 1`, `1.1`, `A.`, `1.`, `a.`, `1)`, `a)`, `(1)`, `(a)`. An
alphabetic level that runs past `z` fails closed rather than writing `aa.`,
because the shared `_ROLE_MARKERS` tables only ever match a single letter, so
`aa.` would produce a document this application could not read back.

An **untracked** marker is inserted **into the paragraph's existing first text
run**, not as new runs. It then inherits that run's character formatting (a
bold heading gets a bold number), and the paragraph's run structure is
unchanged, which is what the run-property invariant in `phase2_invariants.py`
checks. Adding runs would trip that invariant for a change that loses no
formatting at all; the answer is not to widen the invariant.

A **tracked** marker cannot do that -- a revision is a subtree, so it must be
its own run inside `w:ins`, which shifts every later run index and sibling
position. The answer is still not to widen the invariant: `phase2_invariants`
runs the unchanged check against the output with this application's own
revisions projected back out (`_without_own_revisions`), where the run
structure is identical to the source again. The projection is scoped by author,
so a reviewer's pending edits stay in the comparison -- losing run formatting
inside one of those is as damaging as losing it anywhere else.

Note that a round trip is not byte-exact through `PART`: the forward converter
treats a dash or colon after `PART n` as part of the typed marker and removes
it with the marker, so `PART 1 - GENERAL` returns as `PART 1 GENERAL`. That is
existing `csi_to_canadian` behaviour, not a loss introduced coming back, and
requirement text itself is untouched in both directions.

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

`phase1_bundle_manifest.json` identifies format `spec-template-normalizer.phase1`, manifest version 2, bundle ID, UTC creation time, producer/run/classifier identity, the engine fingerprint, prompt hashes, source filename/hash/size, required artifact IDs, and each artifact's path/media type/hash/size/source kind. Version 2 made `producer.engine_fingerprint` required; the formal contract is `schemas/phase1_bundle_manifest.v2.schema.json`.

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

Role expectations come from text signals and effective Word numbering, including numbering inherited through paragraph styles. Do not treat arbitrary `A.`, `1.`, or `a.` text globally as proof of CSI hierarchy. Section numbers (`230500`, `23 05 00`, `23 0500`, `23 05 00.13`) are recognised by the one grammar in `spec_formatter/style_application/core/section_numbers.py`; the classifier, the target token extractor, and the header/footer token patcher must all consume it rather than carrying their own regex. Validate exemplars, role/style coherence, numbering family/level coverage, style inheritance, and style references against the source catalogs.

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

`output_suffix` is one of those fields: the engine's staged output name in
`_build_and_patch_output` and the pipeline's planned output names in
`_plan_output_paths` both read it (`_FORMATTED.docx` for Format-only,
`_CANADIAN_FORMATTED.docx` for Canadian conversion), so the two can never
disagree. `_PHASE2_FORMATTED.docx` is a retired engine-only name that folder
discovery still excludes as legacy output.

### `spec_formatter/style_application/batch_runner.py`

- Loads one validated profile and prepares/classifies targets.
- `process_single_file()` is the one entry point a target reaches, and
  `_apply_classified_target()` is the shared application path underneath it; do
  not duplicate environment/numbering/style sequencing. The prepared-file path
  (`PreparedFile` / `_prepare_file_for_batch` / `_apply_batch_result`) was
  removed with the Batch API classifier it existed to serve: it took an
  already-classified payload as an argument, so it had no observed usage to
  report and could not satisfy the accounting contract below. Do not reintroduce
  a second entry point without carrying `usage` on **both** of its returns.
- Captures target styles/numbering before shell mutation, applies the selected
  policy, produces audit/numbering checks, validates, and packages the result.
  Both returns of `process_single_file()` carry `usage`: a successful target
  costs as much as a failed one, and `DiagnosticsRecorder.record_usage` ignores
  an empty snapshot, so omitting it on either branch silently publishes a run
  with no target usage at all.
- `load_and_validate_shared_config()` accepts only a complete `.phase1` bundle.
  The retired `run_batch_concurrent` / `run_batch_api` entry points, the
  Anthropic Batch API classifier, and the `allow_legacy_bundle` opt-in (which
  referenced the retired `arch_styles_raw.xml`) were removed; the pipeline's
  thread pool is the one *target* concurrency implementation. It is not the
  only pool in the process -- see "Concurrency, retries, and caches".

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
- `core/section_mapping.py` resolves Word's per-type header/footer inheritance
  before comparing the architect's effective section shells. Raw profile
  registries stay immutable, and explicit shell conflicts fail in shared
  profile preflight before target work begins.
- `header_footer_importer.py` and `numbering_importer.py` import bounded
  dependency sets and remap relationships/IDs. Header/footer section metadata
  substitution is restricted to corroborated section/division/title/filename
  slots in just-imported parts, including mirrored DrawingML/VML text boxes;
  ambiguous shells or incomplete target tokens fail closed.
- `phase2_invariants.py` verifies body, numbering, protected structure,
  section, header/footer, relationship, and package contracts, plus
  **effective paragraph geometry**. That last one covers the class every
  other check is blind to: identical text, identical numbering semantics,
  identical run structure and valid XSD, with every paragraph rendering
  somewhere else. It is reachable because a numbering level's `w:pPr` applies
  only while the paragraph is a list member, so cancelling the list drops the
  level's `w:ind` -- and a stylesheet whose list styles carry `w:numPr` and no
  `w:ind` keeps *all* its geometry there.

  The check is **differential, not absolute**. OOXML precedence between a
  paragraph style's `w:ind` and the `w:ind` of the numbering level it
  references is genuinely unsettled: the spec's style hierarchy applies
  paragraph styles after numbering, Word's observed behaviour for a directly
  referenced list is the reverse. So geometry is resolved under *both*
  readings and a paragraph fails only when it changed under both. A rendering
  stable under either precedence is stable whichever one Word implements, and
  a document that moved under only one is exactly where this module cannot
  honestly claim a defect. Do not "fix" it by picking a precedence.

  It runs for every mode. The two forward Canadian modes retarget converted
  paragraphs onto a different list's level indents, so their geometry is meant
  to change; they say so with `ApplicationPolicy.reindents_converted_paragraphs`
  rather than being exempted at the check.
- `docx_decomposer.py` extracts targets safely; `docx_patch.py` assembles and
  validates replacements before publication.

### Architect profile modules

`phase1_pipeline.py` snapshots and analyzes the architect. `phase1_bundle.py`
creates and validates the complete bundle. Root `docx_decomposer.py` builds the
architect slim bundle, derives portable styles without changing the source,
and emits role metadata; it reuses the package's namespace constants,
visible-text extraction, structural element scanner, UTF-8 text helpers,
and package extraction loop instead of carrying its own copies (there is no
root `ooxml_text.py`). `llm_classifier.py`, `paragraph_rules.py`,
`arch_env_extractor.py`, and `phase1_validator.py` own architect classification,
signals, shell capture, and cross-contract validation respectively.

`llm_classifier.classify_document()` applies these deterministic repairs to
the model's instructions before validation, in this order: known editorial
exclusions become ignored paragraphs; roles proven by strong text signals are
added when omitted; role exemplars whose text signals a different role are
corrected; `apply_pStyle` entries contradicted by strong signals are
corrected; and every created style's `basedOn` is set to its exemplar's
source `pStyle` (a note records each such repair, and the classification
audit embeds the notes). The model never decides `basedOn`. The 150,000-token
input cap is a cost guard measured with the API's token counter (estimate
fallback), not a context-window limit; a response that stops at the
output-token limit enters the bounded regeneration loop, while a refusal is
terminal.

### `gui.py`

Owns input collection, background execution, immutable active-run display,
progress/log rendering, and final status. It calls `format_specifications()` and
displays `FormatRunResult`. Lock every run-affecting control while work is
active, display all target processor log lines and audit counts, and open the
actual `run_dir`. Never recreate pipeline business logic in the GUI.

The GUI exposes no template-reuse or worker-count knobs: runs use
`DEFAULT_REUSE_TEMPLATE_ANALYSIS` and `DEFAULT_MAX_WORKERS` (`FormatWorker`
keeps its parameters for headless callers). The worker strips the API key once
so the pipeline and the error redaction see the same string. The target
preview passes the architect as `exclude_discovered` and re-renders when the
architect changes, mirroring the pipeline's folder discovery. A keyring save
that fails unchecks "Remember" and shows `KEYRING_UNAVAILABLE_STATUS` instead
of silently pretending the key was stored. The window opens at 980x930 with an
820x720 minimum.

## Untrusted input and limits

DOCX input and relationship metadata are untrusted.

- Reject absolute, traversal, duplicate, and symbolic-link package members.
- Limits: 10,000 package entries; 512 MiB total uncompressed; 128 MiB per part; compression ratio at most 1,000.
- The bounded, containment-checked ZIP loop exists once:
  `spec_formatter/style_application/docx_decomposer.extract_package_members()`.
  Root `docx_decomposer.extract_docx()` calls it for the architect; the
  limits above are module constants there and nowhere else, so a test that
  lowers a limit patches that module.
- Parse and validate relationship parts. Resolve internal targets only inside the package root.
- Never dereference external relationship targets, local paths, UNC paths, URLs, or encoded traversal.
- Reject malformed relationship XML and broken required relationship metadata.
- Reject any `DOCTYPE` or `ENTITY` declaration in an XML part before it reaches
  the parser: `xml.etree.ElementTree` expands internal entities, so every
  untrusted part (document, styles, numbering, relationships, content types,
  headers/footers, and the engine's own rewritten parts) is parsed only
  through `core/untrusted_xml.parse_untrusted_xml()`, which also wraps parse
  errors with the part name. Never call `ET.fromstring` on package bytes
  directly.
- That rejection is encoding-independent, in three steps over one immutable
  byte payload, so what is screened is always what is parsed. A `str` is
  already decoded, so its declaration is made truthful with
  `prepare_xml_text_for_utf8` before encoding; skipping that reads UTF-8
  bytes back through a stale declared encoding, silently turning `é` into
  `Ã©` under `windows-1252` and failing outright under `utf-16`. A byte scan
  then rejects `<!DOCTYPE`/`<!ENTITY` anywhere in the payload -- deliberately
  broader than XML requires, since it also catches declaration-shaped text in
  comments and CDATA, which is the long-standing contract and must not be
  relaxed as redundant. Finally expat's `StartDoctypeDeclHandler` rejects a
  real declaration in any encoding it can auto-detect, which the ASCII byte
  scan cannot: UTF-16 encodes `<!DOCTYPE` as `<\x00!\x00D\x00...`.
  Keep the byte scan and the expat pass together; either alone has a hole.
- The expat pass judges declarations only. When it finds a payload malformed
  it stays silent and lets `ElementTree` parse the same bytes and word the
  error, so a caller never sees a message that depends on which parser
  noticed first. An encoding expat cannot use (a codec Python lacks, or a
  multi-byte one it refuses internally) becomes `UntrustedXmlError` with the
  part name rather than a bare `LookupError` or `ValueError`.
- Header/footer media limits: 16 MiB per asset and 64 MiB total. They are
  enforced in shared-profile preflight and again at the write site in
  `header_footer_importer._write_hf_parts`; only the `data_base64` payload
  key is accepted, decoded with strict base64 validation.
- Preserve content types from `[Content_Types].xml` where available. Content
  types and document relationships are wired by parsing the part (matching
  `Override` part names case-insensitively and relationships by Type URI) and
  re-serializing it, never by string insertion before a closing tag. The final
  package validator rejects a second theme, settings, numbering, styles, or
  fontTable relationship from the main document part.

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

## Error codes and stages

Engine failures reach `run.json`, every `audit.json`, and the GUI as a stable
`error_code` plus a fixed remediation sentence, never as the raw exception
text (which can echo document text and is redacted to a fingerprint).
`spec_formatter/style_application/core/errors.py` owns the closed code set
in `ERROR_REMEDIATIONS`; raise `EngineError(code, detail)` (a `ValueError`,
so `str(exc)` keeps the developer detail and existing handlers still catch
it) or `attach_engine_error(exc, code)` when the exception type must stay
(an `ImportError`, say). `ApplicationStageError` forwards the code from its
cause, `BatchResult.error_code`/`safe_error` and `TargetFormatResult.error_code`
carry it, and `_write_run_artifacts` prefers it over classifying `error`
text. Adding a code is a contract change: list it here.

The developer detail follows two conventions so a failure can be found in
Word. Canadian architect-contract failures start with `Architect template:`
and name the role. Target-side Canadian messages name the paragraph index
and append a locator built by `_paragraph_locator()` in
`core/csi_to_canadian.py`: `(Section 21 13 13, heading 5)` is the number on
the nearest preceding SectionID paragraph and the paragraph's ordinal among
the PART and numbered-role headings after that SECTION line (`after heading
5` for a non-heading paragraph, `before any SECTION line` when none
precedes it). Locators carry section numbers and counts only, never body
text. Run-property invariant failures in `phase2_invariants.py` likewise
report property names and counts, never XML.

Current codes: `header_footer_target_section_id_required`,
`header_footer_target_section_title_required`, `header_footer_token_residual`,
`canadian_architect_contract`, `canadian_target_hierarchy`,
`canadian_target_markup`, `canadian_numbering_unprovable`,
`canadian_to_csi_hierarchy`, `canadian_to_csi_numbering_unprovable`,
`canadian_to_csi_tracked_hierarchy`, `geometry_not_preserved`,
`conversion_prediction_mismatch`, `builtin_scheme_contract`, `classification_invalid_payload`,
`classification_deterministic_override`,
`classification_coverage_incomplete`, `numbering_importer_unavailable`,
`template_section_shell_conflict`, `template_default_section_conflict`,
`template_duplicate_section_index`.

`stage` is public on `BatchResult`, `TargetFormatResult`, `audit.json`, and
`run.json`: the last checkpoint reached. The sets are closed and tested
(`ENGINE_STAGES`, `RUNNER_STAGES`, `PIPELINE_STAGES` in `core/errors.py`):

- engine (shared application path, in order): `classification_ready`,
  `disposition_verification`, `application_policy`,
  `classification_checkpoint`, `source_catalog_snapshot`,
  `target_token_extraction`, `csi_conversion`, `canadian_to_csi_conversion`,
  `canadian_classification_mapping`, `environment_application`,
  `header_footer_token_patch`, `numbering_import`,
  `header_footer_numbering_remap`, `style_import`,
  `header_footer_style_remap`, `stability_snapshot`,
  `classification_application`, `stability_verification`,
  `geometry_verification`, `application_reporting`, `output_publication`,
  `complete`
- runner (before the shared path): `validation`, `extraction`,
  `bundle_build`, `classification_preflight`, `classification`,
  `application`
- pipeline: `not_started`, `processing`, `publication`, `complete`

## Concurrency, retries, and caches

Three separate mechanisms, often conflated. Describe them precisely.

**Two pools, at different levels.** `pipeline.py` runs a thread pool over
targets: that is the one public target-processing orchestration path, and no
second one may be added. Inside a single target, `core/llm_classifier.py`
runs its own pool over the chunks of that target's slim bundle (at most six
workers). So a run with six targets can have far more than six requests in
flight, which is why the limiter below exists.

**One process-wide request limiter, target-side only.** `_REQUEST_LIMITER` in
`core/llm_classifier.py` is a `BoundedSemaphore` bounding concurrent streams
across every target and every chunk (`SPEC_FORMATTER_MAX_CONCURRENT_REQUESTS`,
default 4). It does **not** cover root `llm_classifier.py`: architect analysis
is single-threaded and one template at a time, so it never contends with
itself. Do not describe the semaphore as a global request cap.

**Two retry policies, deliberately not unified.** Both clients set
`max_retries=0` with the same timeouts, so each owns every attempt rather than
multiplying behind the SDK's hidden retries, and both fail fast on a bad key,
a bad request, or a refusal. Beyond that they differ in three ways, and a
maintainer who assumes one policy will be wrong about the others:

| | Architect (`llm_classifier.py`) | Target (`core/llm_classifier.py`) |
|---|---|---|
| Transport backoff | fixed `2 ** (attempt + 1)` sleeps, initial + 2 transient retries | `_transport_retry_delay`: honours a numeric `Retry-After` on a rate limit, else exponential; `_TRANSPORT_RETRIES = 2` |
| Structured-output compiler failure | retried **once without the schema** (`_is_structured_output_compilation_error`), and that fallback does not consume a transport retry | no equivalent; a non-transient 4xx is terminal |
| Unusable-JSON regeneration | `DEFAULT_RESPONSE_ATTEMPTS = 2` total attempts | `max_regenerations = 2`, so 3 total attempts |

Do not unify them without concrete failure evidence; a retry redesign is its
own change with its own review.

**Inherited-style lookup is already memoized.** `_style_block_index`
(`core/style_import.py`, `maxsize=16`) indexes every `w:style` block in one
structural pass, `_find_style_numpr_in_chain` (`maxsize=8192`) caches the
`basedOn` walk, and `_parsed_style_elements` (`core/classification.py`,
`maxsize=16`) parses `styles.xml` once per distinct text. These were the
engine's second-largest cost before caching. Do not add another cache here on
suspicion; profile first and show the numbers.

**Prompt caching is requested, not guaranteed.** Both classifiers mark their
system block `cache_control: ephemeral`. Zero cache reads in a run identifies
no single cause on its own: concurrent requests can all miss before the first
response returns, the prefix may be under the provider's minimum cacheable
size, the TTL may have expired, or the prefix may simply differ from the
previous run's. Concurrent misses are not by themselves a correctness bug, and
serializing requests to manufacture hits trades latency for them -- measure
before assuming that trade is worth making.

## Observed model usage

`spec_formatter/llm_usage.py` owns the one counting contract for both
classifiers. `UsageCollector` records what the provider reported and, just as
importantly, records when it could not.

- **Missing usage is unknown, not zero.** A request whose final counters never
  arrived increments `requests_with_unknown_usage` and clears `usage_complete`,
  so a total is never quietly understated. A snapshot with `usage_complete`
  false is a lower bound; the provider's invoice stays authoritative.
- A `bool` is not accepted as a counter even though it is an `int`, and a
  non-integer field is dropped rather than coerced.
- `requests_attempted` counts requests sent, `responses_completed` counts
  final messages obtained, and `responses_with_usage` counts those that
  carried recognized counters. Do not collapse these into one "requests"
  number: they answer different questions, and the difference is what makes
  an incomplete total visible.
- Record the response the moment the final message is in hand and **before**
  acting on its stop reason. A refusal and an output-limit response are both
  billed, and both used to raise with their usage unread.
- Classification failures carry their counts out on the exception
  (`attach_usage` / `usage_from_exception`), because everything that fails
  after a request -- regeneration, the overlap re-ask, the chunk merge,
  coverage, style derivation, bundle publication -- would otherwise discard
  work the run already paid for.
- Only the built-in classifier reports usage. `run_phase1` passes a collector
  solely when no classifier was injected; an injected one keeps its existing
  signature. Never probe a callable for telemetry support by calling it and
  catching `TypeError` -- that can repeat a real, paid request.
- **Totals must not depend on verbosity.** Diagnostics events are dropped
  below the configured level, so usage attached only to an INFO phase event
  vanishes at `warning`. `DiagnosticsRecorder.record_usage(scope, snapshot)`
  accumulates per-scope totals outside the event log and `summary()` publishes
  them under `diagnostics.usage` in `run.json`, identically at every level.
  Keep both: the phase event for per-phase detail, the accumulator for cost.
- `usage_complete` requires every counter in `REQUIRED_USAGE_FIELDS`
  (`input_tokens`, `output_tokens`), not merely one recognized field. A
  response reporting input but not output is partly unknown. Cache counters
  are deliberately not required, since a response that neither read nor wrote
  cache legitimately omits them. A scope's total stays incomplete once any
  contribution was incomplete.
- A deterministic-only target reports an explicit zero snapshot, never an
  absent one: "we sent nothing" and "we could not tell you" are different
  answers and must look different.
- Usage travels on its own field (`BatchResult.usage`,
  `TargetFormatResult.usage`, `TemplateProfile.usage`, `Phase1Result.usage`),
  not by reading it back out of diagnostics events. Every conversion between
  those types must carry it; a boundary that drops it silently makes the
  accounting inert without failing anything.
- A reused profile reports no architect usage for the current run. The
  analysis was paid for by the run that created it, and charging it again
  would overstate every later run.
- Usage reaches artifacts as counters only, through the same
  `sanitize_fields` boundary as every other field. Never route provider
  metadata, response text, or error strings there.

Scope: this is observation, not billing reconciliation. A request ledger,
per-request purpose enums, and cross-run rollups are deliberately not built
until measured spend justifies them.

## Run artifacts and public results

Each invocation creates:

```text
<UTC timestamp>_<mode>_<run-id>/
  *_FORMATTED.docx or *_CANADIAN_FORMATTED.docx
  target-<sequence>-<source-hash>.audit.json
  run.log
  diagnostics.jsonl
  run.json
```

`run.json` records the run/mode/status/timestamps; application, policy, and
profile contract versions; output paths; architect path/hash; cache bundle
identity; model/prompt fingerprints; target/output hashes; audit paths;
disposition counts; numbering checks; durations; a top-level `diagnostics`
rollup; and redacted errors. It must not contain API keys or document text.

`diagnostics.jsonl` is the structured detailed-diagnostics stream that
complements the human-readable `run.log`: one JSON object per phase event with
`seq`, `ts`, `level` (`DEBUG`/`INFO`/`WARNING`/`ERROR`), `component`, `event`,
optional `target`, and a `fields` object of counts/timings (per-phase
`duration_ms`, styles imported, numbering remaps, paragraphs modified, and so
on). Engine events are produced on worker threads and folded into the recorder
later, so their `ts` is an ingest time; `fields.t_ms` is the monotonic
production time (a phase's start, in milliseconds since the diagnostics clock
started), and sorting on it orders phases across targets truthfully. It is written after per-target audits and `run.log` but before `run.json`.
`spec_formatter/diagnostics.py` owns the recorder. Every field is reduced to
JSON scalars and short identifier-shaped strings by `sanitize_fields`, so a
value that could carry document text (anything with whitespace or a
document-text key such as `text`/`preview`/`content`) is dropped, never
truncated; the pipeline additionally redacts secrets from every serialized
event. Verbosity is chosen with `format_specifications(diagnostics_level=...)`
or the `SPEC_FORMATTER_DIAGNOSTICS_LEVEL` environment variable, which overrides
the argument. The `diagnostics` block in `run.json` and the per-target
`diagnostics` array in each `audit.json` obey the same boundary. Do not route
free error text, document text, or model-authored strings into a diagnostics
field; emit numbers, bools, and validated identifiers only.

`FormatRunResult` retains `success`, `succeeded`, `failed`, `output_paths`, and
the historical `output_dir`, and adds `run_id`, `conversion_mode`,
`output_root`, `run_dir`, `manifest_path`, and `diagnostics_path`.
`TargetFormatResult` retains its historical fields and adds source/output
SHA-256, `audit_path`, `audit_summary`, application audit details, numbering
checks, and structured `diagnostics` events. `BatchResult` likewise carries a
`diagnostics` list of engine phase events. Additive fields must keep safe
defaults so existing test doubles and callers continue to work.

## Development commands

```bash
pip install -r requirements-dev.txt
python -m pytest -q
python gui.py
python -m pytest tests/test_sanitized_format_only_corpus.py -q
python -m pytest tests/test_builtin_scheme.py tests/test_canadian_to_csi.py \
    tests/test_architect_free_modes.py -q
python -m pytest tests/test_geometry_invariant.py \
    tests/test_conversion_verification.py -q
```

`tests/test_conversion_verification.py` is deliberately written against the
standard library alone and shares no helper code with the engine. Keep it that
way: its value is that it can disagree with the engine's own invariants, which
is exactly what a suite written from the same mental model cannot do.

Rendered geometry is proved outside the test suite, because LibreOffice is not
available everywhere and does not implement pStyle-linked numbering levels
(a style whose `w:numPr` omits `w:ilvl` renders at level 0 there and at its
real level in Word):

```bash
python scripts/proof_render.py reference.docx candidate.docx
```

Use it before trusting a formatting change on live work. Read a non-zero
result as "look at this in Word", never as a verdict.

The GUI tests (`tests/test_gui_modes.py`) import `gui.py`, which needs
`customtkinter` and therefore a Python with `tkinter`. They skip automatically
where `tkinter` is absent (typical Linux CI images), so the rest of the suite
still collects and runs there. The Windows CI job runs everything and is the
authoritative gate.

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

The two architect-free modes pass ``None`` for the template. It stays the first
parameter and must still be passed explicitly, so every existing caller is
unaffected:

```python
result = format_specifications(
    architect_template=None,
    target_specs=[Path("21 13 13 Sprinklers.docx")],
    output_dir=Path("output"),
    api_key="...",
    conversion_mode="csi_to_canadian_standalone",  # or "canadian_to_csi"
)
```

## Change checklist

Before considering a formatter change complete:

1. Confirm architect and target sources remain unchanged.
2. Confirm every classifiable target paragraph is styled or ignored exactly once.
3. Exercise all four application policies; prove Format-only text and numbering
   are unchanged, both Canadian conversions still fail closed, and the
   architect-free modes leave the target's shell alone.
4. Confirm generated style inheritance, collision remapping, and header/footer
   style references resolve correctly without replacing target style IDs.
5. Validate the complete architect bundle and versioned cache compatibility.
6. Validate each output package, audit, `run.log`, `diagnostics.jsonl`,
   `run.json`, hashes, and redaction behavior for success, partial failure, and
   total failure.
7. Test deep Windows paths and folder discovery containing the architect.
8. Run focused tests, the complete suite, and the offline realistic-corpus
   regression in `tests/test_sanitized_format_only_corpus.py`.
9. Render and inspect every page of representative original and output DOCX
   files when formatting or shell behavior changes.
10. Update prompts, schemas, validators, README, and this guide together for
    contract changes.

## Common mistakes

- Adding any surface that writes classifications into the architect package.
- Treating the extraction directory as an output deliverable.
- Treating the selected output root as the concrete run directory.
- Writing loose formatted files or logs directly into the output root.
- Reimplementing mode checks outside `ApplicationPolicy`.
- Accepting an architect template in a mode that does not use one, instead of
  rejecting it.
- Giving the built-in scheme its own, easier validators instead of the
  architect's.
- Applying an architect shell, or any generated one, in an architect-free mode.
- Inventing a CSI marker for a paragraph the source never numbered.
- Adding runs to carry an *untracked* marker rather than joining the
  paragraph's first run. (A tracked marker must be its own run; the invariant
  is satisfied by projecting the application's revisions out, not by widening
  it.)
- Writing a marker from the classified role without proving it matches the
  list level the document actually renders.
- Placing a generated marker inside a field result or tracked insertion.
- Numbering a paragraph whose own mark is an unresolved tracked revision.
- Writing numbers into a document with `w:trackRevisions` on as plain,
  untracked text.
- Cancelling automatic numbering without materializing the geometry the
  numbering level was supplying.
- Reusing a `w:pPr` order table abbreviated for one element to insert another.
- Deriving a conversion's "expected diff" from what the conversion did.
- Writing `w:numPr` before `w:pStyle` inside `w:pPr`.
- Swapping a paragraph's own style for a generated one in a mode that promises
  to change only numbering.
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
- Running a corpus check against a sibling checkout instead of the namespaced
  implementation in this repository.

## Platform and license

Runtime is Python 3.10+ with Windows as the primary GUI platform. CI imports
the application on 3.10 (the floor) and runs the full suite on 3.11. Core
package processing is intended to remain portable.

Copyright 2025 Abraham Borg. Released under the PolyForm Noncommercial License
1.0.0 (`LICENSE`): source-available, with noncommercial use, modification, and
redistribution permitted, and commercial use requiring a separate license. This
is deliberately not an OSI-approved open source license.

Third-party dependency licenses are inventoried in `THIRD_PARTY_NOTICES.md`.
All runtime dependencies are permissive except `certifi` (MPL-2.0, file-level
copyleft, bundled unmodified).

That file is generated, never hand-edited, by
`packaging/windows/generate_third_party_notices.py`. It resolves the full
runtime closure of `requirements.txt` -- transitive dependencies included, with
environment markers evaluated for the target platform -- and reproduces each
distribution's own license file. The Windows release workflow regenerates it
from the real build environment before PyInstaller runs, so the shipped notices
always match the shipped code. Do not maintain the dependency list by hand:
`requirements.txt` pins only the direct runtime dependencies (`anthropic`,
`customtkinter`, `httpx`, `keyring`), so a hand-written list silently omits
whatever pip resolves underneath them. `requirements-dev.txt` adds the test
tools and `requirements-build.txt` the Windows build tools, including
`packaging`, which only the notices generator imports. Workflow actions are
pinned to commit SHAs that `.github/dependabot.yml` keeps current.
