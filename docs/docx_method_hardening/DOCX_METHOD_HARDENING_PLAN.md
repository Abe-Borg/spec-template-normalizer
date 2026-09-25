# DOCX Method Hardening: Implementation Plan

Program identifier: `docx_method_hardening`. Base branch: `master`.
Companion files, all in `docs/docx_method_hardening/`:

| File | Role |
|---|---|
| `DOCX_METHOD_HARDENING_PLAN.md` | This document. The specification for every work item and the rules of the program. |
| `PROGRESS_TRACKER.md` | The single source of truth for status. Machine-checked by `tests/test_docx_method_hardening_tracker.py`. |
| `HANDOFF_PROMPT_TEMPLATE.md` | The template every session fills in for the next session. |
| `handoffs/handoff-for-session-NN.md` | The prompt that starts session `NN`. Written by session `NN-1`. |
| `probes/probe_format_only_gate.py` | Reproduces the Format-only gate blind spots (WI-02). |
| `probes/probe_style_import_w14.py` | Reproduces the Format-only style-import failure (WI-01). |

## 0. Read this first

You are a coding agent with strong reasoning. This plan was written for you,
not for a human reader, so it names files, functions, line-level facts and the
evidence behind every claim. Where a fact was verified against the code on
2026-09-24 it is stated plainly; where it is an inference it says **verify**.
The code moves; when the plan and the code disagree, the code is the fact and
the plan's *intent* is the requirement. Fix the plan text in your PR when you
find such a drift, and say so in your handoff.

The program is executed over many chat sessions. Each session does exactly
one work item and opens exactly one pull request. The rituals in section 3
are not optional: they are how a session that cannot see the previous one
still knows where things stand, and how the human who owns this repository
learns, in unmistakable terms, when everything is done.

`CLAUDE.md` at the repository root is the engineering guide. Every invariant
in it applies to every work item here. Read it completely before touching
code. In particular:

- Architect and target inputs are immutable. Only private snapshots and new
  output files may change.
- Never put document text or secrets in `run.json`, `audit.json`, `run.log`
  or `diagnostics.jsonl`. Counts, booleans and identifier-shaped strings only.
- Adding an error code, a policy field, a stage, or a bundle assumption is a
  contract change: update `CLAUDE.md`, the README, the closed-set tests and
  the schemas together.
- The engine fingerprint (`engine_identity.ENGINE_SOURCE_DIGEST`) covers
  `arch_env_extractor.py`, `docx_decomposer.py`, `llm_classifier.py`,
  `paragraph_rules.py` and `phase1_validator.py`. Touch any of them and
  `tests/test_engine_identity.py` fails until you run
  `python engine_identity.py` and commit the new constant.
- The user's standing preferences: update the README when the implementation
  changes in a user-visible way, update `requirements.txt` when dependencies
  change (none of the items below should need a new dependency), and keep
  `CLAUDE.md` accurate. Do not shorten the README to make room; add only
  what is necessary.

## 1. Origin and evidence

Two working-method documents for hand-formatting fire-protection
specification files were reviewed against this engine. Most of their lessons
are already built in, several in stronger form (appendix C lists them so you
do not redo them). Three lessons exposed real gaps, confirmed by running the
engine's own code, and a handful of smaller ones followed. The disciplines
those documents insist on also shape how each item here must be built:

- **Predict first.** Write down what the change is expected to do to the
  package before running it, and make the verification assert that
  prediction. A check derived from what the transform actually did is
  confirmation, not verification.
- **Verifier independence.** A check that runs the same extraction function
  on both sides can miss a whole class of loss. `tests/test_conversion_verification.py`
  exists for exactly this reason and shares no helper code with the engine.
  Keep it that way, and extend it (WI-08) rather than bending it.
- **A failing check is root-caused, never weakened.** If a check must change,
  the replacement must be at least as strict, and the PR must say why.
- **Character truth is the XML; visual truth is a render.** Whitespace-
  normalized text is neither.

Evidence available in this repository, reproducible from the root:

```
python docs/docx_method_hardening/probes/probe_style_import_w14.py
python docs/docx_method_hardening/probes/probe_format_only_gate.py
```

Output on 2026-09-24, before any work item:

```
format_only      FAIL ValueError: Architect style contains invalid rPr XML: unbound prefix: line 1, column 123
csi_to_canadian  OK   {...}

ACCEPTED  run-level w:tab dropped beside a space
ACCEPTED  tracked-deleted text (w:delText) emptied
ACCEPTED  xml:space=preserve removed from a run ending in a space
ACCEPTED  non-breaking spaces replaced by plain spaces
ACCEPTED  double space collapsed
ACCEPTED  w:softHyphen dropped
ACCEPTED  w:br dropped beside a space
REJECTED  CONTROL: a visible word changed
7 silent mutation(s) accepted by the Format-only gate.
```

Each probe exits non-zero while its defect exists. The work item that fixes a
defect converts its probe into permanent tests and leaves the probe in place,
passing, as documentation.

## 2. Ground rules for every work item

1. **One work item per session, one pull request per session.** The PR
   contains that item's code, tests, documentation and bookkeeping, and
   nothing else. If you discover an unrelated defect, record it in the
   tracker's Notes column or in your handoff; do not fix it in the same PR.
2. **Every PR leaves `master` releasable.** A partial item may be merged only
   if what is merged is complete and validated on its own. Never merge a
   half-applied invariant or a check that is temporarily disabled.
3. **Never skip, weaken, disable or quarantine a test to get green.**
4. **Fail closed.** Where an item cannot prove a property, the target fails
   with a stable error code and a location, per `CLAUDE.md`. It does not
   publish and warn.
5. **Diagnostics carry no document text.** New audit, manifest and
   diagnostics fields are counts, booleans, or identifier-shaped strings
   (`spec_formatter/diagnostics.py` drops anything else).
6. **Invariants report that they ran.** A new check records something in
   `verification_out` on success as well as on failure, as the geometry
   invariant does, so a run can show the check happened.
7. **Reproduce first.** For a defect, add the failing test before the fix
   and show it failing in the PR description.
8. **Read the whole PR diff adversarially before pushing.** Ask what would
   make CI on Windows reject it, then fix that.
9. **Windows is the authoritative CI job** (`.github/workflows/ci.yml`
   runs the full suite on `windows-latest`; Ubuntu runs the hermetic updater
   tests, an import check on Python 3.10, and the untrusted-XML guard).
   Path handling and text encoding must work there.
10. **No new runtime dependencies.** Everything below is standard library
    plus what `requirements.txt` already pins.

## 3. How the program runs

### 3.1 Sessions, branches and pull requests

- Sessions are numbered `00`, `01`, `02`, ... Session `00` created this
  program. Session `NN` is started by the prompt in
  `handoffs/handoff-for-session-NN.md`, which session `NN-1` wrote.
- Each session works exactly one work item, in the order of the required
  table in `PROGRESS_TRACKER.md`, unless the handoff prompt or the user says
  otherwise. Optional items (section 5) are never started unless the user
  asked for them explicitly.
- Branch from `origin/master` after confirming the previous session's PR is
  merged. Use the branch name your session was assigned; if you must create
  one, name it `docx-hardening/wi-NN`. Open the PR against `master`, ready
  for review, never as a draft. There is no PR template in the repository;
  write the body as: what changed, the evidence (failing test before, passing
  after, probe output), the verification you ran, and what the next session
  should do.
- After opening the PR, subscribe to its activity and drive CI to green.
  Review comments are yours to address in the same PR.

### 3.2 Session start ritual

Do these before any code change, in this order:

1. Read `CLAUDE.md`, this plan, `PROGRESS_TRACKER.md`, and your handoff
   prompt, completely.
2. `git fetch origin master`. Confirm the previous session's PR (named in
   the tracker and the handoff) is **merged**. If it is still open, stop and
   tell the user; do not build on a stale base unless the user explicitly
   says to proceed anyway.
3. Create your branch from `origin/master`.
4. Establish the baseline: run `python -m pytest -q` once, unchanged, and
   record the result (counts, skips, failures) in your handoff. If the
   baseline is red, that is the first thing to report.
5. First commit: bookkeeping only. In `PROGRESS_TRACKER.md` set the previous
   item to `merged` with its merge commit (the SHA on `master`), set your
   item to `in_progress` with your session number, and add your session row
   to the session log with the outcome `started`. The tracker test must pass
   on this commit.

### 3.3 Session end ritual

1. Run, in this order, and record the results in the PR and the handoff:
   - the item's own new tests;
   - `python -m pytest -q` (the complete suite);
   - `python -m pytest tests/test_sanitized_format_only_corpus.py -q`;
   - `python -m pytest tests/test_engine_identity.py -q` if you touched a
     fingerprinted file;
   - both probes in `docs/docx_method_hardening/probes/`.
2. Tick every box in the item's **Definition of done** in this plan. If a
   box cannot honestly be ticked, the item is not done: see section 3.6.
3. Update `CLAUDE.md`, the README and `THIRD_PARTY_NOTICES.md` as the item
   requires. The README is user-facing; `CLAUDE.md` is the engineering
   contract; both must remain true after your change.
4. Commit, push, open the PR against `master`.
5. Set the tracker row to `in_review` with the PR URL, update your session
   log row (outcome and PR), commit, push.
6. Write `handoffs/handoff-for-session-NN+1.md` from
   `HANDOFF_PROMPT_TEMPLATE.md`, commit, push. Then paste the prompt from
   that file into chat, verbatim, so the user can start the next session
   without opening the repository.
7. Subscribe to the PR. Drive CI to green. When the PR merges, paste the
   handoff prompt into chat again with the merge line updated to
   `merged as <sha>`. Do not commit that update; the next session's start
   ritual records the merge in the tracker.
8. If your item was the last required item, also do section 3.5.

### 3.4 Progress tracking rules

`PROGRESS_TRACKER.md` has three tables and one status line, and
`tests/test_docx_method_hardening_tracker.py` enforces the following on
every commit. Read the test before editing the tracker; its assertion
messages tell you what it expects.

- The required and optional tables have the columns
  `ID | Title | Status | Session | PR | Merge commit | Notes`.
- Every `WI-NN` in the tables has a `### WI-NN:` section in this plan with a
  **Definition of done** checklist, and vice versa.
- Status values: `not_started`, `in_progress`, `in_review`, `merged`,
  `blocked`, `dropped`; optional items may also be `not_scheduled`.
- `in_progress`, `in_review` and `merged` require a two-digit session.
  `in_review` and `merged` require a GitHub PR URL and every Definition of
  done box ticked. `merged` requires the merge commit SHA. `dropped` and
  `blocked` require a non-empty Notes cell saying why.
- The session log has the columns
  `Session | Date (UTC) | Work item | Outcome | PR | Handoff written`, and
  for every session row `NN` the file `handoffs/handoff-for-session-NN+1.md`
  must exist, be named in the last column, and start with the line
  `# Handoff prompt for session NN+1`.
- The `Program status:` line reads `IN PROGRESS` until every required item
  is `merged` or `dropped`, and `PROGRAM COMPLETE` from then on. The test
  enforces both directions.

Checkboxes are used in this plan only inside Definition of done lists, so
the test can read them without ambiguity. Do not add checkboxes elsewhere.

### 3.5 Completion: tell the owner in big letters

The program is complete when every **required** item is `merged` (or
`dropped` with a recorded reason). Optional items never block completion.

The session whose PR merges the last required item does two things after
the merge event arrives:

1. Pastes the post-merge handoff prompt (section 3.3 step 7). That handoff
   assigns the next session the **closeout**: a bookkeeping-only PR that
   records the final merge in the tracker and flips `Program status:` to
   `PROGRAM COMPLETE`. The closeout session still writes a handoff file for
   the session after it; that file says the program is complete and lists
   the optional items the user may still choose.
2. Prints the banner below in chat, exactly, inside a fenced code block,
   followed by the heading line. The closeout session prints it again when
   its PR merges. Never print this banner while any required item is not
   merged.

```
 █████╗ ██╗     ██╗         ██████╗  ██████╗ ███╗   ██╗███████╗
██╔══██╗██║     ██║         ██╔══██╗██╔═══██╗████╗  ██║██╔════╝
███████║██║     ██║         ██║  ██║██║   ██║██╔██╗ ██║█████╗
██╔══██║██║     ██║         ██║  ██║██║   ██║██║╚██╗██║██╔══╝
██║  ██║███████╗███████╗    ██████╔╝╚██████╔╝██║ ╚████║███████╗
╚═╝  ╚═╝╚══════╝╚══════╝    ╚═════╝  ╚═════╝ ╚═╝  ╚═══╝╚══════╝
```

`# THE DOCX METHOD HARDENING PLAN IS 100% COMPLETE. EVERY REQUIRED WORK ITEM IS MERGED TO MASTER.`

### 3.6 When a session cannot finish, and other exceptions

- **Out of context or time before the item is done.** Push what is
  validated (green tests, no half-applied checks), leave the tracker row at
  `in_progress`, open the PR anyway with `(partial)` in the title, and write
  the handoff saying precisely what remains. The next session continues the
  same item on a new branch after that partial PR merges. Do not leave an
  unpushed branch behind; the container is ephemeral.
- **Blocked.** Set `blocked` with the reason in Notes, write the handoff,
  and tell the user what decision unblocks it.
- **Dropping an item** is the user's call, never the agent's. Record it as
  `dropped` with the user's reason only after they said so.
- **A plan error.** If an item's approach turns out to be wrong, say so in
  the PR and the handoff, fix the plan text in the same PR, and keep the
  item's *intent*. Do not silently implement something else.
- **Reordering.** Only the user or a handoff prompt may reorder items.
  WI-03 depends on WI-02 and WI-08 should run after WI-02 through WI-06;
  everything else is independent.

## 4. Required work items

Effort sizes are rough, for pacing a session: S is under half a session,
M about one session, L may need a partial PR.

### WI-00: Program scaffolding

**Session 00. This item is the PR that created these files.**

**Definition of done**

- [x] `DOCX_METHOD_HARDENING_PLAN.md` written with every required and optional item specified.
- [x] `PROGRESS_TRACKER.md` created with every item, statuses, and the session log.
- [x] `HANDOFF_PROMPT_TEMPLATE.md` created.
- [x] Both probes committed under `probes/` and runnable from the repository root.
- [x] `tests/test_docx_method_hardening_tracker.py` added and passing.
- [x] `CLAUDE.md` repository map and development commands mention the program.
- [x] Pull request opened against `master`.
- [x] `handoffs/handoff-for-session-01.md` written and pasted in chat.

### WI-01: Extension-namespace-safe style import and shell application

**Effort: M. Fixes a hard failure.**

**Why.** Every document current Word creates writes
`<w14:ligatures w14:val="standardContextual"/>` into
`docDefaults/rPrDefault/rPr`, and styles can carry `w14:` children too
(`w14:textFill`, `w14:glow`, `w14:shadow`). The bundle's portable stylesheet
carries the architect's styles verbatim (`build_portable_styles_xml` in the
root `docx_decomposer.py` only appends derived role styles and propagates
namespace declarations for those). The Format-only path then materializes
each body role style's effective run properties through the `basedOn` chain
and the document defaults, and that materialization parses every
run-property block with only the main namespace declared. The parse raises
`unbound prefix`, surfaced as
`ValueError("Architect style contains invalid rPr XML: ...")`, and the target
fails at the `style_import` stage with no error code. The Canadian modes skip
that materialization and succeed on the same template.

**Evidence.** `probes/probe_style_import_w14.py` (fails today).
`tests/fixtures/sanitized_format_only_corpus.py` contains no extension-
namespace content, which is why the corpus never caught it. No test in the
repository runs a stylesheet with `w14:` content through Format-only.

**Root causes, verified 2026-09-24.**

1. `spec_formatter/style_application/core/style_import.py`:
   `_ppr_children_by_name` and `_rpr_children_by_name` wrap the inner XML
   in `<w:pPr xmlns:w=...>` / `<w:rPr xmlns:w=...>` and call
   `ET.fromstring`. Their callers are `_effective_ppr_inner_in_arch`
   (paragraph-style pPr materialization) and
   `_effective_full_rpr_inner_in_arch`, which is called by
   `_materialize_full_rpr_for_detached_body` for Format-only body roots and
   also walks the docDefaults fallback. Both key their inheritance maps by
   *local* name, so even after parsing, a `w14:` child could shadow a `w:`
   child of the same local name.
2. Second hole, same family: when cloned blocks are written into the target
   `word/styles.xml`, and when `apply_doc_defaults`
   (`spec_formatter/style_application/arch_env_applier.py`) splices the
   registry's docDefaults into the target stylesheet, nothing reconciles
   the target root's namespace declarations. A `w14:`-bearing fragment
   written into a target whose `w:styles` root does not declare `w14`
   produces ill-formed XML. `docx_patch.validate_xml_wellformedness` then
   refuses to build the package, which is fail-closed, but it is still a
   total run failure with an unhelpful message. **Verify** by grepping
   `xmlns` in `style_import.py`: as of 2026-09-24 the only namespace
   handling there is the two ET wrappers above.

**Approach.**

1. Replace the ET round-trip in `_ppr_children_by_name` and
   `_rpr_children_by_name` with the lexical child walker that already
   exists: `iter_direct_child_xml_blocks` in
   `spec_formatter/style_application/core/xml_helpers.py` (used by
   `phase2_invariants` for exactly this purpose). Return the source
   fragments byte-for-byte. Key `resolved`/`order` by the **qualified**
   name (`w:rFonts`, `w14:ligatures`). Keep the existing exclusions
   (`pStyle`, `numPr`, `sectPr`, `pPrChange` for pPr; `rStyle`, `rPrChange`
   for rPr) by qualified name under the `w:` prefix.
   *As implemented (session 01, after review of PR 61):* keyed by the
   **expanded** name instead -- namespace URI and local name, each prefix
   resolved through the architect root's declarations -- because two
   prefixes bound to one namespace name one property, and keying by the
   prefix as written kept a parent's value beside the child's override.
2. Ordering of materialized children: keep the current first-seen order
   for `w:` children; place every extension-namespace child after all `w:`
   children, in first-seen order. Word writes extension children last in
   `w:rPr`, and the strict schema does not know them at all.
3. Decide, and record in `CLAUDE.md` invariant 5, that extension children
   are **carried**, not dropped. The method documents drop them for
   hand-transplants; here dropping would silently change formatting the
   architect specified (`w14:textFill` is visible), and the Phase 1 side
   already preserves them (`test_portable_styles_preserve_extension_properties_and_declare_namespace`
   in `tests/test_core_fix_regressions.py`).
4. Add one namespace helper module or a section of `xml_helpers.py` with:
   `root_namespace_declarations(xml_text) -> dict[prefix, uri]` (from the
   root element's open tag), `prefixes_used(fragment) -> set[str]` (element
   and attribute prefixes; ignore `xml` and `xmlns`; *as implemented, also
   the prefixes named in markup-compatibility values such as
   `mc:Choice Requires="w14"`*), and
   `ensure_root_declarations(part_xml, needed: dict[prefix, uri], ignorable: set[str]) -> str`
   which adds missing `xmlns:` declarations to the root open tag, merges
   `mc:Ignorable` tokens for prefixes the architect root lists as ignorable
   (declaring `mc` when needed), and fails closed when the target already
   binds the prefix to a different URI. Use it in
   `import_arch_styles_into_target` for every block written into the target
   stylesheet, and in `apply_doc_defaults` for the docDefaults fragment.
   The prefix-to-URI map comes from the portable stylesheet root
   (`arch_styles_xml` is already in `batch_runner`'s shared path; thread it
   into `apply_environment_to_target`, or read the bundle's
   `portable_styles.xml` root from `registry_dir`, which
   `apply_environment_to_target(target_extract_dir, registry, log, ...,
   registry_dir=None)` already accepts and passes to `apply_settings`).
   Whichever you choose, the map is built once per target and shared.
5. The prefix conflict is a new error code: `style_import_namespace_conflict`
   with a fixed remediation sentence, added to `ERROR_REMEDIATIONS` in
   `core/errors.py`, to the "Current codes" list in `CLAUDE.md`, and to the
   README's error documentation if it lists codes. Note
   `tests/style_application_regression/test_engine_errors.py` bounds the
   size of `ERROR_REMEDIATIONS` (`15 <= len <= 21` on 2026-09-24); raise the
   upper bound in the same PR.
6. Wrap the remaining `ValueError("Architect style contains invalid ...")`
   paths so a genuinely malformed fragment still fails with a message that
   names the style ID, never the XML.

**Tests (add before the fix, show them failing).**

- Unit, in `tests/style_application_regression/` next to the existing
  style-import tests: (a) docDefaults with `w14:ligatures`, Format-only body
  root import succeeds, the materialized rPr keeps the child after the `w:`
  children, and the output stylesheet root declares `w14` and parses;
  (b) a `Normal` with a `w14:` child cloned for Canadian mode into a target
  whose root declares only `w`: output declares and parses; (c) target binds
  `w14` to a different URI: `EngineError` with the new code; (d) architect
  root lists `w14` in `mc:Ignorable`, target has no `mc:Ignorable`: output
  root gains `xmlns:mc` and `mc:Ignorable="w14"`; (e) a `w:` child and a
  `w14:` child sharing a local name do not shadow each other.
- End to end through `format_specifications` with the injected
  deterministic classifier pattern from `tests/test_unified_roundtrip.py`
  (`_deterministic_classifier`), architect styles carrying ligatures in
  docDefaults and in `Normal`, target stylesheet with a bare root, mode
  `format_only`: the run succeeds, the output package passes
  `validate_docx_package`, and the corpus regression still passes.
- Convert `probes/probe_style_import_w14.py` into a test; keep the probe
  runnable and passing.

**Documentation.** `CLAUDE.md`: invariant 5 (carry extension children,
declare their namespaces, fail closed on a conflict), the
`style_import.py` module section, the error-code list, and a
"Common mistakes" bullet ("parsing an OOXML fragment with only the `w`
namespace declared"). README: the safety-guarantees list gains one line.
Engine fingerprint: unaffected unless you touch the root
`docx_decomposer.py`.

**Risks.** Switching from `ET.tostring` to raw fragments changes the
whitespace of materialized blocks (`<w:keepNext />` becomes
`<w:keepNext/>`); tests asserting exact strings will need updating, and
that is fine. Do not "normalize" the raw fragments to keep old assertions
passing.

**Definition of done**

- [x] Both rows of `probe_style_import_w14.py` (`format_only` and `csi_to_canadian`) print OK, and the probe exits 0.
- [x] Failing tests were added first and are now passing, including the end-to-end Format-only run.
- [x] Materialization keys by expanded name (namespace and local name, so neither a shared local name nor a second prefix for one namespace confuses it) and carries extension children after `w:` children.
- [x] Target stylesheet roots gain the declarations (and `mc:Ignorable` tokens) the inserted fragments need; a prefix bound to a different URI fails closed with `style_import_namespace_conflict`.
- [x] `apply_doc_defaults` is covered by the same guarantee.
- [x] `CLAUDE.md` invariant 5, module notes, error-code list and README updated; `test_engine_errors.py` bound raised.
- [x] Full suite and corpus regression green on the PR.

### WI-02: Exact run-content signature at the Format-only gate

**Effort: M.**

**Why.** The final gate for Format-only,
`_verify_format_only_body_invariants` in
`spec_formatter/style_application/phase2_invariants.py`, compares
`paragraph_text_from_block` (in `core/xml_helpers.py`) before and after.
That helper removes `w:del`, `w:moveFrom` and `w:instrText`, maps `w:tab`,
`w:br` and `w:cr` to a space, collapses all whitespace and strips. It is the
same function on both sides, so it cannot see: tracked-deleted text
changing, a run-level tab or break dropped beside a space, a lost
`xml:space="preserve"` (Word then discards the edge space on open), a
non-breaking space turned into a plain one, a collapsed double space, or a
dropped soft hyphen. The probe demonstrates all seven. The byte-level diff
contract in `apply_phase2_classifications`
(`_normalize_paragraph_for_contract`, `core/classification.py`) covers the
application step only; the final gate is what covers numbering import,
style import, shell application and repackaging, and it is what the README's
"body text comes through unchanged" claim rests on.

**Evidence.** `probes/probe_format_only_gate.py` (seven rows accepted today).

**Approach.**

1. Add `paragraph_run_content_signature(p_xml) -> tuple` to
   `core/xml_helpers.py`, next to `paragraph_text_from_block`, specified in
   appendix A. It walks every run in the paragraph after
   `strip_out_of_scope_subtrees`, including runs nested in `w:ins`, `w:del`,
   `w:moveTo`, `w:moveFrom`, `w:hyperlink`, `w:smartTag`, `w:sdt`,
   `w:customXml`, `w:fldSimple`, `w:dir` and `w:bdo`, and records each run's
   content children in order, with no whitespace normalization of any kind.
   `w:tab` counts only as a run child; the tab-stop definitions in
   `w:pPr/w:tabs` and inside `w:pPrChange` use the same element name and
   must not count. `w:lastRenderedPageBreak` and `w:rPr` are excluded.
2. In `_verify_format_only_body_invariants`, compare signatures paragraph
   by paragraph where the blocks differ (identical blocks need no
   extraction, as now). Keep the existing normalized-text comparison and
   its message `FORMAT_ONLY INVARIANT FAIL: target body text changed at
   paragraph index N` (a test matches on it); add a second failure
   `FORMAT_ONLY INVARIANT FAIL: target run content changed (<kind>) at
   paragraph index N` where `<kind>` is the item kind that differed
   (`text`, `deleted_text`, `tab`, `break`, `preserve_space`, ...), never
   the text.
3. Record `body_signature_paragraphs_compared` in `verification_out`.
4. Text decoding: compare decoded text, not raw escaped text, but decode
   with an XML-only unescape (the five predefined entities plus numeric
   references), not `html.unescape`, which also accepts HTML entities that
   are not XML.

**Tests.** Turn every row of the probe into a test in
`tests/style_application_regression/test_final_package_validation.py`
(seven rejections plus the control). Add a positive test proving a
paragraph whose only difference is a stripped contracted `w:rPr` child
still passes, so the signature does not fire on the edits Format-only
legitimately makes. Corpus regression must stay green: Format-only never
edits a text node, so a corpus failure here means the signature counted
something it must not (look first at `w:pPr/w:tabs` and at
`w:txbxContent`).

**Documentation.** `CLAUDE.md` invariant 2: replace "compares paragraph
blocks ... that is where unchanged text and numbering is proven" with a
description of the signature and what it includes. README safety
guarantees: the Format-only line now says text, tracked deletions, tabs,
breaks and space preservation are proven unchanged.

**Definition of done**

- [x] `probe_format_only_gate.py` prints REJECTED for every row and exits 0.
- [x] `paragraph_run_content_signature` implemented per appendix A with unit tests for each item kind and for the `w:pPr/w:tabs` exclusion.
- [x] The Format-only gate compares signatures and reports the kind, never text; `body_signature_paragraphs_compared` recorded on success.
- [x] Existing message contract kept; probe rows are permanent tests.
- [x] `CLAUDE.md` invariant 2 and README updated.
- [x] Full suite and corpus regression green on the PR.

### WI-03: Final-gate text identity for every mode (identity except the enumerated diff)

**Effort: M. Depends on WI-02.**

**Why.** `verify_phase2_invariants` runs the body text check only when
`policy.preserve_target_numbering` is true, so for the three conversion
modes the final gate proves no text property at all. The converters verify
their own edits in memory, at the `csi_conversion` /
`canadian_to_csi_conversion` stages, before environment application,
numbering import, style import, classification application and packaging.
Anything later that damaged text would publish. The method documents' rule
is "byte identity except an explicitly enumerated diff list, asserted
exactly"; the reverse converter already builds that list
(`ConversionPlan` and `_verify_prediction` in `core/canadian_to_csi.py`),
and the forward converter computes each converted paragraph's expected body
(`_verify_changed_paragraph` in `core/marker_tools.py`).

**Approach.**

1. Define a small frozen type, `ExpectedParagraphChanges`, holding for each
   changed paragraph index the expected visible text and, where the
   converter knows it, the expected text skeleton (forward: source skeleton
   minus the removed tab; reverse: marker plus source text). Have both
   converters return it alongside their document (find where each
   converter's result is consumed in `_apply_classified_target_impl` in
   `spec_formatter/style_application/batch_runner.py`). Verified names on
   2026-09-24: the forward converter is `plan_csi_to_canadian` /
   `apply_csi_to_canadian` in `core/csi_to_canadian.py`, with the frozen
   dataclasses `MarkerEdit`, `CanadianConversionReport` and
   `ConversionPlan`, and it calls `_verify_changed_paragraph` per converted
   paragraph inside `plan_csi_to_canadian`; the reverse converter is
   `plan_canadian_to_csi` / `apply_canadian_to_csi` in
   `core/canadian_to_csi.py` with its own `ConversionPlan`.
2. Thread it into `_build_and_patch_output` and from there into
   `verify_phase2_invariants(expected_paragraph_changes=...)`. Format-only
   passes an empty set.
3. In the gate, for every paragraph **not** in the set, require identical
   run-content signatures between source and output, after projecting out
   this application's own tracked insertions (`_without_own_revisions`,
   author `MARKER_REVISION_AUTHOR`) so tracked markers do not trip it.
   For every paragraph **in** the set, re-run the mode's per-paragraph text
   assertion against the expected values, independently of the converter's
   in-memory result. Fail with `conversion_prediction_mismatch` for the
   Canadian modes (it already exists and means exactly this) and the
   Format-only message for Format-only.
4. Record `body_signature_paragraphs_compared` and
   `body_paragraphs_expected_changed` in `verification_out`.

**Tests.** For `csi_to_canadian`, `csi_to_canadian_standalone` and
`canadian_to_csi`: inject a mutation into a non-converted paragraph after
conversion (the test doubles in `tests/test_unified_roundtrip.py` and
`tests/test_architect_free_modes.py` show how to drive each mode without an
API key) and assert the gate fails; mutate a converted paragraph beyond its
marker and assert the gate fails; the happy paths still pass and the new
counters appear in the `build_output` diagnostics event.

**Documentation.** `CLAUDE.md` "Target shell, packaging, and invariants"
gains a bullet describing the enumerated-diff gate; README safety
guarantees: "every mode proves that only the paragraphs the conversion
predicted have changed".

*As implemented (session 03):*

- **The expectation is exact run content, not a text skeleton.** A skeleton
  (the paragraph XML with `w:t` contents blanked) cannot be compared at the
  final gate: classification application legitimately rewrites `pPr` and
  `rPr` in the Canadian modes, so every skeleton would differ. Each
  `ExpectedParagraphChange` (`core/expected_changes.py`) instead holds the
  visible text, the exact run-content signature, and the signature with this
  application's own tracked insertions projected out.
- **Worked out from the source, on the signature.**
  `_predict_marker_removal` (`marker_tools`) and `_predict_marker_insertion` /
  `_predict_marked_paragraph` (`canadian_to_csi`) restate each edit on the
  source paragraph's signature, never on the XML the edit writes.
  - Each converter asserts its assembled document against the prediction
    exactly, beside its existing normalized checks, so every existing
    converter test also exercises the model.
  - A prediction that cannot be formed is reported only after the edit had
    its chance to refuse the paragraph in its own, actionable terms.
  - The exact prediction also catches a real converter defect: the forward
    edit decodes with `html.unescape`, which turns `&#x80;` into a euro sign.
    Such a paragraph now fails closed.
- **No projection for unpredicted paragraphs.** Approach step 3 projects this
  application's insertions out before the identity check. For a paragraph
  outside the prediction that would *hide* a marker inserted where none was
  predicted, so unpredicted paragraphs are compared exactly, which is
  stricter. The projection is used only for predicted paragraphs, where it
  proves a tracked marker is still inside its own revision.
- **Interfaces.** `apply_csi_to_canadian` / `apply_canadian_to_csi` return a
  `ConversionResult` (the text-free report beside the prediction), so the
  prediction never rides on a published report; `ConversionPlan` carries it
  too. The prediction also records the classified roles, so a gate failure is
  placed by SECTION and heading. An omitted prediction allows no change, and
  Format-only refuses a non-empty one.
- `MARKER_REVISION_AUTHOR` and the own-revision projection moved to
  `core/expected_changes.py`, so the converters and the gate share one
  definition without an import cycle; `canadian_to_csi` re-exports the
  constant.
- **After review of PR 63.** An unpredicted paragraph is compared by visible
  text as well as by signature, because the signature does not record a run's
  container: a run wrapped in `w:del` or `w:moveFrom` kept its signature while
  its text vanished. The counters are recorded as the body check starts, in
  both branches, so a failure before any comparison (a paragraph added or
  removed, a Format-only word changed) still shows the check ran.
- **Owner-approved extra fix, in its own commit.** With Track Changes on,
  `_insert_marker` found the run to precede by searching back for `<w:r`,
  which also matches `<w:rPr`. For every formatted run it put the `w:ins`
  inside the run, ahead of its properties: invalid OOXML that published as a
  success. The prediction asserts the documented placement, so the
  placement was fixed in this item's PR at the owner's request.

**Definition of done**

- [x] Both converters return an `ExpectedParagraphChanges` value consumed by the final gate.
- [x] Non-predicted paragraphs are proven identical by run-content signature in every mode; predicted ones are re-asserted at the gate.
- [x] Mutation tests fail the gate for all three conversion modes; happy paths pass.
- [x] Counters recorded on success; no document text in any new field.
- [x] `CLAUDE.md` and README updated.
- [x] Full suite and corpus regression green on the PR.

### WI-04: Even-page header parity follows the architect

**Effort: S to M.**

**Why.** `import_headers_footers` in
`spec_formatter/style_application/header_footer_importer.py` wires the
architect's `default`, `first` and `even` header and footer references into
every target section (`_rewire_document_sectpr`), and
`verify_phase2_invariants` proves the references are there. The switch that
makes Word render even-page headers, `w:evenAndOddHeaders` in
`word/settings.xml`, is global, and `apply_settings` in
`arch_env_applier.py` carries only the architect's `w:compat` block by
design (protection, tracking and mail-merge state stay target-owned). So an
architect with distinct odd and even headers renders in the output with its
default header on every page, and an architect without them applied to a
target whose switch is on leaves even pages with no header. Both are silent.
The string `evenAndOddHeaders` appears in no engine module on 2026-09-24;
`tests/test_unified_roundtrip.py` builds an architect with the switch and an
even header but asserts nothing about the output settings. `w:titlePg` is
per-section and already handled.

**Approach.**

1. Derive the architect's parity from the registry's
   `settings.settings_xml` (already captured, canonicalized, by
   `extract_settings` in `arch_env_extractor.py`), so the profile contract
   and the engine fingerprint are untouched. Treat
   `<w:evenAndOddHeaders w:val="0|false|off"/>` as off, mirroring
   `_TRACK_REVISIONS_RX` in `core/canadian_to_csi.py`.
2. Under `policy.apply_full_architect_shell` only, set or clear the element
   in the target's `word/settings.xml` to match the architect. Insert with
   the complete `CT_Settings` child order in appendix B (verify it against
   the ECMA-376 schema text before committing; an abbreviated table is the
   exact mistake the method documents record). Creating the minimal settings
   part when the target has none is already handled by `apply_settings`;
   reuse that path.
3. Do **not** add a preflight rejection for an architect whose sections
   reference an `even` header or footer while its switch is off. That is a
   valid and common state: Word keeps the even part dormant and renders the
   default header on even pages, and step 2 reproduces exactly that in the
   target by clearing the switch. The importer already wires the dormant
   reference as it finds it; leave that alone. No new error code is needed
   for this item.
4. Invariant: in the header/footer section of `verify_phase2_invariants`,
   when architect header/footer data is present, require the output
   settings parity to equal the architect's; record `header_parity_checked`
   in `verification_out`.

**Tests.** Extend `tests/test_unified_roundtrip.py` so the existing
even-header architect asserts `<w:evenAndOddHeaders/>` in the output
settings; add the inverse (target has the switch, architect does not: it is
removed); an architect-free mode leaves the target settings byte-identical;
an architect with a dormant even reference and the switch off imports the
reference unchanged and leaves the target switch cleared; a unit test for
the schema-order insertion at each neighbour position
(before `w:bookFoldRevPrinting`, after `w:defaultTableStyle`, into an empty
settings body).

**Documentation.** `CLAUDE.md`: the shell description in invariant 2 and
the `arch_env_applier.py` bullet. README: Format-only and
Canadian mode sections mention that even/odd header parity follows the
template.

*As implemented (session 04):*

- **The switch follows the header set, not the shell alone.** Approach step 2
  sets or clears the switch under the full shell, and step 4 checks it only
  when architect header/footer data is present. The two disagree where the
  importer keeps the target's own headers (the architect has no header parts,
  or none a mapped section references): changing the switch there would do to
  the target's headers exactly what this item fixes for the architect's. So
  the switch follows the architect exactly when the architect's set replaced
  the target's (`HeaderFooterImportResult.replaced_target_parts`, set by
  `import_headers_footers`), and is otherwise the target's own and left as
  written. `apply_header_parity` (`arch_env_applier.py`) runs after the
  header/footer import, still under `apply_full_architect_shell` only.
- **One module, `core/header_parity.py`.** `even_and_odd_headers` reads the
  switch namespace-aware and fails closed (`HeaderParityError`, a
  `ValueError`) on a repeated element or a `w:val` outside the six `ST_OnOff`
  spellings (trimmed), rather than mirroring `_TRACK_REVISIONS_RX`'s "anything
  else is on". `architect_even_and_odd_headers` reads the registry's
  `settings.settings_xml`: `None` is a template with no settings part (off); a
  registry that does not record the field is unknown, not off.
  `set_even_and_odd_headers` edits lexically and proves afterwards that the
  part reads as intended, is byte-identical to the original once every switch
  is taken out of both, and has the switch at its schema position.
- **Insertion point.** After the last child the sequence places before the
  switch (or just inside the root), rather than before the first child it
  places after it. The two agree on any part whose children are all known and
  ordered; this reading also keeps an extension element such as `w14:docId`
  from moving it. A part with no valid position (a child that must follow the
  switch sitting before one that must precede it) fails closed.
- **The table was checked against the schema text**, element by element, in
  two copies that agree: ISO/IEC 29500-4:2012 as printed (transitional
  `wml.xsd`, `CT_Settings` from schema line 2896, page 922) and the
  transitional `wml.xsd` in python-docx's `ref/xsd` (commit `e454546`). Both
  match appendix B exactly, and `tests/test_header_parity.py` holds the
  constant equal to appendix B.
- **Preflight.** No rejection for a dormant `even` reference. One check was
  added, under `applies_shell`: a template with headers or footers whose
  switch cannot be read fails once, before any target work. Two hand-built
  test registries gained `settings.settings_xml`, which Phase 1 always writes.
- **The gate** resolves each package's settings part through its document
  relationship. When the architect's set was imported the output's switch must
  read as the architect's; when the target kept its own it must be exactly as
  written, compared uninterpreted. It records `header_parity_checked`,
  `header_parity_follows_architect` and `even_and_odd_headers` as it starts;
  `apply_environment` records `header_parity_follows_architect`,
  `even_and_odd_headers` and `header_parity_changed`.
- **After review of PR 64.** The writer edited `word/settings.xml` by name
  while the gate reads the part the document relates, so a target relating
  its settings under another name failed the gate. `apply_header_parity` now
  writes into the related part (packaged by `_build_and_patch_output` when it
  is not `word/settings.xml`), creates `word/settings.xml` only when nothing
  is related, and refuses to relate a stray unrelated `word/settings.xml`.

**Definition of done**

- [x] Parity derived from the registry's settings XML; set or cleared under the full shell only.
- [x] Insertion uses the complete `CT_Settings` order table, verified against the schema.
- [x] A dormant even reference with the switch off is imported unchanged and the target switch is cleared; no preflight rejection was added.
- [x] Final gate verifies parity and records `header_parity_checked`.
- [x] Tests listed above added and passing; architect-free modes proven untouched.
- [x] `CLAUDE.md` and README updated.
- [x] Full suite and corpus regression green on the PR.

### WI-05: Package-level change whitelist invariant

**Effort: S to M.**

**Why.** `patch_docx` (`spec_formatter/style_application/docx_patch.py`)
copies every source ZIP member unchanged except the replacements, in
source order and compression. That is the transform. Nothing at the final
gate independently proves that parts outside the mode's remit are byte-
identical: `validate_docx_package` checks structure, not identity with the
source, and `_verify_target_header_footer_preserved` covers header and
footer parts only. The README's "applies no document shell of any kind"
claim for the architect-free modes therefore rests on the transform being
right. `_build_and_patch_output` in `batch_runner.py` passes
`word/styles.xml`, `word/settings.xml`, `word/theme/theme1.xml`,
`word/fontTable.xml`, `word/numbering.xml`, `[Content_Types].xml` and
`word/_rels/document.xml.rels` as replacements in every mode (read back
from the extraction directory), so a stray write to any of them in an
architect-free mode would publish.

**Approach.**

1. In `verify_phase2_invariants`, when `new_docx` is given, hash every
   member of source and output (`zipfile` names and SHA-256; the source
   package is already opened there) and classify each as unchanged,
   changed, added or removed.
2. Derive the allowed set from the policy:
   - always: `word/document.xml`;
   - `applies_role_styles` or `apply_full_architect_shell`: `word/styles.xml`;
   - `import_body_numbering` or `apply_full_architect_shell`:
     `word/numbering.xml`, `[Content_Types].xml`, `word/_rels/document.xml.rels`;
   - `apply_full_architect_shell`: `word/settings.xml`,
     `word/theme/theme1.xml`, `word/fontTable.xml`, header and footer parts
     and their `.rels` (added, changed or removed), and `word/media/*`
     additions, each cross-checked against the importer's manifest
     (`env_result["header_footer_import"]` names in `batch_runner`); thread
     that manifest into the gate.
   Anything else changed, added or removed fails with
   `INVARIANT FAIL: package member outside the mode's remit changed: <name>`
   (a part name is not document text).
3. Record `package_members_compared`, `package_members_changed`,
   `package_members_added` and `package_members_removed` in
   `verification_out`; list the names in the run log only.

**Tests.** Per mode, a happy-path test asserting the counters; a test that
alters `word/footnotes.xml`, `word/comments.xml` and a `[trash]/0000.dat`
member in the output and expects failure; for `canadian_to_csi`, any change
to `word/styles.xml`, `word/settings.xml`, `word/numbering.xml` or the
theme fails; for `csi_to_canadian_standalone`, `word/numbering.xml` and its
wiring may change and `word/styles.xml` may not.

**Documentation.** `CLAUDE.md` "Target shell, packaging, and invariants"
bullet; README safety guarantees: "every part outside the mode's remit is
proven byte-identical to the source".

**Definition of done**

- [ ] Whitelist derived from `ApplicationPolicy` fields and the header/footer manifest, not from a hard-coded mode name.
- [ ] Any out-of-remit change, addition or removal fails the target; counters recorded on success.
- [ ] Tests listed above added and passing for all four modes.
- [ ] `CLAUDE.md` and README updated.
- [ ] Full suite and corpus regression green on the PR.

### WI-06: Revision accounting, discarded-revision warnings, collision-proof revision ids

**Effort: M.**

**Why.** Three related lessons. (1) The method documents count
`w:ins`/`w:del` before and after every transform and assert the exact
delta; the engine's run-property invariant catches lost revision *subtrees*
indirectly (through run paths) but there is no explicit census, and
`_without_own_revisions` (`without_own_revisions` in
`core/expected_changes.py` since WI-03) projects out only `w:ins` by
`MARKER_REVISION_AUTHOR`, not the `w:pPrChange` the tracked reverse
conversion also writes. (2) `_remove_existing_hf_files` in
`header_footer_importer.py` deletes the target's header and footer parts
with a log line; any pending tracked change inside them is discarded
silently, and pending changes inside the architect's parts are imported as
accepted-looking content. (3) `_MARKER_REVISION_ID_BASE = 900000` in
`core/canadian_to_csi.py` allocates revision ids from a fixed base
"without needing to scan"; ids must be unique across every annotation in
the document (revisions, bookmarks, comments), and a scan is one regex.

**Approach.**

1. Census: in `verify_phase2_invariants`, count per kind in
   `word/document.xml` before and after: `w:ins`, `w:del`, `w:moveFrom`,
   `w:moveTo`, `w:rPrChange`, `w:pPrChange`, `w:sectPrChange`,
   `w:tblPrChange`, `w:trPrChange`, `w:tcPrChange`, `w:numberingChange`.
   Expected delta: zero for every mode, except tracked `canadian_to_csi`,
   which adds exactly one `w:ins` and one `w:pPrChange` per converted
   paragraph, all authored `MARKER_REVISION_AUTHOR`; additionally require
   that the multiset of revisions by every *other* author is unchanged
   (author and kind, never content). Record `revisions_before`,
   `revisions_after`, `revisions_added_by_application` in
   `verification_out`.
2. Warnings: count revision elements in each target header/footer part
   before `_remove_existing_hf_files` deletes it and in each architect part
   `_write_hf_parts` writes; return the counts through the import manifest;
   log a `WARNING:` line in the run log naming the part and count; record
   `header_footer_revisions_discarded` and
   `header_footer_revisions_imported` in the environment-application
   diagnostics event and the audit. Counts only.
3. Ids: replace the fixed base with `max(existing) + 1`, where existing is
   every `w:id` attribute on annotation elements across `word/document.xml`
   and the header, footer, footnote, endnote and comment parts present in
   the extraction directory (bookmark and comment ranges included). Keep
   allocation monotonic within one conversion. Do **not** assert that every
   `w:id` in the output is unique: valid OOXML repeats an id across paired
   annotations (`w:bookmarkStart`/`w:bookmarkEnd`,
   `w:commentRangeStart`/`w:commentRangeEnd`/`w:commentReference`), so that
   assertion would reject ordinary documents. Assert instead, in
   `tests/test_conversion_verification.py`, that every id the application
   allocated is absent from the source's annotation ids, that the allocated
   ids are distinct from each other, and that ids are unique among revision
   elements (`w:ins`, `w:del`, `w:moveFrom`, `w:moveTo` and the
   `w:*Change` family) in the output.

**Tests.** Census happy paths per mode; a mutated output with one `w:del`
removed fails; tracked reverse conversion reports the exact expected delta;
a target with a `w:ins` in its footer produces the warning count; a source
whose existing ids already exceed the old base gets non-colliding markers.

**Documentation.** `CLAUDE.md` "Canadian to CSI" (ids), the invariants
bullet, and the "Run artifacts" section for the new audit fields. README
safety guarantees: revisions are counted and discarded ones are reported.

**Definition of done**

- [ ] Revision census invariant with exact expected delta per mode; counters recorded on success.
- [ ] Discarded and imported header/footer revisions counted, logged and recorded as counts.
- [ ] Revision ids allocated above every existing annotation id; the independent verifier asserts the allocated ids collide with nothing and revision-element ids are unique, without requiring paired annotation ids to differ.
- [ ] Tests listed above added and passing.
- [ ] `CLAUDE.md` and README updated.
- [ ] Full suite and corpus regression green on the PR.

### WI-07: Marker placement before leading structural run content

**Effort: S.**

**Why.** In `core/canadian_to_csi.py`, `_insert_marker` places an untracked
marker immediately before the paragraph's first `w:t`
(`_FIRST_TEXT_RX`). A `w:tab`, `w:br`, `w:cr`, `w:ptab` or `w:sym` that
precedes the first text node, in the same run or in an earlier text-less
run, therefore ends up *before* the number: Word rendered "1.1, suffix tab,
paragraph tab, text" and the output reads "paragraph tab, 1.1, tab, text".
`_verify_marked_paragraph` and `_verify_prediction` compare whitespace-
normalized text with leading whitespace stripped, so they cannot see it, and
`test_structural_children_change_only_by_the_marker_tabs` counts tabs
without positions. The tracked path places its `w:ins` run before the whole
run that holds the first `w:t`, which covers a leading tab in that same run
but not one in an *earlier*, text-less run; verify the tracked variants too.

*Session 03 notes.* The tracked placement was not even that for any run
carrying properties: a search back for `<w:r` matched `<w:rPr` and put the
revision inside the run. WI-03's PR fixed that. WI-03 also added
`_predict_marker_insertion`, which restates today's placement (before the
first `w:t`) on the run-content signature. This item must move that
prediction with the edit, or the converter's own prediction check refuses
every paragraph it changes.

**Approach.** Insert the untracked marker text node and its tab as the
first content children (after `w:rPr`, and after an inert
`w:lastRenderedPageBreak`) of the first run that contains any content,
which may precede the first `w:t` run; the marker then still inherits that
run's properties. Keep the existing refusal when that run is inside a field
or tracked subtree. Add a structural post-check to `_verify_marked_paragraph`
and to `_verify_prediction`: the marker's text node is the first content
child of the first content run of the paragraph, checked on the XML, not on
normalized text.

**Tests.** `tests/test_canadian_to_csi.py`: `<w:r><w:tab/><w:t>GENERAL</w:t></w:r>`,
`<w:r><w:tab/></w:r><w:r><w:t>GENERAL</w:t></w:r>`,
`<w:r><w:lastRenderedPageBreak/><w:t>...</w:t></w:r>`, a leading `w:br`,
and the tracked variants of each. `tests/test_conversion_verification.py`:
a paragraph with a leading tab in the fixture, and an independent assertion
that the marker is the first run content.

**Documentation.** `CLAUDE.md` "Common mistakes": placing a marker after a
leading tab or break.

**Definition of done**

- [ ] Untracked markers lead the paragraph's run content structurally, not merely after whitespace normalization.
- [ ] Structural post-check added to the edit verification and the prediction check.
- [ ] Tests listed above added and passing.
- [ ] `CLAUDE.md` updated.
- [ ] Full suite green on the PR.

### WI-08: Independent stdlib verifier for the other three modes

**Effort: M. Run after WI-02 through WI-06 so it can assert their properties.**

**Why.** `tests/test_conversion_verification.py` is the repository's
implementation of the method documents' independence discipline: a verifier
that shares no helper code with the engine and can therefore disagree with
it. It covers `canadian_to_csi` only. The other three modes get the same
treatment.

**Approach.** A new stdlib-only module (`zipfile`, `hashlib`,
`xml.etree.ElementTree`, `re`; nothing from `spec_formatter` except the
public entry point used to *produce* the output) named
`tests/test_independent_verification_all_modes.py`. Drive the real engine
without an API key the way `tests/test_unified_roundtrip.py`
(`_deterministic_classifier`) and `tests/test_architect_free_modes.py` do.
For each mode assert, with its own extraction code: the package whitelist
(computed independently from the mode, not from `ApplicationPolicy`);
exact text including deleted text and structural children for every
paragraph the mode did not predict to change, with predictions hand-written
in the fixture and spelled out with `\t` and breaks; the revision census;
revision-id uniqueness scoped as WI-06 defines it; settings parity (WI-04);
and that every output XML part
parses with its declared namespaces. The fixture predictions must be
written before the engine is run, in the test source, not derived from
output.

**Documentation.** `CLAUDE.md` "Development commands" lists the new test
beside the existing independent verifier and repeats the rule that it
shares no code with the engine.

**Definition of done**

- [ ] New stdlib-only verifier covers `format_only`, `csi_to_canadian` and `csi_to_canadian_standalone`.
- [ ] Predictions are hand-written in the test source; no assertion derives its expectation from engine output.
- [ ] The module imports nothing from the engine except the public entry point that produces the output.
- [ ] `CLAUDE.md` updated.
- [ ] Full suite green on the PR.

## 5. Optional work items

Never start these without the user's explicit request in a handoff or in
chat. They do not block completion. Their tracker status stays
`not_scheduled` until the user schedules them.

### WI-09: Stale cross-reference and field-result sweep after numbering conversions

**Effort: M. Report-only; never fails a target.**

**Why.** After a numbering-scheme conversion, literal in-text references
written against the old scheme ("Section 3.4C", "paragraph 2.15D",
"Article 1.02 B") go stale, and REF, PAGEREF, NOTEREF, TOC and SEQ fields
display cached results that Word refreshes only on update. The method
documents flag these to the author rather than rewriting them, because
rewriting is content editing.

**Approach.** In both Canadian conversion modes and the reverse mode, scan
the visible text of body paragraphs for reference-shaped patterns and count
fields whose instruction begins with those keywords. Publish per target, in
`audit.json`, a list of `{paragraph_index, location}` entries using the
validated `ErrorLocation` shape from `core/errors.py` (no free text), plus
counts in diagnostics and a `WARNING:` line in the run log. Surface the
counts in the GUI results. Patterns must be conservative and documented;
false positives cost the user a glance, false negatives cost nothing new.

**Definition of done**

- [ ] Sweep implemented for the three conversion modes, report-only.
- [ ] Audit entries carry indices and validated locations only; counts in diagnostics; run-log warning; GUI shows the counts.
- [ ] Tests for each pattern class and for the no-text guarantee.
- [ ] `CLAUDE.md` and README updated.

### WI-10: Revision dates in Word's own convention

**Effort: S.**

**Why.** The method documents observed that Word writes `w:date` as local
wall-clock time with a `Z` suffix and the true UTC in `w16du:dateUtc`, and
shows true-UTC values hours off in the reviewing pane. `_utc_revision_date`
in `core/canadian_to_csi.py` writes true UTC, which is right by the
standard and wrong by Word's pane. Decide with the user which they want;
if Word's convention, write local time with the `Z` suffix and record
`marker_date_convention` in the audit. Writing `w16du:dateUtc` requires
declaring that namespace on the document root and merging `mc:Ignorable`,
which WI-01's helper makes possible.

**Definition of done**

- [ ] User's decision recorded in the tracker Notes.
- [ ] Chosen convention implemented, tested, and recorded in the audit.
- [ ] `CLAUDE.md` and README updated.

## Appendix A: run-content signature specification (WI-02)

`paragraph_run_content_signature(p_xml)` returns a tuple of per-run tuples.
It is computed on the paragraph after `strip_out_of_scope_subtrees`, so
drawings, text boxes and objects are excluded (they are compared byte-exact
elsewhere). Runs are every `w:r` in document order at any depth below the
paragraph except inside excluded subtrees. Each run's tuple lists its
content children in order:

| Child | Item |
|---|---|
| `w:t` | `("t", decoded_text, preserve)` where `preserve` is whether `xml:space="preserve"` is present |
| `w:delText` | `("delText", decoded_text, preserve)` |
| `w:instrText`, `w:delInstrText` | `("instr", decoded_text)` / `("delInstr", decoded_text)` |
| `w:tab` | `("tab",)` |
| `w:ptab` | `("ptab", alignment, relativeTo, leader)` |
| `w:br` | `("br", type_attr_or_empty, clear_attr_or_empty)` |
| `w:cr` | `("cr",)` |
| `w:noBreakHyphen` | `("noBreakHyphen",)` |
| `w:softHyphen` | `("softHyphen",)` |
| `w:sym` | `("sym", font, char)` |
| `w:fldChar` | `("fldChar", type)` |
| `w:footnoteReference`, `w:endnoteReference`, `w:commentReference` | `("ref", kind, id)` |
| `w:separator`, `w:continuationSeparator`, `w:footnoteRef`, `w:endnoteRef`, `w:annotationRef`, `w:dayShort`, `w:dayLong`, `w:monthShort`, `w:monthLong`, `w:yearShort`, `w:yearLong`, `w:pgNum` | `(local_name,)` |
| `w:lastRenderedPageBreak`, `w:rPr` | excluded |
| anything else | `("other", qualified_name)` so an unknown child is a difference, never silently ignored |

Decoded text uses an XML-only unescape. Nothing is trimmed, collapsed or
mapped. Two paragraphs are identical when their tuples are equal.

*As implemented (session 02):* `kind` in `("ref", kind, id)` is the element's
local name (`footnoteReference`, `endnoteReference`, `commentReference`), and
`type` in `("fldChar", type)` is `w:fldCharType`. Decoded text is exactly what
an XML parser reports: line ends are normalized first (XML 1.0 section 2.11; a
literal CR survives no parser, so Word never sees one, while `&#13;` is
content), then references are expanded with `xml_unescape`; a CDATA section
is literal, and comments and processing instructions are not content. A run
nested in a run's own child (`w:ruby`) follows the run that holds it. The
companion `run_content_difference` names the first differing item from the
closed set `RUN_CONTENT_DIFFERENCE_KINDS`, which adds `preserve_space` (only a
text node's `xml:space` changed) and `run_boundary` (the same items regrouped
into runs, or an empty run added or removed) to one kind per item tag.

## Appendix B: `CT_Settings` child order (WI-04)

The element sequence of `w:settings` in ECMA-376 Part 1, to be verified
against the schema text before it is committed as a constant:

```
writeProtection, view, zoom, removePersonalInformation, removeDateAndTime,
doNotDisplayPageBoundaries, displayBackgroundShape, printPostScriptOverText,
printFractionalCharacterWidth, printFormsData, embedTrueTypeFonts,
embedSystemFonts, saveSubsetFonts, saveFormsData, mirrorMargins,
alignBordersAndEdges, bordersDoNotSurroundHeader, bordersDoNotSurroundFooter,
gutterAtTop, hideSpellingErrors, hideGrammaticalErrors, activeWritingStyle*,
proofState, formsDesign, attachedTemplate, linkStyles, stylePaneFormatFilter,
stylePaneSortMethod, documentType, mailMerge, revisionView, trackRevisions,
doNotTrackMoves, doNotTrackFormatting, documentProtection, autoFormatOverride,
styleLockTheme, styleLockQFSet, defaultTabStop, autoHyphenation,
consecutiveHyphenLimit, hyphenationZone, doNotHyphenateCaps, showEnvelope,
summaryLength, clickAndTypeStyle, defaultTableStyle, evenAndOddHeaders,
bookFoldRevPrinting, bookFoldPrinting, bookFoldPrintingSheets,
drawingGridHorizontalSpacing, drawingGridVerticalSpacing,
displayHorizontalDrawingGridEvery, displayVerticalDrawingGridEvery,
doNotUseMarginsForDrawingGridOrigin, drawingGridHorizontalOrigin,
drawingGridVerticalOrigin, doNotShadeFormData, noPunctuationKerning,
characterSpacingControl, printTwoOnOne, strictFirstAndLastChars,
noLineBreaksAfter, noLineBreaksBefore, savePreviewPicture,
doNotValidateAgainstSchema, saveInvalidXml, ignoreMixedContent,
alwaysShowPlaceholderText, doNotDemarcateInvalidXml, saveXmlDataOnly,
useXSLTWhenSaving, saveThroughXslt, showXMLTags, alwaysMergeEmptyNamespace,
updateFields, hdrShapeDefaults, footnotePr, endnotePr, compat, docVars,
rsids, m:mathPr, attachedSchema*, themeFontLang, clrSchemeMapping,
doNotIncludeSubdocsInStats, doNotAutoCompressPictures, forceUpgrade,
captions, readModeInkLockDown, smartTagType*, sl:schemaLibrary,
shapeDefaults, doNotEmbedSmartTags, decimalSymbol, listSeparator
```

`w:evenAndOddHeaders` sits after `w:defaultTableStyle` and before
`w:bookFoldRevPrinting`; `w:compat` comes much later. Insert by walking the
target's existing children and placing the new element before the first
child whose position in this table is greater.

*Verified (session 04):* the list above matches, element for element, the
`CT_Settings` sequence of the transitional `wml.xsd` in ISO/IEC 29500-4:2012
and in python-docx's `ref/xsd/wml.xsd`. It is committed as
`CT_SETTINGS_CHILD_ORDER` in `core/header_parity.py`, which places the element
after the last child that precedes it (see WI-04's *as implemented* note).

## Appendix C: method-document lessons already covered by the engine

Do not re-implement these; they are listed so you can recognise them.

- Non-breaking spaces and tabs inside section numbers:
  `core/section_numbers.py` accepts them.
- Header/footer placeholders fragmented across runs:
  `_replace_visible_ranges` in `header_footer_importer.py` maps ranges over
  concatenated text back to text nodes.
- Text-box text mirrored in `mc:Choice` and `mc:Fallback`:
  `_validate_relevant_alternate_content_mirrors`.
- Complete `w:pPr` order table: `_PPR_CHILD_ORDER` in `core/canadian_to_csi.py`.
- `xml:space="preserve"` on written and trimmed text nodes: `_with_preserve_space`
  in `core/marker_tools.py`, and the marker writers.
- Repacking in source order and compression, keeping `[trash]` items:
  `docx_patch.patch_docx` and `validate_docx_package`.
- Untracked markers join the first run; tracked ones are projected out for
  the run-structure check: `_insert_marker`, `_without_own_revisions`.
- Predict-first verification of the reverse conversion: `_verify_prediction`.
- An independent, stdlib-only verifier: `tests/test_conversion_verification.py`.
- Rendered geometry proof with `pdftotext -bbox`: `scripts/proof_render.py`.
- Settings applied as compatibility flags only, so `trackRevisions` and
  protection survive: `apply_settings`.
- Ignored paragraphs proven byte-identical: end of `apply_phase2_classifications`.
- An empty numbered paragraph still consumes a number: the reverse
  converter fails closed on any unconverted member of a converted list.
- Word's `w:id` on `w:pPrChange`: always written by `_ppr_change`.
- LibreOffice's missing pStyle-linked numbering levels: documented in
  `scripts/proof_render.py` and `CLAUDE.md`.
- `Normal` is never replaced: collision-safe cloning in `style_import.py`.
- The rule against regex-editing XML is deliberately **not** adopted: this
  engine is regex-based by design, with the defensive machinery `CLAUDE.md`
  describes, and rewriting it is a different project.
