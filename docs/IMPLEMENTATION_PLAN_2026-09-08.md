# Specification Formatter: implementation and validation handoff

**Prepared:** 2026-09-08
**Revised:** 2026-09-08, after implementation review and reproduction on the supported runtime.
**Repository:** `spec-template-normalizer`
**Baseline inspected:** `b66258a`; revision verified against `b156679`
**Status:** **Complete.** W1 (§4), W2-reduced (§6) and W4 (§8) shipped. The §5
spend gate ran on 2026-09-09 and returned *small*, which closes W3 and W5–W8
(§7). Results and the decision table are in
`docs/IMPLEMENTATION_REPORT_2026-09-08.md`.
**Audience:** Coding agents capable of independent investigation, implementation, adversarial testing, and integration review.

**Reading guide:** Section 0 records what this revision changed and why. Sections 1-3 hold the decisions, verified evidence, and invariants. Section 4 is the initial deliverable and can be implemented on its own. Sections 5-7 are the conditional follow-on work and the gate that decides whether any of it happens. Sections 8-11 cover integration, working arrangement, and handoff.

## 0. What this revision changed

The first draft was correct in its findings and wrong in its proportions. It made supporting work mandatory, and its headline W1 recipe omitted a condition that its own detailed requirements supplied. Both are corrected here.

| Area | First draft | This revision | Why |
|---|---|---|---|
| W1 design | Decode bytes, scan the decoded text, re-encode, parse | Keep the existing byte prescan and add an Expat `StartDoctypeDeclHandler` validator over the same immutable payload | Decode-first is defeated by BOM-less UTF-16 unless the NUL rule in the old §5.3.4 is also applied; the validator makes rejection encoding-independent by construction instead of by discipline |
| W1 as replacement | Implied the byte scan would be replaced | Union: nothing is removed | An Expat-only guard silently relaxes declaration-shaped text in comments and CDATA and changes two existing rejection messages; both are verified below |
| Decoded-text handling | Encode a `str` to UTF-8 and parse it | Normalize the declaration first with the existing `prepare_xml_text_for_utf8` | The shipped guard parses UTF-8 bytes back through a stale declaration: `windows-1252` yields mojibake and `utf-16` fails outright. Found in review of this plan; §4.3 records it |
| W0 | A separate work package with a baseline report and ownership assignment | Folded into W1 as ordinary verification | One focused fix does not need a preparatory phase |
| W2 | A complete request ledger across both classifiers | Two items: record architect response usage, and preserve observed usage when either classifier fails | The rest needs a demonstrated purpose |
| W3 | Build the analyzer and evaluation harness | Gated behind inspecting existing spend first | The draft required evidence before optimizing but not before instrumenting |
| Corpus | 12-20 documents across several templates | Existing fixtures plus a few representative hard cases | Gold labels need owner adjudication; that is a scheduling commitment, not an agent deliverable |
| Delegation | Four agents, per-agent report templates, reviewer questionnaire | One implementing agent and an independent review | The coordination machinery was larger than the work it governed |
| Evidence | Reproduction on Python 3.14.6 / Expat 2.8.1, with a caveat about supported versions | Reproduced on Python 3.11.15 / Expat 2.6.1 | Closes the environment gap in the original investigation |

Deferred work is unchanged in substance: W5-W8 remain conditional, and the correctness argument that rejects the review brief's cache key (§7.2) is retained in full.

## 1. Decision and scope

Ship the confirmed XML rejection fix first and on its own. Then look at existing spend before building anything to measure it. Then decide on a small accounting patch. Everything else stays conditional.

**Sequence:**

1. **W1 — encoding-independent rejection of prohibited XML declarations.** Independent, self-contained, verified defect. No dependency on anything else in this document.
2. **Spend inspection.** Read the provider's existing usage reporting for recent real runs. Minutes of work. Decides whether §6 and §7 are worth anything.
3. **W2 (reduced) — record architect response usage, and preserve observed usage when either classifier fails.** Worth doing for honest run diagnostics regardless of what spend shows, but sized by it.
4. **W4 — documentation corrections**, alongside whichever of the above ship.
5. **W3, W5, W6, W7, W8 — conditional.** Implement only against a demonstrated purpose. A measured decision to leave them unimplemented is a complete outcome, not unfinished work.

### 1.1 Authorization and handoff boundaries

- Creating or revising this plan does not authorize application changes. An implementing agent follows the owner's instructions in its own session.
- Use offline fixtures and fake provider responses by default. Do not infer permission to send private specifications to a provider, incur evaluation charges, publish releases, or merge from this document alone.
- Do not modify the original review brief in Downloads. Preserve it as historical evidence; record corrections here or in an implementation report.
- Do not clean unrelated files. Untracked `.pytest_tmp_*` directories in a working checkout are not authorization to remove them.
- Inspect the actual checkout and follow the owner's current branch instructions rather than any branch name carried over from the brief.

## 2. Verified evidence

Everything in this table was checked against the code at revision time. Line references are navigation aids; locate the current function before editing.

| Observation | Evidence | Consequence |
|---|---|---|
| The declaration guard is bypassable by encoding | `core/untrusted_xml.py` matches `<!(?:DOCTYPE\|ENTITY)` against raw bytes; UTF-16 encodes `<!DOCTYPE` as `<\x00!\x00D\x00...` and does not match. A tiny entity payload was accepted and expanded to 16 characters | W1 is warranted as a rejection-contract fix |
| Reproduced on the supported runtime | Python 3.11.15 / Expat 2.6.1 — the Windows CI version. UTF-16 LE with BOM and UTF-16 BE both accepted and expanded; UTF-8 and `str` inputs correctly rejected | Closes the environment gap; the original draft had only Python 3.14.6 / Expat 2.8.1 |
| Decode-first alone does not close it | BOM-less UTF-16 without an XML declaration: `decode_xml_bytes` sniffs it as UTF-8, returns NUL-interleaved text the text regex cannot match, re-encoding round-trips to the original bytes, and Expat re-detects UTF-16 and expands | The NUL rule is part of the design, not a caveat. §4.2 avoids the question entirely |
| Raw-byte callers exist and read untrusted input | `header_footer_importer.py:134,481,716` pass `read_bytes()` from the extracted target package; `phase2_invariants.py:426` passes `zf.read(name)` for every XML part in a package; `core/registry.py:603` passes `path.read_bytes()` | Establishes exposure. Earlier decoding gates may still front-run particular application paths; trace before claiming a specific end-to-end exploit |
| A decoded `str` is parsed through its stale declaration | The guard encodes a `str` to UTF-8 and parses it with the original declaration still in place. A `windows-1252` declaration yields `'Ã©'` for `é`; a `utf-16` declaration fails outright. `read_xml_text` preserves declarations and OPC permits non-UTF-8 parts | A live defect, not only a design-sketch issue. §4.2 step 0 fixes it with the existing `prepare_xml_text_for_utf8` helper |
| End-to-end denial of service is not established | Modern Expat has amplification countermeasures | Fix the prohibition independently of severity. Do not run an unbounded payload to demonstrate it |
| Architect response usage is not accumulated | Root `llm_classifier.py::_call_api` reads `get_final_text()` and `stop_reason`, never `.usage`, and raises on `max_tokens`/`refusal` *after* the final message is in hand | The counts exist at that moment and are discarded. This is the seam for W2 |
| Target usage is returned only on success | `core/llm_classifier.py` calls `_record_usage` correctly before the refusal and `max_tokens` raises, but `result["usage"] = dict(usage_totals)` sits on the success path | A refusal, exhausted regeneration, or merge failure drops every observed count |
| Diagnostics totals are computed over filtered events | `diagnostics.py` drops events below `_min_level` before storing them; `summary()` rolls up `snapshot()` | Emitting usage as INFO events and rolling them up there loses all totals at WARNING |
| Target classification depends on template role definitions | `build_phase2_slim_bundle(..., role_specs=...)`; `role_specs` drives deterministic matching in `core/classification.py:503,569` | Role *names* alone cannot identify cache compatibility. See §7.2 |
| Style lookup already has memoization | `@functools.lru_cache(maxsize=16)` on `_style_block_index` (`core/style_import.py`) and `_parsed_style_elements` (`core/classification.py`) | Do not add another cache on the brief's performance suspicion |
| Two concurrency levels exist | `pipeline.py` runs a target pool; `core/llm_classifier.py` runs a chunk pool; a process-wide `BoundedSemaphore` caps open streams | Document both accurately; do not rewrite concurrency as part of this work |
| The classifiers differ specifically in `Retry-After` handling | Target honours a numeric `retry-after` header (`_retry_after_seconds`); architect `_call_api` uses fixed exponential sleeps. Both set `max_retries=0` with identical timeouts | Instrument the actual policies. Do not unify them; the delta is narrow and a redesign needs its own evidence |
| Boolean rejection already exists target-side | `_usage_numbers` skips `bool` and non-`int` values | Existing behaviour to preserve in any shared contract, not new work |
| The token-count guard transmits document text | `_count_input_tokens` sends the system prompt and the full user message, which carries the slim bundle | An offline dry run must not call the count endpoint. The count is free and is not billed generation usage |
| The target prompt is ambiguous about formatting | `core/prompts/phase2_master_prompt.txt` permits indentation as evidence (line 42) and says "Do NOT reference formatting" (line 57) | Clarify evidence versus output instructions before removing hints. Prompt edits belong in W5 |
| Reported payload savings are character measurements | The brief measured JSON character lengths on a small constructed sample | Do not label them token, cost, or accuracy improvements |
| The suite is green | `1049 passed, 3 skipped in 11.57s` | The brief's figure is current. Confirm in the implementing environment rather than assuming it stays so |

## 3. Non-negotiable invariants

Read `CLAUDE.md` before implementation. Retain these throughout.

1. Architect and target source files remain immutable. Process private snapshots and publish new outputs only.
2. In `format_only`, preserve target body text and effective numbering semantics. This does not mean byte-identical DOCX packages; presentation and package serialization legitimately change.
3. In `csi_to_canadian`, retain the existing supported conversion boundary and rejection behaviour.
4. Every classifiable paragraph has exactly one styled or ignored disposition. Keep deterministic-override rejection, coverage checks, and current out-of-scope treatment.
5. Keep source-derived styles, collision-safe style IDs, effective inheritance resolution, protected subtrees, and target numbering ownership.
6. Keep the single shared application path and the immutable application policy. Do not add a second formatter implementation for caching, evaluation, or error handling.
7. Keep strict profile validation and the committed engine identity. Do not weaken cache invalidation to avoid the cost of a change.
8. Keep per-run isolation, atomic output publication, source rechecks, and independent per-target failure handling.
9. Persist only code-defined identifiers, counts, booleans, safe timing values, and existing approved provenance in diagnostics. Never persist prompts, responses, paragraph text, API keys, HTTP bodies, or arbitrary provider objects.
10. Preserve public and injected seams where practical. New optional parameters and additive result fields need safe defaults. Do not require existing injected classifiers to emit usage.
11. Python 3.10 remains the source floor; Windows remains the primary platform.
12. Do not change model defaults, effort, retry ceilings, concurrency limits, or chunk overlap in a security or telemetry patch.

## 4. W1 — encoding-independent rejection of prohibited XML declarations

This is the initial deliverable and stands alone. **Implemented.** The union guard
ships in `core/untrusted_xml.py` with the §4.5 matrix in
`tests/style_application_regression/test_untrusted_xml.py`. Measured on the real
corpus payloads (109 parses, 0.29 MB): 6.55 ms to 8.86 ms, so **+2.32 ms per
corpus run** — 35% of parse time and about 0.5% of the run's wall clock. Two
further encoding leaks found during implementation are recorded in §4.8.

### 4.1 Intended behaviour

Every untrusted XML entry through `parse_untrusted_xml` rejects a DOCTYPE or ENTITY declaration before entity expansion, regardless of whether the caller supplied bytes or text and regardless of encoding. Valid supported OOXML still parses with correct Unicode content. Unsupported or malformed encodings fail predictably. **Every rejection the current guard performs is still performed, with the same error type and message.**

### 4.2 Design: union guard

Keep the existing byte prescan. Add an Expat validator. Parse the same immutable payload with all three steps.

```python
# 0. A str is already decoded. Make its declaration truthful before
#    encoding, or both parsers will read the UTF-8 bytes back through the
#    stale declared encoding. Reuses the existing shared helper.
if isinstance(data, str):
    payload = prepare_xml_text_for_utf8(data).encode("utf-8")
else:
    payload = data

# 1. Existing conservative prescan - unchanged, nothing removed.
if _DOCTYPE_RE.search(payload):
    raise UntrustedXmlError(f"{context}: DOCTYPE/ENTITY declarations are not allowed ...")

# 2. Expat rejects a real declaration in any encoding it can auto-detect.
#    StartDoctypeDeclHandler fires as Expat begins the document-type
#    declaration, before any entity is expanded.
# 3. Then parse that same payload with ElementTree.
```

Four properties make this the right shape:

1. **One immutable payload** feeds the prescan, the validator, and the parse. The old draft's requirement that "the representation that is checked and the representation that is parsed must be equivalent" becomes true by construction rather than something an implementer must argue. There is no second encoding interpretation to keep in sync.
2. **Nothing is removed**, so the current conservative screening policy survives intact — including declaration-shaped text inside comments and CDATA, which the plan already required preserving for this patch.
3. **`UntrustedXmlError` handling is untouched.** It stays a `ValueError` subclass, keeps the part-name context, and keeps its existing message for every case that reaches it today.
4. **A decoded `str` keeps its characters**, because step 0 makes the declaration agree with the bytes actually produced. See §4.3.

Step 0 fixes a latent defect in the current guard rather than merely preserving it; §4.8 records that as a deliberate behaviour change. `prepare_xml_text_for_utf8` is idempotent, so the existing pre-normalizing call at `arch_env_applier.py:223` becomes redundant but stays harmless — leave it or remove it, but do not make removal a condition of this patch.

Both the validator and `ElementTree` use the same Expat build, so they cannot disagree about encoding detection or well-formedness.

### 4.3 Two designs that do not work, with evidence

Record both in the implementation report so neither is reintroduced.

**Decode-first is insufficient on its own.** Decoding with the shared helper, scanning the decoded text, then re-encoding and parsing is defeated by BOM-less UTF-16 with no XML declaration: the sniff falls through to UTF-8, the decoded text is NUL-interleaved so the text regex cannot match, re-encoding round-trips to the original bytes, and Expat re-detects UTF-16 and expands the entity. Rejecting XML-invalid literal NULs closes it, which is why that rule is load-bearing rather than a test case. §4.2 removes the need for the rule entirely.

**Expat-only, replacing the byte scan, is a regression.** Measured across nine cases:

| Case | Today (byte scan) | Expat only | Union (§4.2) |
|---|---|---|---|
| Real doctype, UTF-8 | REJECT | REJECT | REJECT |
| Doctype after a prolog comment | REJECT | REJECT | REJECT |
| `<!doctype html>` lowercase | REJECT | **ExpatError** | REJECT |
| `<!ENTITY` in element content | REJECT | **ExpatError** | REJECT |
| DOCTYPE-shaped text inside a COMMENT | REJECT | **accept** | REJECT |
| DOCTYPE-shaped text inside CDATA | REJECT | **accept** | REJECT |
| DOCTYPE-shaped text escaped as content | accept | accept | accept |
| Real doctype, UTF-16 with BOM | **accept** | REJECT | REJECT |
| Real doctype, BOM-less UTF-16 LE | **accept** | REJECT | REJECT |

Replacing the scan silently relaxes two cases the plan requires preserving, and changes two more from `UntrustedXmlError: DOCTYPE/ENTITY` to a wrapped parse error — which fails the existing parametrized test in `tests/style_application_regression/test_untrusted_xml.py`. Only the union is correct on all nine.

Note that `<w:p>...<!ENTITY a "b"></w:p>` is not well-formed XML at all; the current guard is deliberately stricter than well-formedness requires, and the union keeps that intent. A later cleanup must not drop it as redundant.

**Encoding a decoded `str` without normalizing its declaration is a live defect today.** `parse_untrusted_xml` currently does `data.encode("utf-8")` on a `str` and hands the result to a parser that still believes the stale declaration. OPC permits non-UTF-8 parts, and `read_xml_text` returns decoded text with its original declaration intact, so this is reachable rather than theoretical:

| `str` input | Current guard | With step 0 |
|---|---|---|
| `encoding="windows-1252"`, content `é` | `'Ã©'` — silent mojibake | `'é'` |
| `encoding="utf-16"`, content `é` | `UntrustedXmlError: XML parse error` | `'é'` |
| `encoding="UTF-8"` or no declaration | correct | correct |

Both parsers agree on the mis-declared payload, so the union guard is internally consistent either way — they simply agree on the *wrong* interpretation. Only normalizing the declaration fixes it.

Reachable callers passing decoded text straight through include `core/classification.py:385` (`numbering_xml_text`) and `core/classification.py:1464` (`styles_xml_text`). `arch_env_applier.py:223` is the one site that already normalizes, which is evidence the hazard was known and handled locally rather than at the boundary.

Two fixes were considered. Handing the original `str` to `ElementTree` also decodes correctly, but it makes the checked and parsed representations differ by type, cannot feed the byte prescan without re-encoding anyway, and leans on parser-specific `str` handling across the supported interpreter range. Normalizing with the existing `prepare_xml_text_for_utf8` keeps one bytes payload for all three steps and reuses a helper the codebase already relies on in `write_xml_text`. Prefer it.

### 4.4 Files

Primary:

- `spec_formatter/style_application/core/untrusted_xml.py`
- `tests/style_application_regression/test_untrusted_xml.py`

`core/ooxml_text.py` needs no change *for the guard itself*; §4.2 step 0 calls its
existing `prepare_xml_text_for_utf8`. A separate defect in that helper was found
during implementation and is stated here as this section requires:
`_TEXT_DECLARED_ENCODING` was unanchored, so it rewrote the first
declaration-shaped text *anywhere* in a part. A part with no prolog but with
`<?xml ... encoding="..."?>` inside CDATA had that content silently edited --
through `write_xml_text` onto disk and through `docx_patch.py:135` into a
published part, which is target document content the formatter must never
change. The pattern is now anchored to the start of the document, where a
declaration can legally appear, with an optional leading BOM.

Trace and exercise callers in `header_footer_importer.py` (`_remove_existing_hf_files`, `_rebuild_document_rels`, `_ensure_content_types`), `phase2_invariants.py::validate_docx_package`, `core/registry.py` bundle-artifact loading, `docx_patch.py::validate_xml_wellformedness`, and `arch_env_applier.py` content-type and relationship preparation.

### 4.5 Test matrix

Use tiny payloads. Do not run an unbounded amplification payload.

| Case | Expected |
|---|---|
| Valid UTF-8 XML bytes, with and without BOM | Correct root and text |
| Valid UTF-16 LE/BE with BOM and matching declarations | Correct root and text |
| BOM-less UTF-16 LE/BE, with and without declarations | Correct parse or documented rejection; never unchecked expansion |
| Declared single-byte encoding with non-ASCII text (e.g. windows-1252) | Characters preserved |
| Already-decoded text declaring `windows-1252`, `utf-16`, `UTF-8`, and no declaration | Characters preserved in every case; specifically `é` never becomes `Ã©` |
| Tiny DOCTYPE plus internal entity in each supported encoding | `UntrustedXmlError` before expansion |
| External SYSTEM/PUBLIC declarations | Rejected without filesystem or network dereference |
| Every existing case in the current parametrized rejection test | Same exception type and message as today |
| Declaration-shaped text inside comments and CDATA | Still rejected; conservative policy unchanged |
| UTF-32 variants | Documented behaviour; Expat does not support UTF-32 and `ET.fromstring` already rejects it today, so raw UTF-32 bytes are no regression |
| Truncated data, unknown encoding, literal NUL, invalid Unicode | Predictable wrapped failure |
| Valid escaped text, comments, namespaces, non-ASCII attributes | No regression |

A focused assertion that a prohibited payload never reaches the parse call is justified here, because "before expansion" is the security contract rather than an implementation detail. Pair it with behaviour-level tests.

### 4.6 Reachability evidence

1. Build minimal synthetic packages carrying prohibited declarations in `word/_rels/document.xml.rels` and `[Content_Types].xml`.
2. Exercise both the direct helper callers and the normal application and package-validation paths.
3. Trace shell application order. Theme, settings, and font-table helpers may already decode or reject parts before header/footer import. Record those earlier gates rather than claiming every raw-byte site is independently exploitable.
4. Include a valid package with non-UTF-8 parts as a positive control.
5. Confirm failing targets publish no partial DOCX and that source hashes are unchanged.
6. Cover bundle-artifact parsing separately from the target package path.

Keep the confirmed bypass of a stated prohibition distinct from a demonstrated denial of service. Neither distinction is a reason to delay the fix.

### 4.7 Performance

Measured on the sanitized corpus regression, which is the only realistic parse workload in the repository.

- Actual volume for one corpus run: **109 `parse_untrusted_xml` calls, 0.29 MB total, median part 943 B, p90 3.2 KB, largest 31 KB.**
- Added cost of the Expat validator at that size distribution: about 40% of parse time at the median, 33% at p90, 23% at the largest part. The percentage is worst on small parts because parser construction dominates.
- Absolute cost: **roughly 6 ms added per corpus run**, against a run that makes model calls taking seconds.

Two honest caveats. These fixtures are small, and real specification sections are considerably larger; the cost scales with parse volume and stays near a quarter of parse time, so even a 1 MB `document.xml` adds single-digit milliseconds. And this is a fixture-level measurement, not an application-wide result — confirm by timing the corpus regression before and after the change, which is cheap once implemented.

### 4.8 Verification and acceptance

Ordinary verification, not a separate phase:

```powershell
git rev-parse HEAD
python --version
python -c "import pyexpat; print(pyexpat.EXPAT_VERSION)"
python -m pytest -q
```

Then reproduce the defect on the actual starting commit, implement the union guard, and run:

```powershell
python -m pytest -q tests/style_application_regression/test_untrusted_xml.py tests/style_application_regression/test_ooxml_text.py tests/test_ooxml_text.py tests/style_application_regression/test_docx_patch.py tests/style_application_regression/test_final_package_validation.py
python -m pytest -q
python -m pytest -q tests/test_sanitized_format_only_corpus.py
```

Acceptance:

- The original bypass fails for bytes and text, across the tested encodings.
- Every existing rejection keeps its current exception type and message.
- Three deliberate behaviour changes, all fixing defects rather than relaxing the
  contract. A decoded `str` with a non-UTF-8 declaration now parses with correct
  characters instead of mojibake, and a `utf-16`-declared `str` now parses instead
  of raising. An encoding expat cannot use — a codec Python lacks (`LookupError`)
  or a multi-byte one it refuses internally (`ValueError`) — now raises
  `UntrustedXmlError` with the part name instead of escaping unwrapped past every
  caller that handles it. That last one was found by an adversarial encoding sweep
  during implementation, not by the original review.
- Valid content retains its Unicode across supported encodings.
- Package validation, both application modes, and the corpus regression are unaffected.
- Python and Expat versions are recorded for the parser results: implemented and
  verified on Python 3.11.15 / Expat 2.6.1.
- Corpus regression timed before and after: 6.55 ms to 8.86 ms of parse time over
  its 109 parses, so +2.32 ms per run.
- The change is one isolated commit. If compatibility regresses, revise the approach; never restore acceptance of prohibited declarations.

W1 does not touch `engine_identity.py`'s covered files, so it does not invalidate cached profiles.

## 5. Spend inspection gate

**Ran 2026-09-09. Result: small. W3 and W5–W8 are closed.** The figure itself
is deliberately not recorded here — this repository is source-available, and
the owner's real API spend is private financial information that does not
belong in a public file. The decision is what the record needs. See
`docs/IMPLEMENTATION_REPORT_2026-09-08.md` §4 for the decision table, and
re-run this gate if the workload changes materially.

The original reasoning follows, retained because it explains why the gate
came before the work rather than after it.

Before building anything to measure cost, look at what is already known.

Read the provider's existing usage reporting for recent real runs and answer one question: **is architect and target classification spend material enough to justify optimization work?**

- If it is not, W3 and W5-W8 are closed. Record the figure and the date. W2 shrinks to §6, which is worth doing for honest diagnostics on its own.
- If it is, W3 becomes justified, and §7 defines what it must produce.

This costs minutes. The first draft required an evaluation harness before establishing that optimization mattered, which inverted its own evidence-first principle.

## 6. W2 (reduced) — honest usage accounting

**Implemented.** `spec_formatter/llm_usage.py` owns the shared contract;
`Phase1Result.usage` carries architect counts; both classifiers hand their
observed counts out on the exception when they fail. The engine digest moved to
`cdfe7148986b940e` because `llm_classifier.py` is a covered file, so cached
architect profiles are invalidated once — the accepted cost recorded in §6.4.

Two items, both worth doing independently of what §5 shows, because both currently produce untruthful run artifacts:

1. **Record architect response usage.** `_call_api` obtains the final message and then raises for `max_tokens` and `refusal` without reading `.usage`. Capture usage immediately after the final message and before any stop-reason handling.
2. **Preserve observed usage when either classifier fails.** Target-side counts are collected correctly but published only on the success path. Use `finally` or an equivalent reliable handoff so a refusal, exhausted regeneration, merge failure, or later application failure still carries what was observed into the failed run artifacts.

A comprehensive request ledger, a shared usage module, per-request purpose enums, and cross-run aggregation need a demonstrated purpose from §5. If that purpose does not materialize, implement the two items above with the smallest sound contract and stop.

### 6.1 Semantics to settle before coding

Whatever the eventual size, these meanings must be unambiguous. Publish the field table in the engineering guide.

| Concept | Definition |
|---|---|
| `requests_attempted` | Actual model stream/create invocations, including transport retries and grammar fallback requests |
| `responses_completed` | Requests for which the final message was obtained; includes refusals and output-limit responses |
| `responses_with_usage` | Completed responses carrying at least one valid recognized usage field; this alone does not prove complete accounting |
| `input_tokens` | Valid provider-reported uncached input tokens under the provider's current usage contract |
| `output_tokens` | Valid provider-reported output tokens, including billed thinking tokens where the provider includes them |
| `cache_read_input_tokens` / `cache_creation_input_tokens` | Provider-reported cache read and cache write input tokens |
| `usage_complete` | Whether all usage necessary for the reported scope is known; false when a request lacks required final counters |
| `requests_with_unknown_usage` | Requests whose final billable usage cannot be established from observed metadata |
| `transport_retries` / `regenerations` | Additional transport attempts, distinct from JSON regeneration under existing policies |

Rules:

- **Missing usage is unknown, not zero.** Accept zero only when reported or when a documented provider contract supplies it.
- Rejecting booleans and non-integers as counters is **existing target-side behaviour to preserve**, not new work.
- Do not add uncached input counts to an already inclusive total. Verify the provider schema before writing any cost formula.
- The input-size token guard is separate from billed generation usage, and `_count_input_tokens` transmits document text — it must not run in an offline path.
- Preserve request purpose with fixed enums and numeric identifiers. Never document titles, paragraph excerpts, response text, or exception strings.
- Do not sum nested phase timings as run wall-clock time.

### 6.2 Propagation constraints

- Thread an optional keyword-only collector through production calls. `classify_document` keeps returning only the validated instruction object; telemetry never enters instruction JSON or its schema.
- Record attempted requests even when opening or consuming the stream fails, and mark usage unknown rather than manufacturing an estimate.
- On the target path, let running chunk work settle per existing executor behaviour before finalizing that target's accounting, so a failing chunk cannot hide a sibling's usage.
- Choose one authoritative contribution per request. The existing `result["usage"]` may remain a compatibility summary but must not be added again to request-level totals.
- Injected classifiers without telemetry support must keep working and report usage unavailable. Pass new arguments only through an explicit supported signature. **Never detect an unsupported keyword by catching `TypeError` and calling again** — that can repeat a real paid request.
- Observers must not change successful output or replace the primary classification exception.
- **Totals must survive verbosity.** `DiagnosticsRecorder.summary()` rolls up level-filtered events, so usage emitted only as INFO events vanishes at WARNING. Accumulate validated usage independently of verbosity, and do not claim suppressed events were written to `diagnostics.jsonl`.
- Report a reused architect profile as no new architect request. Never charge a historical creation's cost to the current run.
- Report deterministic-only target classification as zero model requests with known zero usage.
- Keep dollar estimates out of runtime code.

### 6.3 Tests

Provider fakes only; no paid calls in the suite. Cover: one architect response; malformed-then-valid JSON; output limit then regeneration; terminal refusal; targeted coverage patch; structured-output compiler fallback; transport failure then success; fresh analysis followed by style or bundle failure; reused profile; strict legacy injected fakes; multiple target chunks; target refusal or exhausted regeneration; one chunk failing while another finishes; overlap re-ask; merge failure after responses; deterministic-only target with no client constructed; missing or malformed usage fields; parallel runs; every verbosity level; and provider metadata shaped like secrets or body text.

Primary files: `tests/test_llm_classifier_safety.py`, `tests/test_phase1_pipeline.py`, `tests/test_unified_pipeline.py`, `tests/test_diagnostics.py`, `tests/style_application_regression/test_llm_classifier.py`, `tests/style_application_regression/test_batch_runner_failure_diagnostics.py`, `tests/style_application_regression/test_batch_runner.py`.

### 6.4 Cache-invalidation timing

Changing root `llm_classifier.py` changes `ENGINE_SOURCE_DIGEST`, invalidating every cached architect profile and forcing a fresh paid analysis per template on the next run. Keep the conservative digest — do not weaken it with a comment-stripping heuristic. Where practical, land architect telemetry together with other already-planned changes to the covered engine files so one invalidation covers both. Do not invent changes merely to amortize it, and do not delay W1 for it. Recompute with `python engine_identity.py` after final integration and explain the decision in the PR.

## 7. Conditional work

**All closed by the §5 gate on 2026-09-09.** None was rejected on merit; each
is closed because the evidence that would justify it does not exist and, at
this spend, is not worth generating. The specifications below are retained so
any of them can be reopened on the same terms if the workload changes — in
particular §7.2, whose correctness argument must be read before anyone
attempts a target-classification cache.

None of this is authorized by this document. Each needs a demonstrated purpose from §5 and its own review.

### 7.1 W3 — workload measurement and evaluation

If §5 shows material spend, build a developer tool with a tested importable core that uses production extraction and slim-bundle logic rather than recreating numbering or classification rules.

Required behaviour: accept explicit targets and a validated profile or fixture role definitions; snapshot through existing bounded helpers into tool-owned temporary directories, never a new unchecked ZIP loop; **make no provider calls by default, including the token-count endpoint**; produce a machine-readable counts report plus a short Markdown note on sample and limits; use opaque identifiers and keep document text and private paths out of shareable reports; report skipped, failed, and empty documents alongside successful ones.

Report per target/profile: `classifiable_total`, `deterministic_classified`, `deterministic_ignored`, `unresolved_sent`, `unresolved_fraction` (not-applicable when the denominator is zero), `out_of_scope` counted separately, request characters by component, chunk count, paragraph presentations including overlap, and candidate representation sizes. Check that `deterministic_classified + deterministic_ignored + unresolved_sent == classifiable_total` using the production index and disposition contract, not raw `filter_report` entries.

**Corpus:** start with existing sanitized fixtures plus a few purpose-built cases for structures they miss — automatic numbering inherited through style chains, typed markers at depth, unmarked continuation prose, repeated Canadian numeric markers, Roman/alpha ambiguity, cross-references, heading-shaped requirements, editorial boilerplate, and chunk-boundary and overlap disagreements. Expand only when a specific optimization justifies the owner's adjudication time. Gold labels for ambiguous CSI material require domain review by the owner; that is a scheduling commitment and the real constraint on this package, not a deliverable an agent can complete. Separate development and holdout material by document and template family.

Keep measured, estimated, and unavailable quantities in separate columns. Character reduction is never a token or dollar measurement. When any request has unknown usage, label the dollar result a lower bound; the invoice remains authoritative.

### 7.2 W6 — target-classification cache

Retained in full, because it is the strongest correction to the review brief.

**The brief's proposed key is incorrect, not merely coarse.** `role_specs` — the architect's portable numbering patterns — flows into `build_phase2_slim_bundle` and drives deterministic classification (`core/classification.py:503,569`). Two architect templates with identical role *names* but different numbering patterns produce different deterministic dispositions, a different unresolved set, and a different request. Keying on target hash plus role names would serve a wrong cached classification.

If ever implemented, cache model-derived dispositions for an exact classification request plan. Rebuild the target bundle and deterministic dispositions every run, and reapply every local validator, deterministic-override check, and coverage check before application. Never cache formatted DOCX outputs or bypass source snapshots, run isolation, application policy, package validation, or publication.

Cache identity must cover: target source identity and paragraph-index universe; exact template-derived role definitions including numbering patterns, provenance, and counter constraints; the actual unresolved paragraph data and all evidence sent to the model; system and user instructions, serialized role ordering, response schema, and wire-projection version; provider and model identity, effort, and output constraints; chunking, overlap, and re-ask strategy; classification preprocessing and merge semantics with a compatible fingerprint; and an explicit cache contract version. The five-file architect engine digest does not cover target preprocessing. Conversion mode may be omitted only after a test proves it changes neither the cached request nor its interpretation.

Storage: versioned private namespace, bounded entries and bytes, explicit pruning, atomic writes, no partial entries. Treat corruption as a safe miss. Checksums detect corruption but do not authenticate a hostile local writer — state that boundary honestly. Disable reuse for injected classifiers without stable identity. Report zero new API requests on a hit and never charge historical usage to the current run.

Acceptance requires a two-template, same-role-list regression that fails under the brief's proposed key. Without that correctness argument and an economic case, do not implement.

### 7.3 W5, W7, W8

- **W5 (payload reduction):** evaluate before changing any default. Keep the full slim bundle as local authority and build a separate wire projection; never mutate shared paragraph dictionaries. Candidates in order: remove the duplicated role list from the user message; omit only fields with explicitly defined default semantics via a per-field allowlist (never a recursive truthiness filter — it drops `False`, zero, paragraph index zero, and meaningful empty collections); remove duplicated neighbour text only where exact context is present; and resolve the formatting-evidence ambiguity in the prompt. Evaluate candidates independently. An inconclusive evaluation means keep production unchanged.
- **W7 (architect replay cache):** do not remove `ENGINE_SOURCE_DIGEST` from profile validation or replace it with a heuristic. A source-file hash does not identify the model input, since a slim-bundle change alters the request for identical DOCX bytes, and stored post-repair instructions cannot stand in for a pre-repair response after repair logic changes. Do not implement this tier merely to avoid one invalidation caused by W2.
- **W8 (effort, model, chunking, deterministic rules):** keep current defaults. One experiment at a time on the same adjudicated corpus. A new deterministic rule needs a concrete unresolved pattern, a defensible structural discriminator, and close negative cases; the metric is correct resolution with no precision regression, not a lower unresolved percentage. Do not enlarge chunks because a context window is larger, and do not serialize requests for cache hits without measuring latency.

Also excluded: another inherited-style cache without profiling, splitting or renaming classifier modules, reviving the Batch API or retired runners, and altering GUI, updater, licensing, or CI platform coverage without a task-specific reason.

## 8. W4 — documentation corrections

**Implemented.** Items 1 and 2 shipped with W1 and W2. Items 3, 4, 5 and 9 are
now one "Concurrency, retries, and caches" section in `CLAUDE.md`; items 6 and
7 sit with the invariants they qualify, and item 6 also appears in `README.md`
in user-facing terms. Item 10 was already in the change checklist. Item 11 is
not applicable: no evaluation tooling was built, because §5 has not run. Item 8
is closed by the corrections table below rather than by a doc edit, since the
brief it corrects is not in this repository.

Update `CLAUDE.md` and `README.md` alongside whatever ships. Keep the engineering detail; do not replace substantive guidance with a shorter summary.

1. Document encoding-independent declaration rejection, the preserved conservative screening policy, and the distinction between a parser-contract failure and a demonstrated exploit.
2. Explain usage fields, observed versus unknown counts, failure accounting, profile reuse, deterministic-only work, and verbosity-independent totals — scoped to what actually ships.
3. Describe concurrency accurately: one public target orchestration path, an internal chunk pool in the target classifier, and a process-wide request semaphore that does not cover root architect calls.
4. Describe each classifier's retry behaviour specifically: the target honours `Retry-After`; the architect uses fixed exponential backoff. Both disable SDK retries.
5. State that inherited-style lookup already uses `lru_cache`. Do not add speculative performance claims without profiling.
6. Distinguish exact body-text and numbering preservation from whole-package byte identity, and note that unedited ignored-paragraph XML does not by itself prove unchanged rendered appearance when the shell or document defaults change.
7. Distinguish zero LLM involvement from correct classification.
8. State that payload size measurements are not token or dollar measurements.
9. Document that prompt-cache reads are affected by concurrency, minimum prefix size, TTL, and prefix identity; zero reads alone identifies no single cause.
10. Keep visual inspection of representative Word output in the validation checklist for formatting, numbering, or shell changes.
11. Describe any evaluation tooling as developer tooling, not a step in the GUI workflow.

Do not edit production prompts for wording during W4; the formatting-evidence ambiguity is a model-input change belonging to W5.

### Brief corrections to close in the implementation report

| Brief claim | Correct treatment |
|---|---|
| Omit empty/false fields with zero information loss | Define field defaults and evaluate; absence can differ from explicit false or null |
| Target hash plus role names suffices for cache reuse | Incorrect — include template-dependent role definitions and exact classification inputs (§7.2) |
| Style numbering lookup needs memoization | Already cached; profile before further work |
| Zero cache reads means prefix invalidation | Check concurrency, warm-up, TTL, and minimum size too |
| Only one thread pool exists | One public runner plus internal chunk concurrency |
| Classification is intrinsically a poor use of high effort | A workload-specific hypothesis, not a general fact |
| More deterministic resolution always improves quality | Requires precision evidence and adversarial coverage |
| Payload reductions are measured token savings | Character reductions until measured with the actual tokenizer |
| A review verdict must have no hedging | State a clear verdict with explicit coverage and uncertainty |

## 9. Integration and verification

### 9.1 Change sequence

| Sequence | Scope | Must stand on its own |
|---|---|---|
| PR 1 | W1 union guard, encoding matrix, documentation | Safe parser behaviour; no telemetry or optimization changes |
| — | Spend inspection (§5) | A recorded figure and date; no code |
| PR 2 | W2 reduced accounting and its tests and docs | Correct accounting with no prompt or request-policy change |
| PR 3+ | Only what §5 justifies | Its own evidence and compatibility review |

### 9.2 Version and resource checklist

| Surface | When to update |
|---|---|
| `engine_identity.py::ENGINE_SOURCE_DIGEST` | Any covered root source changes; compute after final integration. W1 does not touch these files |
| `PIPELINE_VERSION` | Only if architect pipeline compatibility requires it; not a substitute for the digest |
| Profile contract / manifest versions | Only if consumer-visible profile assumptions change |
| Run audit / manifest versions | Assess additive usage fields against readers; document the decision |
| `requirements.txt` | New or changed direct runtime dependency only. Neither W1 nor reduced W2 adds one — both use the standard library |
| `THIRD_PARTY_NOTICES.md` | Regenerate through the build process if dependencies change; never hand-edit |
| PyInstaller spec | New dynamically imported runtime modules or resources |
| `README.md`, `CLAUDE.md` | Changed behaviour, diagnostics, constraints, validation requirements |

### 9.3 Runtime and platform checks

1. Run the full suite on Windows / Python 3.11, the authoritative gate.
2. Run W1's tests on Python 3.10 as well. The Linux CI job installs `requirements.txt` after its updater tests; a small focused step after that installation covers the new parser behaviour. Preserve pinned action SHAs and avoid unrelated CI redesign.
3. Record Python and Expat versions with any parser result.
4. Confirm a frozen Windows build imports the changed code and loads prompts as before, using `docs/RELEASE_WINDOWS.md`. Do not publish a release automatically.

### 9.4 Document and visual acceptance

Exercise both application modes with offline classifier seams and representative fixtures. Verify source hashes, body text, numbering semantics, protected structures, collision-safe styles, relationships, package validity, and no partial publication.

For W1 and reduced W2, unchanged request and output behaviour plus regression tests are the primary evidence; telemetry needs no new visual standard. For any change that affects resulting documents, inspect representative output in Word against the architect template and original target, rendering every page when formatting or shell behaviour changes. Use disposable copies. Record the Word version, font availability, cases selected, and any limitation.

### 9.5 Acceptance checklist

- [ ] W1 rejects prohibited declarations before expansion across the tested encodings.
- [ ] Every rejection the current guard performs still happens, with the same exception type and message.
- [ ] Valid supported non-ASCII content is preserved and failure behaviour is stable.
- [ ] Corpus regression timing recorded before and after W1.
- [ ] Existing spend inspected and recorded, with a decision on §7.
- [ ] Architect response usage is captured, and observed usage survives classifier failure.
- [ ] Unknown usage stays explicitly unknown; no silent zero-cost claims.
- [ ] Usage totals survive every diagnostic verbosity level.
- [ ] No document text, secrets, raw responses, or arbitrary provider metadata in artifacts.
- [ ] Injected seams still work with no duplicate request from compatibility handling.
- [ ] Model, effort, retry, concurrency, and formatting defaults unchanged.
- [ ] Documentation, versions, and digests match what shipped.
- [ ] Focused and full suites pass in the required environments, with skips explained.
- [ ] Each conditional package has an implemented, deferred, or rejected decision with evidence.
- [ ] Unrelated files and prior outputs untouched.

## 10. Working arrangement

One implementing agent and one independent review are sufficient for W1 and for reduced W2. Implement sequentially. If more agents are ever warranted, keep them to independent files and evidence gathering, and never create competing usage schemas or a second extraction or classification implementation.

A completion report should give the concrete problem and resulting behaviour, evidence confirmed and any hypothesis disproved, files changed with reasons, tests run with interpreter and platform, invariants touched and how preservation was shown, compatibility and fingerprint implications, remaining risks, and the commits.

An independent reviewer should be able to answer these from the diff:

- Can the parser interpret a representation different from the one screened?
- Does every rejection the old guard performed still happen, with the same message?
- Are valid non-ASCII inputs preserved and unsupported encodings rejected consistently?
- Can a paid response occur before an exception that discards its observed usage?
- Can a worker finish after the accounting snapshot was finalized?
- Can the same request contribute through both a callback and a returned summary?
- Can a reused profile import historical usage into a new run's totals?
- Can logging verbosity remove cost information?
- Can an injected fake cause a second real invocation during argument fallback?
- Did telemetry change prompt construction, model settings, retries, or validators?
- Are any reported savings character counts presented as observations?

## 11. Handoff artifacts

1. Reviewable commits or PRs for what actually shipped.
2. `README.md` and `CLAUDE.md` updated for the parser contract, and for accounting if it ships.
3. Focused adversarial parser tests and accounting regression tests.
4. A short implementation report, suggested `docs/IMPLEMENTATION_REPORT_2026-09-08.md`, with actual results rather than copied baseline claims — including the recorded Python and Expat versions, corpus timing, the recorded spend figure, and a decision table for §7.

Do not deliver speculative code for conditional packages. The desired result is a safer parser, honest run diagnostics, and an evidence-based path to optimization that may correctly end in no optimization at all.

## 12. References

Recheck provider contracts before implementing live accounting or evaluation.

- [Expat XML security](https://libexpat.github.io/doc/xml-security/): amplification protections and the limits of a small entity-expansion reproduction. Consulted 2026-09-08.
- [`xml.parsers.expat` — `StartDoctypeDeclHandler`](https://docs.python.org/3.11/library/pyexpat.html#xml.parsers.expat.xmlparser.StartDoctypeDeclHandler): fires as Expat begins the document-type declaration, which is the rejection point §4.2 relies on. Consulted 2026-09-08.
- [Python XML processing and security](https://docs.python.org/3/library/xml.html): runtime and parser-version considerations. Consulted 2026-09-08.
- [Anthropic prompt caching](https://platform.claude.com/docs/en/build-with-claude/prompt-caching): usage categories, minimum prefix sizes, TTL behaviour. Consulted 2026-09-08.
- [Anthropic pricing](https://platform.claude.com/docs/en/about-claude/pricing): dated rates for optional cost estimates. Consulted 2026-09-08. Do not hard-code rates into the formatter.

The original review brief in the owner's Downloads remains contextual source material. Its instructions, branch name, measurements, and proposed optimizations are not independently verified authority; §7.2 and the §8 table record where it is wrong.
