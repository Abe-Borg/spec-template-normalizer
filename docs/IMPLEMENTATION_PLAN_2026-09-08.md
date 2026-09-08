# Specification Formatter: implementation and validation handoff

**Prepared:** 2026-09-08  
**Repository:** `C:\Github-Repos\spec-template-normalizer`  
**Baseline inspected:** `b66258a`  
**Status:** Implementation plan; application changes have not been made.  
**Audience:** Coding agents capable of independent investigation, implementation, adversarial testing, and integration review.

**Reading guide:** Start with sections 1-3 for decisions and invariants. Sections 4-8 specify the initial implementation. Sections 9-11 define conditional optimization work. Sections 12-14 contain integration checks, agent assignments, and required handoff artifacts.

## 1. Executive decision and scope

Implement the confirmed XML rejection fix, complete model-usage accounting, correct misleading engineering guidance, and build the evidence needed to make further optimization decisions. Preserve the current document-transformation contracts and model defaults while doing this work.

Do not treat the earlier review brief as an approved backlog. Several of its proposals have incomplete correctness arguments, and some suspected problems are already addressed in the implementation. This plan supersedes its recommendations for this work, while retaining its useful architectural context.

The initial deliverable is a small series of independently reviewable changes:

1. Encoding-independent rejection of prohibited XML declarations at the shared untrusted-parser boundary.
2. Accurate, privacy-preserving usage accounting for architect and target classification, including unsuccessful work.
3. Offline workload analysis and an evaluation harness that separates character savings, token savings, dollar estimates, and classification quality.
4. Documentation and validation updates reflecting the actual implementation and the limits of the evidence.

Payload changes are conditional on evaluation. Target-classification caching, architect-classification replay, model/effort changes, and broader deterministic rules are later decisions, not automatic follow-on implementation tasks.

### 1.1 Authorization and handoff boundaries

- The owner requested this Markdown plan. Creating this plan does not itself authorize this planning agent to implement application changes.
- An implementing agent must follow the owner's instructions in its own session. When assigned this plan for implementation, complete the unconditional work without repeatedly asking about routine design choices.
- Use offline fixtures and fake provider responses by default. Do not infer permission to send private specifications to a provider, incur evaluation charges, publish releases, or merge changes from this document alone. Follow any explicit authorization already supplied by the owner.
- Do not change the original review brief in Downloads. Preserve it as historical evidence; record corrections in repository documentation or an implementation report.
- Do not force the historical `claude/magical-mccarthy-3m54f8` branch from the brief. Inspect the actual checkout and follow the owner's current branch instructions. If a new branch is appropriate and no name was requested, use a `codex/` prefix.
- Do not clean unrelated files. At planning time the checkout contained three existing untracked `.pytest_tmp_header_token_final_edge*` directories. Their presence is not authorization to remove them.

### 1.2 Work package decisions

| ID | Work package | Decision | Dependency | Expected risk |
|---|---|---|---|---|
| W0 | Establish baseline and ownership | Required | None | Low |
| W1 | Harden the shared XML guard | Implement | W0 | Moderate: encoding compatibility |
| W2 | Complete model-usage accounting | Implement | W0 | Moderate: retry, failure, and concurrency paths |
| W3 | Workload measurement and evaluation harness | Implement offline capability; live evaluation needs suitable authorization | W0; consume W2 when ready | Low for offline tooling |
| W4 | Correct documentation and review assumptions | Implement | W1/W2 findings; can draft alongside them | Low |
| W5 | Reduce classification payload | Evaluate first; merge only supported changes | W2 + W3 | Moderate: model behavior |
| W6 | Target-classification cache | Deferred unless measured repetition justifies it | W2 + W3; settle W5 request representation first | High: stale classifications |
| W7 | Separate architect response cache from profile cache | Deferred; no initial implementation | W2 + evidence of expensive invalidations | High: replay semantics |
| W8 | Effort/model/chunk/deterministic-rule tuning | Deferred; one experiment at a time | W3 | High: classification quality |
| W9 | Integration and acceptance | Required | Every package selected for the release | Moderate |

Completion of W0-W4 and W9 is a valid, complete initial implementation. A measured decision to leave W5-W8 unchanged is an acceptable outcome, not unfinished work.

## 2. Evidence and corrections to preserve

The planning review read the brief, selectively inspected source at `b66258a`, and ran one small in-memory XML reproduction with bytecode writing disabled. It did not run the full test suite, process real specifications, measure production costs, or establish a complete security audit.

| Observation | Evidence at the inspected baseline | Consequence |
|---|---|---|
| UTF-16 bytes bypass the declaration guard | `core/untrusted_xml.py::parse_untrusted_xml`; small entity payload accepted and expanded to 16 characters | W1 is warranted as a rejection-contract fix |
| Decoded text containing the same declaration is rejected | Same in-memory reproduction | Raw-byte and text callers currently have different protection |
| Local reproduction used Python 3.14.6 / Expat 2.8.1 | Runtime output during the review | Do not confuse this environment with supported Python 3.10 or Windows CI's Python 3.11 |
| End-to-end exploitability is not established | Header/footer helpers accept raw bytes, but earlier shell helpers may decode and reject the same parts first | Trace actual paths; distinguish helper bypass from application-level exploit |
| Architect response usage is not accumulated | Root `llm_classifier.py::_call_api` reads final text and stop reason, but not response usage | Add accounting without changing classification instructions |
| Target usage is returned only on successful completion | Target classifier attaches `result["usage"]` after chunk collection/merge; runner reads it after return | Failed calls can consume tokens that never reach persisted accounting |
| Current diagnostics summary does not total tokens | `spec_formatter/diagnostics.py::DiagnosticsRecorder.summary` totals events and phase durations | Recording fields alone will not produce a run cost profile |
| Target classification depends on template role definitions | `build_phase2_slim_bundle(..., role_specs=...)`; `_numbering_role_candidates` and counter-conflict checks consume those definitions | Available role names alone cannot identify cache compatibility |
| Style lookup already has memoization | `_style_block_index` and `_find_style_numpr_in_chain` in `core/style_import.py`; `_parsed_style_elements` in `core/classification.py` | Do not add another cache based on the brief's performance suspicion |
| There are target and chunk thread pools | `pipeline.py` and target `core/llm_classifier.py` both instantiate `ThreadPoolExecutor` | Document the two levels accurately; do not rewrite concurrency as part of this work |
| The target prompt explicitly permits indentation evidence | `core/prompts/phase2_master_prompt.txt` also says "Do NOT reference formatting" | Clarify evidence versus output instructions before removing hints |
| Reported payload savings are character measurements | Brief measures JSON character lengths on a small constructed sample | Do not label them measured token, cost, or accuracy improvements |

The source-file names below are relative to the repository root. Line numbers from the historical brief are navigation aids only; locate current functions before editing.

### 2.1 Important qualification on the security finding

Modern Expat has entity-amplification countermeasures. Accepting a small internal entity proves that the application-level prohibition is bypassable; it does not prove unlimited expansion or a practical denial of service in the shipped application. Fix the explicit prohibition independently of the severity assessment. Document the actual Python/Expat versions used for both the reproduction and packaged validation.

Do not run an unbounded billion-laughs payload to demonstrate the issue. Tiny inputs and assertions that prohibited content never reaches the parsing operation are sufficient.

### 2.2 Additional correction discovered while preparing this plan

The brief describes retry behavior as though both classifiers share the same policy. At this baseline, target classification has a `retry-after`-aware transport policy, while root `_call_api` uses its own bounded exponential sleeps. W2 must instrument the actual policies without silently unifying or changing them. A retry-policy redesign would require its own concrete failure evidence and review.

## 3. Non-negotiable invariants

Read the current `CLAUDE.md` before implementation. Retain these contracts throughout all work packages:

1. Architect and target source files remain immutable. Process private snapshots and publish new outputs only.
2. In `format_only`, preserve target body text and effective numbering semantics. Do not describe this as requiring byte-identical entire DOCX packages; presentation and package serialization legitimately change.
3. In `csi_to_canadian`, retain the existing supported conversion boundary and rejection behavior.
4. Every classifiable paragraph has exactly one styled or ignored disposition. Keep deterministic-override rejection, coverage checks, and current out-of-scope treatment.
5. Keep source-derived styles, collision-safe style IDs, effective inheritance resolution, protected subtrees, and target numbering ownership.
6. Keep the single shared application path and the immutable application policy. Do not add a second formatter implementation for caching, evaluation, or error handling.
7. Keep strict profile validation and committed engine identity. Do not weaken cache invalidation to avoid the cost of this change series.
8. Keep per-run isolation, atomic output publication, source rechecks, and independent per-target failure handling.
9. Persist only code-defined identifiers, counts, booleans, safe timing values, and existing approved provenance in diagnostics. Never persist prompts, responses, paragraph text, API keys, HTTP bodies, or arbitrary provider objects there.
10. Preserve public and injected seams where practical. New optional parameters and additive result fields must have safe defaults. Do not require existing injected classifiers to emit usage.
11. Python 3.10 remains the source floor; Windows remains the primary platform. Do not introduce newer syntax or dependencies accidentally.
12. Do not change model defaults, effort, retry ceilings, concurrency limits, or chunk overlap in a telemetry/security patch.

## 4. W0 — establish baseline and assign ownership

### Tasks

1. Record the actual commit, current branch, working-tree changes, interpreter, Expat version, and installed dependency versions relevant to the selected work.
2. Read `CLAUDE.md`, the application policy, current test guidance, and any applicable repository instructions. Resolve conflicts in favor of the owner's current task instructions.
3. Run the existing suite in the implementing environment before editing. Record failures and skips rather than assuming the brief's `1049 passed, 3 skipped` is current.
4. Reproduce the tiny UTF-16 helper bypass using synthetic data. Record whether it still exists on the actual starting commit.
5. Identify injected classifier/analyzer/processor call signatures and tests that use strict fakes. These determine how telemetry must be threaded without breaking callers.
6. Assign one integration owner for shared files: `spec_formatter/pipeline.py`, `spec_formatter/diagnostics.py`, `README.md`, `CLAUDE.md`, and `engine_identity.py`.
7. Use separate commits or worktrees for independent work. Do not have agents concurrently overwrite the same shared file.

### Baseline commands

Run from the repository root using the intended interpreter. Resolve the interpreter explicitly if the shell's `python` is ambiguous.

```powershell
git --no-optional-locks status --short
git rev-parse HEAD
python --version
python -c "import pyexpat; print(pyexpat.EXPAT_VERSION)"
python -m pytest -q
```

Implementation tests may create test-owned temporary files. That is different from the read-only planning review; use the authorization and filesystem rules of the implementation session. Never remove unrelated temporary directories to get a cleaner baseline.

### Acceptance

- A concise baseline report identifies actual failures, skips, and environment limitations.
- The agent can explain the distinction between architect analysis, target classification, and shared application.
- No unrelated changes have been overwritten or cleaned up.

## 5. W1 — make prohibited XML rejection encoding-independent

### 5.1 Intended behavior

Every untrusted XML entry through `parse_untrusted_xml` must reject a DOCTYPE or ENTITY declaration before entity expansion, regardless of whether the caller supplied bytes or text and regardless of a supported input encoding. Valid supported OOXML must still parse with correct Unicode content. Unsupported or malformed encodings must fail predictably.

Prefer one central fix over hand-patching the currently known callers. A caller list cannot provide lasting protection when new raw-byte calls are added later.

### 5.2 Primary files

- `spec_formatter/style_application/core/untrusted_xml.py`
- `spec_formatter/style_application/core/ooxml_text.py`, only if existing decoding behavior needs a small shared correction
- `tests/style_application_regression/test_untrusted_xml.py`
- `tests/style_application_regression/test_ooxml_text.py`
- `tests/test_ooxml_text.py`

Trace and exercise callers in:

- `header_footer_importer.py::_remove_existing_hf_files`, `_rebuild_document_rels`, `_ensure_content_types`
- `phase2_invariants.py::validate_docx_package`
- `core/registry.py` bundle-artifact loading
- `docx_patch.py::validate_xml_wellformedness`
- `arch_env_applier.py` content-type and relationship preparation

### 5.3 Design requirements

1. Reuse the shared encoding-aware helpers rather than inventing a second BOM/XML-declaration decoder.
2. A suitable default design is: decode bytes using the shared policy; scan the resulting text for prohibited declarations; prepare a truthful UTF-8 declaration; parse only the checked, canonical representation.
3. The representation that is checked and the representation that is parsed must be equivalent. Do not decode and check one representation, then hand the original bytes back to an auto-detecting parser.
4. Guard against encoding disagreement. For example, BOM-less UTF-16 bytes can decode under an incorrect codec to a Python string containing literal NULs and later be auto-detected differently if re-encoded. Reject XML-invalid literal NUL characters and test these cases explicitly. Do not assume "decode first" alone is a complete proof.
5. Handle text inputs with an existing non-UTF-8 XML declaration correctly. Python `str` is already decoded; a stale declaration must not cause its newly encoded bytes to be misinterpreted.
6. Treat malformed bytes, unknown codecs, invalid Unicode, and unsupported encodings as errors. Keep `UntrustedXmlError` a `ValueError` subclass and retain useful part context without exposing source content through public diagnostics.
7. Do not silently broaden the parser's supported encoding set through arbitrary Python codec names. Establish the existing supported behavior, and either retain it safely or explicitly reject unsupported cases. Document intentional compatibility changes.
8. Preserve the current conservative declaration-screening policy for this patch. Do not combine it with a redesign to accept declaration-shaped literals inside comments or CDATA. Such behavior can be assessed separately.
9. Retain XML syntax validation. Canonicalization must not turn structurally malformed XML into accepted content or repair mismatched document structure.
10. Avoid a new runtime dependency unless the shared-helper approach cannot meet these requirements. If a dependency is necessary, explain the reason, test the frozen build, update direct runtime requirements, and regenerate notices through the existing process.

### 5.4 Required test matrix

| Case | Expected result |
|---|---|
| Valid UTF-8 XML bytes, with and without BOM | Correct root and text |
| Valid UTF-16 LE/BE with BOM and matching declarations | Correct root and text |
| BOM-less UTF-16 LE/BE with XML declarations | Correct decode or deliberate documented rejection; never bypass the guard |
| BOM-less UTF-16 without declarations | Explicitly test detection/rejection; never reinterpret unchecked content |
| UTF-32 variants supported by the shared reader | Correct canonical parse or documented rejection; never unchecked expansion |
| Existing supported declared single-byte encoding with non-ASCII text | Preserve characters |
| Already-decoded text with UTF-16 declaration | Correct safe handling |
| Tiny DOCTYPE plus internal entity in each supported encoding | `UntrustedXmlError` before expansion |
| External SYSTEM/PUBLIC declarations | Rejected without filesystem or network dereference |
| Existing uppercase/lowercase declaration-screening fixtures | Current rejection contract retained |
| Truncated encoded data, unknown encoding, literal NUL, invalid Unicode | Predictable wrapped failure |
| BOM/declaration disagreement | No unchecked alternate interpretation; supported policy documented |
| Entity/DOCTYPE-shaped content currently conservatively rejected | No accidental relaxation in this patch |
| Valid escaped text, comments, namespaces, and non-ASCII attributes | No unrelated regression |

Use tiny payloads. A focused spy asserting that a prohibited input does not reach the underlying parse call is justified here because "before parsing" is the security contract, not incidental implementation detail. Pair that assertion with behavior-level tests.

### 5.5 Reachability and integration evidence

1. Construct minimal synthetic packages with prohibited declarations in `word/_rels/document.xml.rels` and `[Content_Types].xml`.
2. Exercise both direct helper callers and normal application/package-validation paths.
3. Trace shell application order. Theme/settings/font-table helpers may already decode or reject parts before header/footer import. Record these earlier gates rather than claiming every listed raw-byte site is exploitable.
4. Include a valid package with non-UTF-8 parts as a positive control, not only malicious inputs.
5. Confirm failing targets do not publish partial DOCX outputs and source hashes remain unchanged.
6. Cover bundle-artifact parsing independently of the target package path. Integrity validation does not remove the need for parser safety.

### 5.6 Acceptance and rollback

- The original tiny bypass fails for bytes and text.
- Equivalent valid text retains its Unicode content across supported encodings.
- Existing error handling, package validation, both application modes, and corpus regressions remain intact.
- The implementation report separates helper-level rejection, reachable application paths, and any unresolved severity questions.
- The change is an isolated commit. If compatibility regresses, revise the canonicalization approach; do not silently restore acceptance of prohibited declarations.

## 6. W2 — complete usage accounting without changing classification behavior

### 6.1 Problem to solve

The current architect path has no response-usage accounting. The target path collects usage from completed responses but returns it only after successful chunk processing and merge. A refusal, exhausted regeneration, or a later failure can therefore consume tokens without those counts reaching the run artifacts. Existing run diagnostics summarize timing and events rather than token totals.

The required outcome is a trustworthy account of **observed usage**, including unsuccessful work, with explicit uncertainty when a provider response does not expose final usage. This is not a billing reconciliation system, and it must not invent zero-cost claims for interrupted requests.

### 6.2 Files and integration seams

| Layer | Files/functions | Required responsibility |
|---|---|---|
| Small shared usage contract | A focused module such as `spec_formatter/llm_usage.py` if warranted | Normalize known numeric fields, maintain thread-safe counts, expose immutable snapshots |
| Architect provider boundary | Root `llm_classifier.py::_call_api` | Record each stream attempt and final response usage before stop-reason handling |
| Architect regeneration and coverage patches | `_request_json_response`, `classify_document` | Carry one call-scoped collector through every request, including targeted patches |
| Architect orchestration | `phase1_pipeline.py::run_phase1`; `spec_formatter/template_analysis.py` | Preserve instruction schema and propagate telemetry through supported seams |
| Profile preparation | `pipeline.py::prepare_template_profile` | Distinguish fresh analysis, cache reuse, and failed analysis; retain observed usage on failure |
| Target provider boundary | Target `core/llm_classifier.py::classify_target_document` | Record every chunk, regeneration, and overlap re-ask; preserve counts if a worker or merge fails |
| Target orchestration | `style_application/batch_runner.py` | Publish telemetry in success and failure results without changing classification dispositions |
| Run artifacts | `pipeline.py`, `diagnostics.py` | Aggregate architect and targets exactly once; persist totals independent of log verbosity |

Use a small typed or explicitly validated contract. Do not build a general observability framework.

### 6.3 Accounting semantics to define before coding

| Field/concept | Definition |
|---|---|
| `requests_attempted` | Number of actual model stream/create invocations, including transport retries and grammar fallback requests |
| `responses_completed` | Number of requests for which the final message was obtained; includes refusals and output-limit responses |
| `responses_with_usage` | Completed responses with at least one valid recognized usage field; this alone does not prove complete accounting |
| `input_tokens` | Sum of valid provider-reported uncached input tokens under the provider's current usage contract |
| `output_tokens` | Sum of valid provider-reported output tokens, including billed thinking tokens where the provider includes them |
| `cache_read_input_tokens` | Provider-reported input tokens read from cache |
| `cache_creation_input_tokens` | Provider-reported cache-write input tokens |
| `usage_complete` | Whether all usage necessary for the reported scope is known; false for requests lacking required final counters |
| `requests_with_unknown_usage` | Requests for which final billable usage cannot be established from observed provider metadata |
| `transport_retries` | Additional transport attempts, distinct from JSON regeneration |
| `regenerations` | Re-requests for unusable JSON, validation failure, or output-limit responses under existing policies |
| `coverage_patch_requests` | Architect targeted requests for missing dispositions |
| `overlap_reask_requests` | Target overlap-conflict requests |

These are proposed names; the integration owner may align them with repository conventions before implementation. Publish a stable field table in the engineering guide. Do not leave ambiguous meanings such as whether "requests" means attempted or successful responses.

Additional rules:

- Missing usage is unknown, not zero. Accept zero only when it is reported or follows a documented provider contract, such as a known absent cache category in a complete current response.
- Reject booleans as integer counters. Ignore/report as incomplete malformed, negative, or unknown numeric fields without letting provider metadata alter document output.
- Do not add uncached input counts to an already inclusive total. Verify the provider schema before writing a cost formula.
- Token counting for the input-size guard is separate from model generation requests and billed response usage.
- If cache writes are later split by TTL, the aggregate cache-creation field and its breakdown overlap; do not sum both as separate usage.
- Preserve request purpose with fixed enums and numeric target/chunk/attempt identifiers. Do not use document titles, paragraph excerpts, response text, or arbitrary exception strings.
- Do not sum nested phase timings as run wall-clock time. Request duration sums and elapsed run duration describe different things under concurrency.

### 6.4 Recommended propagation design

1. Create one bounded, run-scoped accounting path, with independent architect and per-target collectors or scopes. No module-global mutable counters.
2. Thread an optional keyword-only collector/observer through production architect calls. Keep `classify_document` returning only the validated instruction object; do not insert telemetry into instruction JSON or its schema.
3. Capture final usage immediately after obtaining the final message and before throwing for `refusal`, `max_tokens`, or other stop reasons.
4. Record attempted requests even if opening or consuming the stream fails. Mark usage unknown if reliable final counts are unavailable. Do not manufacture a token estimate and present it as observed usage.
5. Use `finally` or an equivalent reliable handoff to retain already observed counts when parsing, validation, style derivation, bundle publication, chunk merging, or later application fails.
6. On the target path, wait for running chunk work to settle according to the existing executor behavior before finalizing that target's accounting. Otherwise a failing chunk can hide usage from siblings that finish later.
7. Choose one authoritative contribution per request. Existing target `result["usage"]` may remain as a compatibility summary, but it must not be added again to request-level totals.
8. Injected classifiers/analyzers that do not support telemetry must continue to work and report unavailable usage. Pass new arguments only through an explicit supported adapter/signature, not indiscriminately to every callable.
9. Never detect unsupported keyword arguments by catching `TypeError` from a classifier and calling it again. That can repeat real requests and conceal an internal bug.
10. Optional observers must not change successful document output or replace the primary classification exception. Do not expose a failing observer's arbitrary error text in logs. Keep programming errors visible through controlled tests and safe diagnostics.
11. If adding result fields such as a usage snapshot to `Phase1Result` or `BatchResult`, use additive defaults. A successful result field alone is insufficient for the failure path; maintain the independent collector/handoff.

### 6.5 Run summary and failure artifacts

- Add a documented, additive usage summary under `run.json` diagnostics, grouped by architect/target stage and model. Include total observed counters and completeness status.
- Preserve per-target usage in the existing target audit/diagnostics path, including failure results.
- Account for architect failure in the initialization-failure artifact path before any target starts.
- Report a reused architect profile as no new architect request for this run. Do not count the cost of its original creation again.
- Report deterministic-only target classification as zero model requests with known zero usage for that stage.
- Keep usage totals available at `warning` and `error` log levels. `DiagnosticsRecorder.summary()` currently uses stored, filtered events; merely adding INFO fields will lose totals when those events are suppressed. Accumulate validated usage independently of verbosity, or otherwise explicitly preserve this accounting contract.
- Keep existing event count/log consistency assertions valid. Do not pretend suppressed usage events were written to `diagnostics.jsonl`.
- Distinguish a complete observed total from a lower bound. Runs with unknown requests must be labeled incomplete.
- Keep dollar estimates out of the formatter's initial runtime change. W3 can calculate optional estimates from a dated price table without hard-coding prices into production logic.

### 6.6 Required tests

Use provider fakes; no paid calls in the normal suite.

| Scenario | Assertion |
|---|---|
| One architect response | Exact counters and one attempt/completion |
| Architect malformed JSON then valid JSON | Both responses counted once |
| Architect output limit then regeneration | Output-limit usage retained |
| Architect terminal refusal | Usage retained; no extra request |
| Architect targeted coverage patch | Initial and patch requests included |
| Structured-output compiler fallback | Attempts distinguished; usage not invented for a response without counters |
| Transport failure then success | Attempts and unknown-usage status truthful |
| Final stream interruption | Failure remains primary; accounting marked incomplete |
| Fresh analysis followed by style/bundle failure | Observed usage reaches failed run artifacts |
| Reused profile | No historical creation usage charged to current run |
| Strict legacy injected classifier/analyzer fake | Existing signature and return shape still work |
| Multiple target chunks | Totals equal independently known fake responses |
| Target refusal or exhausted regeneration | Observed usage survives target failure |
| One chunk fails while another finishes | Late sibling usage retained; no double counting |
| Overlap re-ask | Additional request included |
| Merge/coverage failure after responses | Usage retained even though no classifications return |
| Deterministic-only target | Zero model requests; client not constructed |
| Missing/partial/malformed usage fields | No crash; no false completeness claim |
| Parallel independent runs | No cross-run leakage of counts |
| INFO/DEBUG/WARNING/ERROR verbosity | Same usage totals; event logs still obey verbosity |
| Secret/body-text-shaped provider metadata | Excluded from every persisted artifact |
| Same usage exposed through callback and compatibility result | Exactly one contribution |

Primary test files:

- `tests/test_llm_classifier_safety.py`
- `tests/test_phase1_pipeline.py`
- `tests/test_unified_pipeline.py`
- `tests/test_diagnostics.py`
- `tests/style_application_regression/test_llm_classifier.py`
- `tests/style_application_regression/test_batch_runner_failure_diagnostics.py`
- `tests/style_application_regression/test_batch_runner.py`

Add a focused test module for the shared usage contract if a new module is created. Prefer assertions about totals, failure preservation, privacy, and behavior over snapshots of incidental helper calls.

### 6.7 Compatibility, versioning, and acceptance

- Changing root `llm_classifier.py` changes the committed engine fingerprint. Recompute it after integration and accept the existing cache invalidation rather than bypassing it.
- Keep `.phase1` instruction/audit semantics unchanged. Usage belongs to the current run, not the semantic identity of a cached profile.
- Additive diagnostics fields need documented compatibility handling; change manifest/audit version constants only if consumer assumptions actually change. Explain the decision in the PR.
- Verify any new module is discoverable in the frozen Windows build.
- Acceptance requires exact fake-provider totals on success and failure, stable classification results, no privacy regression, and no changes to request policy other than observation.

## 7. W3 — build workload measurement and an evaluation harness

### 7.1 Questions the evidence must answer

1. How many paragraphs actually reach the target model in representative work?
2. Which document structures account for unresolved paragraphs?
3. What are the measured request characters, provider-reported tokens, observed cache reads/writes, retries, and elapsed durations?
4. How much architect analysis is fresh versus reused, and how often does engine invalidation cause a new analysis?
5. Are repeated target classifications common enough to justify a persistent cache?
6. Do candidate payload or effort changes preserve the correct dispositions on difficult cases?

Do not answer these questions using synthetic examples alone. Synthetic cases establish correctness boundaries; a representative, appropriately authorized corpus establishes workload relevance.

### 7.2 Offline analyzer

Create a developer-facing tool, for example `tools/analyze_classification_workload.py`, with a tested importable core. It must use production extraction and slim-bundle logic rather than recreating numbering or classification rules. Proposed tooling paths are new deliverables, not existing commands to invoke during baseline setup.

Required behavior:

- Accept explicit target files and a validated architect profile or explicit fixture role definitions.
- Snapshot/extract through the existing bounded helpers into tool-owned temporary directories. Never unpack untrusted ZIPs with a new unchecked extraction loop.
- Build the production target slim bundle with the actual `role_specs` and available roles.
- Make no provider calls by default. A dry run must not transmit document text through a token-count endpoint either.
- Produce a machine-readable counts report and a short Markdown explanation of the sample and limits. Keep document text and private source paths out of the default shareable report.
- Use opaque fixture/document identifiers. Store any necessary mapping to owner files privately and separately, not in committed fixtures or public reports.
- Preserve hashes/identities according to the existing privacy policy; do not assume a hash is anonymization for arbitrary short text.
- Report skipped, failed, and empty documents as well as successful ones. Do not calculate savings only on the documents that happen to work.

For each target/profile pair report at least:

| Measurement | Definition |
|---|---|
| `classifiable_total` | Unique paragraph indices in the production classifiable universe |
| `deterministic_classified` | Unique classifiable indices assigned a role locally |
| `deterministic_ignored` | Unique classifiable indices ignored locally |
| `unresolved_sent` | Unique unresolved indices before chunk overlap |
| `unresolved_fraction` | `unresolved_sent / classifiable_total`; use not-applicable when the denominator is zero |
| `out_of_scope` | Separately counted unique excluded paragraphs; do not add overlapping subtree reports into the classifiable denominator |
| Request characters | Length of the actual serialized system/user/schema components, with each component identified |
| Chunk count | Production chunk count, including small/large edge cases |
| Paragraph presentations | Sum across chunks, including intentional overlap |
| Field contribution | Explicitly defined serialization-size comparison; never describe character attribution as token attribution |
| Candidate representation size | Characters under each evaluated wire representation, with no change to production defaults |

Check the arithmetic `deterministic_classified + deterministic_ignored + unresolved_sent == classifiable_total`. Use the production index/disposition contract, not a raw count of `filter_report` entries.

### 7.3 Corpus design

Start with existing sanitized fixtures, then add small purpose-built synthetic cases for missing structures. Use real owner documents only when their use is authorized, and sanitize any new fixture before committing it.

Cover at least these families:

- Fully automatic numbering, including numbering inherited through style chains.
- Typed markers with ordinary and deep hierarchy.
- Unmarked continuation prose and genuinely ambiguous headings.
- Repeated Canadian numeric markers where indentation and context matter.
- Roman/alpha ambiguity, short cross-references, and heading-like requirement sentences.
- Editorial/boilerplate content that should be ignored.
- Tables, drawings, text boxes, section breaks, and tracked/protected structures under current ownership rules.
- Mixed formatting, non-ASCII text, and supported XML encodings.
- Different architect profiles exposing the same role names but different numbering definitions.
- Large targets, chunk boundaries, gaps caused by deterministic filtering, and overlap disagreements.

As a planning target, aim for approximately 12-20 distinct target documents across several architect templates, plus synthetic adversarial cases. This is a starting collection goal, not a statistical guarantee or a reason to fabricate representative documents. Report actual coverage and any absent family.

Separate development and holdout documents by document/template family. Splitting neighboring paragraphs from the same document across those sets leaks context and exaggerates quality.

### 7.4 Gold labels and correctness criteria

- Gold labels require human/domain review for genuinely ambiguous material. A previous model response is not ground truth merely because it passed schema validation.
- Record natural CSI role, effective role under the template's allowed fallback policy, ignore disposition where appropriate, and any genuinely unresolved labeling disagreement.
- Keep paragraph indices tied to source document identity. A label file for different bytes must not be silently accepted.
- Evaluate exact effective-role agreement, classified-versus-ignored errors, per-role confusion, document-level failures, deterministic override failures, and exact coverage.
- Report rare/deep roles separately; high overall accuracy can hide a serious failure in a small category.
- Separate deterministic decisions from model decisions. A higher deterministic resolution rate is not evidence of higher precision.
- Treat every new wrong classified/ignored decision and every new invariant failure as a release blocker until understood. Do not trade correctness for a favorable average.
- Label uncertain gold cases separately and exclude them from claimed definitive accuracy until adjudicated; still include them in qualitative review.

### 7.5 Live evaluation runner, when authorized

Create a developer tool such as `tools/evaluate_classification.py` only as much as needed to run controlled comparisons. Reuse the production request construction, validation, merge, and retry behavior. Avoid a second classifier implementation.

Before live execution, require an explicit dataset, model/effort variant, maximum trial count, and spending/request limits appropriate to that session. An API key in the environment is not by itself an instruction to run a large paid experiment.

Required controls:

1. A no-network preview lists the trial matrix and estimated upper-bound exposure without dumping document content.
2. A live run records every attempt, retry, refusal, and incomplete-usage condition through W2 accounting.
3. Bound the number of trials and reserve conservatively for in-flight requests. Stop launching new work if the configured limit cannot accommodate another trial. Do not promise exact invoice enforcement where provider usage is unavailable.
4. Use the same dataset and validation rules for baseline and candidate. Change one variable at a time.
5. Record cache state and ordering. Cold and warm trials answer different questions; avoid giving one arm warm caches while presenting the comparison as otherwise identical.
6. Use repeated trials where stochastic variation matters. State the number of repeats and observed variability; do not claim a single equal score proves equivalence.
7. Keep raw prompts/responses out of default reports. Any diagnostic response retention must have a deliberate private location and retention policy suitable for the authorized corpus.
8. Do not run paid or credential-dependent evaluations in normal CI.

### 7.6 Cost calculations and reporting

Keep measured, estimated, and unavailable quantities distinct. A recommended report has separate columns for characters, observed token categories, estimated dollars, accuracy, failures, and elapsed time.

For a verified provider usage schema with disjoint token categories, the basic estimate is:

```text
estimated_cost =
    uncached_input_tokens * input_rate
  + cache_read_input_tokens * cache_read_rate
  + cache_creation_input_tokens * applicable_cache_write_rate
  + output_tokens * output_rate
```

Use rates per token, or divide per-million rates by one million. Include the rate source, retrieval date, model, provider/platform, and cache TTL. Do not apply a single cache-write multiplier if the experiment mixes TTLs.

When any request has unknown usage, label the dollar result incomplete or a lower bound. The provider's invoice remains authoritative. Character estimates must not be substituted into the observed-token columns.

Record saved API usage separately from developer effort, maintenance cost, and end-to-end wall time. A representation that saves tokens but increases retries or misclassifications is not a win.

### 7.7 Tests and acceptance

- Offline analysis makes no network calls, preserves source hashes, and uses production bundle construction.
- Counts handle empty documents, deterministic-only documents, filtered gaps, duplicate report entries, and overlapping chunks correctly.
- Dataset labels are rejected when source identity or indices do not match.
- Report generation excludes document text and private paths by default.
- Cost math is tested against small hand-calculated examples, including unknown usage and overlapping aggregate/breakdown fields.
- Fake-provider experiments exercise trial limits and failure accounting without paid calls.
- The final evidence report states which decisions can be made and which remain unsupported. If representative owner material or live authorization is unavailable, deliver the complete offline capability and name that precise limitation; do not fabricate a cost profile.

## 8. W4 — align documentation with actual contracts and evidence

### Required edits

Update `CLAUDE.md` and `README.md` alongside the implementation. Keep the engineering detail; do not replace substantive guidance with a shorter generic summary.

1. Document encoding-independent declaration rejection, supported encoding behavior, and the distinction between parser-contract failure and demonstrated application exploitability.
2. Explain usage fields, observed versus unknown counts, failure accounting, profile reuse, deterministic-only work, and verbosity-independent totals.
3. Describe concurrency accurately: one public target-processing orchestration path, with an internal chunk pool in the target classifier. The target request semaphore should not be described as covering root architect calls unless the code actually does so.
4. Describe each classifier's actual retry behavior. Do not imply identical policies where they differ.
5. State that the current inherited-style lookup already uses caches. Do not add speculative performance claims without profiling.
6. Distinguish exact target body-text/numbering preservation from whole-package byte identity. Explain that ignored paragraph XML remaining unedited does not, by itself, prove unchanged rendered appearance when document defaults or the shell change.
7. Distinguish zero LLM involvement from correct classification. Deterministic rules require adversarial tests and precision checks too.
8. Explain that payload size measurements are not token or dollar measurements.
9. Document that prompt-cache reads can be affected by concurrency, minimum prefix size, TTL, and prefix identity. Zero reads alone do not identify which factor caused a miss.
10. Keep visual inspection of representative Word outputs in the validation checklist when classification, formatting, numbering, or shell behavior changes.
11. Describe the evaluation tooling as developer tooling, not a new mandatory step in the normal user's GUI workflow.

Do not edit production prompts merely to improve wording during W4. The target prompt's formatting-evidence ambiguity is a model-input change and belongs in W5 with evaluation. Documentation may identify the ambiguity before it is resolved.

### Reference brief corrections

The implementation report should explicitly close or qualify these historical claims:

| Brief claim | Correct treatment |
|---|---|
| Omit empty/false fields with zero information loss | Define field defaults and evaluate; absence can differ from explicit false/null |
| Target hash plus role names is enough for target cache reuse | Include template-dependent role definitions and exact classification inputs/contracts |
| Style numbering lookup needs memoization | Existing caches already cover the suspected path; profile before further work |
| Zero cache reads means prefix invalidation | Check concurrency, warm-up, TTL, and minimum size as well |
| Only one thread pool exists | Distinguish one public runner from internal chunk concurrency |
| Classification is intrinsically a poor use of high effort | Treat as a workload-specific hypothesis, not a general fact |
| More deterministic resolution always improves quality | Require precision evidence and adversarial coverage |
| Payload reductions are measured token savings | They are character reductions until measured with the actual tokenizer/provider |
| Review verdict must have no hedging | State a clear verdict with explicit coverage and uncertainty |

### Acceptance

- Documentation matches shipped behavior and selected work packages.
- Current limitations and deferred decisions are explicit.
- API/model/pricing assertions that remain in documentation have current primary-source references.
- No claim of full-suite success, production savings, or Word validation is copied from the brief without being reproduced for the actual implementation.

## 9. W5 — evaluate payload improvements before changing the default

### 9.1 Decision gate

Proceed only after W3 can compare actual production inputs and W2 can measure request usage. A candidate may be prototyped behind an evaluation-only option. Do not ship an unevaluated representation as the default merely because it is smaller.

Keep the full slim bundle as the local authority for validation and audit. Build a separate wire projection for the model; do not delete fields from the internal bundle or mutate shared paragraph dictionaries in place.

### 9.2 Candidate sequence

**Candidate A: remove duplicated role-list text from the user message.** The same role list already exists in the system block, but verify every production request/retry/re-ask path still includes it. Measure the modest saving and check difficult fallback cases. Treat even this as a prompt change, not a guaranteed behavioral no-op.

**Candidate B: omit only fields with explicitly defined default semantics.** Use a per-field allowlist and documented default table. Never use a generic recursive truthiness filter: it can drop `False`, numeric zero, paragraph index zero, empty-but-meaningful collections, or counter values. Preserve all non-default numbering evidence and semantic-conflict information.

**Candidate C: remove duplicated neighbor text only when its exact context is present.** Use paragraph identity and the actual wire text, not merely equal strings. A neighbor may be absent because deterministic filtering removed it, or its body and neighbor-preview truncation may differ. Preserve context at chunk boundaries, gaps, and overlap re-asks whenever equivalent context is not present.

**Candidate D: resolve formatting-evidence instructions.** Separate these two concepts explicitly: formatting can be evidence for hierarchy where appropriate; the model returns roles/dispositions rather than formatting commands. Test indentation-dependent Canadian and unmarked-prose cases. Do not delete `pPr_hints`, `rPr_hints`, or indentation wholesale based on the ambiguous phrase "Do NOT reference formatting."

Evaluate candidates independently before combining them. Keep the existing paragraph-text cap, output schema, effort, model, overlap, and retry policy fixed in these comparisons.

### 9.3 Files and tests

Likely files:

- Target `core/llm_classifier.py::_build_user_message`, `_system_blocks`, chunk/re-ask construction
- `core/prompts/phase2_master_prompt.txt` and `core/prompts/phase2_run_instruction.txt` if actual prompt changes are selected
- `core/classification.py` only if a small explicit context identity is needed in the internal bundle
- Target classifier/bundle tests and W3 evaluation cases
- `README.md`, `CLAUDE.md`, and request/provenance identity code where applicable

Required tests include preservation of paragraph index zero, explicit false flags, non-empty numbering metadata, inherited numbering, default omission semantics, non-mutating projection, filtered neighbor gaps, unequal preview truncation, duplicate paragraph text at different indices, chunk boundaries, and overlap re-asks.

### 9.4 Acceptance gate

Merge only candidates with:

- No new invariant violations or classified/ignored errors in the regression corpus.
- No unexplained new role errors in adjudicated holdout cases, with sample size and repetitions stated.
- A demonstrated reduction in observed tokens or end-to-end cost worth the added complexity; character reduction alone is insufficient for a cost claim.
- No material increase in retries, refusals, incomplete responses, or document-level failures.
- Updated prompt/request fingerprints and explicit representation identity if a later cache depends on them.
- A straightforward rollback to the previous request representation, preserving test evidence.

An inconclusive evaluation means keep the production representation unchanged. Record the result rather than broadening the experiment opportunistically.

## 10. W6 — target-classification caching, only if justified

### 10.1 Why this is deferred

The brief's suggested key omits template-dependent inputs. A persistent cache also adds invalidation, corruption, retention, concurrency, privacy, and provenance obligations. Implementing it before measuring repeated unresolved work could add substantial risk for little saving.

Authorize implementation only when W3 shows recurring identical classification work with meaningful observed cost or latency, and when the request representation is sufficiently stable. If deterministic classification resolves nearly everything, defer the cache.

### 10.2 Safe design direction if selected

Cache model-derived classifications at a clearly defined boundary, preferably the LLM-only dispositions for an exact classification request plan. Rebuild the current target bundle and deterministic dispositions every run. Reapply current local validators, deterministic-override checks, and exact coverage checks before application.

Do not cache formatted DOCX outputs or bypass source snapshots, run isolation, application policy, package validation, or output publication.

The cache identity must cover all behavior-affecting inputs, including:

1. Target source identity and paragraph-index universe.
2. Exact template-derived role definitions relevant to classification, including numbering patterns, provenance, and counter constraints. Same role names are not sufficient.
3. Actual unresolved paragraph data and all context/formatting/numbering evidence sent to the model.
4. System and user instructions, role ordering as actually serialized, response schema, and wire-projection version.
5. Provider/model identity, thinking/effort, output constraints, and any request options that affect classification.
6. Chunking/overlap/re-ask strategy and the effective request plan, if caching combined document-level results.
7. Classification preprocessing, deterministic rules, validation/merge semantics, and their compatible version/fingerprint.
8. An explicit cache contract version and an immutable implementation identity suitable for a frozen build.

Hash a canonical exact request representation plus the relevant local semantic contract. A source hash and model name alone are insufficient. Do not assume the existing five-file architect engine digest covers target preprocessing changes.

Conversion mode may be omitted only after a test proves that, for the same source and profile inputs, it does not change the cached classification request or its interpretation. Reuse across modes is an optimization to prove, not an assumption based on a missing function parameter.

### 10.3 Storage and behavior requirements

- Use a versioned private cache namespace with bounded entry count/bytes and an explicit pruning policy.
- Store only the validated data required for reuse. Do not persist prompts, paragraph text, raw responses, arbitrary notes, or free-form provider metadata by default.
- If ignored reasons are model-authored strings, explicitly settle safe storage/normalization semantics before caching them; do not quietly change the classification contract to fit the cache.
- Write atomically. Never publish partial entries. Test concurrent readers/writers and crash remnants.
- Reject malformed, oversized, incompatible, or checksum-invalid entries. Treat ordinary cache corruption as a safe miss and perform fresh classification under normal policy; do not reuse partially validated data.
- Checksums detect corruption; they do not authenticate a hostile writer with access to the cache. State the local cache trust boundary honestly.
- Disable reuse for injected classifiers lacking a stable, explicit identity.
- Report current-run cache hits/misses and zero new API requests for hits. Retain historical provenance separately if needed; never charge historical token usage to the current run.
- Failure to store a new cache entry should not invalidate an otherwise correct formatting output unless an existing explicit product contract requires it.
- Retain fresh per-run audits and manifests, even when classification is reused.

### 10.4 Required cache tests

| Change/scenario | Expected behavior |
|---|---|
| Same target, exact profile semantics and request contract | Reuse validated classifications |
| Same role names, different template numbering patterns | Cache miss |
| Same target, changed context/preprocessing or deterministic rule | Cache miss |
| Changed system/user prompt, schema, model, effort, or wire projection | Cache miss |
| Changed chunk plan or merge contract | Cache miss for combined results |
| Switched mode with proven identical classification inputs | Hit may be allowed; application still follows selected mode |
| Valid JSON containing wrong, duplicate, missing, or unknown indices | Reject entry |
| Cache result attempts deterministic override | Reject entry |
| Corrupt/partial/oversized entry | Safe miss; no output corruption |
| Concurrent runs and pruning | No partial reads, lost selected entries, or cross-run writes |
| Injected classifier without stable identity | No persistent reuse |
| Cache hit and later application failure | Fresh failed run artifacts, truthful zero new model requests |

### 10.5 Acceptance

The cache must reduce measured repeated work, preserve every validator, and pass a two-template same-role-list regression that would fail under the original brief's proposed key. If that correctness argument or the economic case remains incomplete, do not implement the cache.

## 11. W7 and W8 — deferred structural and quality changes

### 11.1 Architect response replay cache

Do not remove `ENGINE_SOURCE_DIGEST` from profile validation or replace it with a comment-stripping heuristic. The coarse digest is deliberately conservative.

A future two-tier cache requires evidence that repeated architect analysis due to engine changes is materially expensive. Its design must distinguish:

- The exact model request, including the derived slim bundle and request schema.
- The model response before deterministic repairs.
- Current deterministic repairs, coverage patching, validation, and profile derivation.
- The current profile's complete engine identity and shell-capture behavior.

The source file hash alone does not identify the model input: a change to the slim-bundle builder can change the request for identical DOCX bytes. Stored post-repair instructions cannot safely stand in for a pre-repair response after repair logic changes. Targeted coverage patches further complicate replay.

If this work is selected later, specify a versioned replay transcript or equivalent validated raw-stage representation, its private-data handling, bounded storage, and how current validators rerun. Add tests for changed bundle construction, repair order, patch logic, prompt/schema, and invalid stored results. Continue deriving and validating profiles under the current full engine identity. Do not implement this tier merely to avoid one cache invalidation caused by W2.

### 11.2 Effort or model changes

Keep current defaults in the initial series. If W3 justifies experimentation, compare current settings against one lower-effort setting before considering a different model. Use the same adjudicated corpus and request representation, and account for retries and cache effects.

No claim that classification inherently benefits or fails to benefit from reasoning effort should substitute for workload evidence. Do not disable thinking, introduce confidence-based cascades, or create a GUI model selector as part of this work.

### 11.3 Deterministic-rule expansion

Only add a rule for a concrete unresolved pattern that the corpus establishes and that has a defensible structural discriminator. Include close negative cases, especially cross-references, heading-shaped requirements, Roman/alpha ambiguity, and template-specific numbering conflicts.

The acceptance metric is correct resolution with no known precision regression, not simply a lower unresolved percentage. Keep ambiguous material unresolved when it cannot be proven locally.

### 11.4 Chunking and prompt-cache warm-up

Do not enlarge chunks because a provider advertises a larger context window. Measure request sizes, output headroom, regeneration exposure, latency, and overlap cost. The present character-based chunk budget is an estimate, not proof of provider token fit.

Do not serialize requests merely to create cache hits without measuring the latency tradeoff. Concurrent requests can miss before the first response begins; this is not by itself a correctness bug. Retain the current semaphore and executor behavior unless a separate experiment justifies a change.

### 11.5 Performance and maintenance exclusions

- Do not add another inherited-style cache without profiling current cache behavior.
- Do not split large modules or rename the root/package classifier modules solely because the brief calls them large or confusing.
- Do not bring back Batch API or retired public runners.
- Do not alter GUI settings, updater behavior, licensing, or broad CI platform coverage without a task-specific reason.
- Do not expand this project into a complete namespace/OOXML rewrite. New independent findings should receive their own reproduction, severity, scope, and owner decision where they exceed this plan.

## 12. W9 — integration, validation, and release readiness

### 12.1 Change sequence

Recommended PR/commit sequence:

| Sequence | Scope | Must stand on its own |
|---|---|---|
| PR 1 | W1 XML fix, adversarial tests, relevant documentation | Safe parser behavior; no telemetry or optimization changes |
| PR 2 | W2 usage contract, architect/target propagation, summaries, tests/docs | Correct accounting without prompt/request-policy changes |
| PR 3 | W3 offline analysis and evaluation tooling, W4 remaining corrections | No network by default; no changed production classification |
| PR 4, only if warranted | W5 selected payload change and its evidence | Explicit quality/cost gate met |
| Later independent work | W6/W7/W8 if selected | Separate evidence and compatibility review |

An integration owner may split PR 2 into the shared contract, architect propagation, and target/rollup integration, but must not present partial success-only accounting as the finished outcome.

### 12.2 Version and resource checklist

| Surface | When to update |
|---|---|
| `engine_identity.py::ENGINE_SOURCE_DIGEST` | Any covered root source changes; compute after final integration |
| `PIPELINE_VERSION` | If architect pipeline behavior/compatibility requires a version change; do not substitute it for the digest |
| Profile contract/manifest versions | Only if consumer-visible profile assumptions change; W2 should normally avoid semantic bundle changes |
| Run audit/manifest versions | Assess additive usage fields against readers; document compatibility decision |
| Prompt fingerprints / wire representation identity | Any selected W5 request change |
| `requirements.txt` | New or changed direct runtime dependency only |
| `requirements-dev.txt` | Necessary development-only dependency; prefer existing standard library and pytest |
| `THIRD_PARTY_NOTICES.md` | Regenerate through the existing build process if dependencies change; never hand-edit |
| PyInstaller spec/resource handling | New dynamically imported runtime modules or resources not covered by current collection |
| `README.md` and `CLAUDE.md` | Changed behavior, diagnostics, constraints, and validation requirements |
| JSON schemas | Only actual contract changes; do not add usage to semantic instructions to avoid plumbing |

The current PyInstaller spec collects `spec_formatter` submodules and explicitly names root modules. Verify the real build rather than assuming this covers every new import pattern.

### 12.3 Focused test groups

Use these existing paths as a starting point. Add the new work-package tests after they exist.

```powershell
python -m pytest -q tests/style_application_regression/test_untrusted_xml.py tests/style_application_regression/test_ooxml_text.py tests/test_ooxml_text.py tests/style_application_regression/test_docx_patch.py tests/style_application_regression/test_final_package_validation.py
python -m pytest -q tests/test_llm_classifier_safety.py tests/test_phase1_pipeline.py tests/test_diagnostics.py tests/test_unified_pipeline.py tests/style_application_regression/test_llm_classifier.py tests/style_application_regression/test_batch_runner_failure_diagnostics.py
python -m pytest -q tests/test_engine_identity.py tests/test_resources.py
```

After final integration:

```powershell
python engine_identity.py
python -m pytest -q
python -m pytest -q tests/test_sanitized_format_only_corpus.py
```

`python engine_identity.py` prints the computed digest; update the committed constant when required and rerun its focused test. Do not change the constant to suppress a failure without checking which covered files changed.

The full suite already includes the corpus test in normal collection. A separate corpus invocation is useful for an explicit acceptance record or the repository's required workflow; do not repeatedly rerun everything when no new changes or failures justify it.

### 12.4 Supported-runtime and Windows checks

1. Run the full suite on Windows/Python 3.11, matching the existing authoritative CI gate.
2. Run W1 and W2's relevant non-GUI tests on Python 3.10, not just an import check. The current Linux job installs runtime dependencies after its updater tests; a small focused step after that installation can cover the new shared-parser/usage behavior. Preserve pinned workflow actions and avoid unrelated CI redesign.
3. Record Python and Expat versions for parser results. Success on the local Python 3.14 runtime alone is insufficient evidence for the supported floor.
4. Confirm a frozen Windows app can import the new runtime code and load prompts as before. Use existing build/self-check procedures from `docs/RELEASE_WINDOWS.md`; do not invent release commands or publish a release automatically.
5. If a dependency or frozen-import arrangement changes, complete the existing packaging validation and notices generation before calling the change release-ready.

### 12.5 Document and visual acceptance

- Exercise both application modes with offline classifier seams and representative fixtures.
- Verify source hashes, text, numbering semantics, protected structures, collision-safe styles, relationships, package validity, and no partial output publication.
- For W2 alone, unchanged requests and output behavior plus regression tests are the primary evidence; telemetry does not require a new visual format standard.
- For any W1 encoding behavior or W5 classification change that affects resulting documents, inspect representative output in Word against the architect template and original target. Render every page of selected representative documents when formatting/shell behavior changes, as the repository guide requires.
- Use disposable output copies for Word inspection; do not save originals or treat Word's automatic field updates as part of the formatter's output.
- Check numbering/restarts, hierarchy depth, indents, headers/footers, page breaks, non-ASCII text, and content around filtered/protected paragraphs.
- Record renderer/Word version, font availability, selected cases, and any inspection limitations. XML invariants and visual review complement each other; neither replaces the other.

### 12.6 Final acceptance checklist

- [ ] W0 baseline and actual starting commit documented.
- [ ] W1 rejects prohibited declarations consistently before expansion across tested encodings.
- [ ] W1 preserves valid supported non-ASCII/encoded content and stable failure behavior.
- [ ] W2 accounts for architect and target requests, retries, refusals, patches, and late failures.
- [ ] Unknown usage remains explicitly unknown; no silent zero-cost claims.
- [ ] Current-run totals exclude historical cache creation and do not double-count compatibility summaries.
- [ ] Usage summaries remain available at all diagnostic verbosity levels.
- [ ] No document text, secrets, raw responses, or arbitrary provider metadata enter artifacts.
- [ ] Existing injected seams remain usable; no accidental duplicate request from compatibility handling.
- [ ] Model, effort, retry, concurrency, and formatting defaults remain unchanged unless separately evaluated and selected.
- [ ] W3 runs offline by default and reports honest denominators and uncertainty.
- [ ] Documentation and relevant versions/digests/resources match the final implementation.
- [ ] Focused and full tests pass in the required environments, with skips explained.
- [ ] Required Word/visual checks are complete, or the report precisely identifies why release acceptance is pending.
- [ ] Each conditional package has a clear implemented/deferred/rejected decision and evidence.
- [ ] Unrelated user files and prior outputs remain untouched.

## 13. Agent delegation and handoff instructions

This section describes how the owner or lead implementing agent can divide the work. It does not mean that agents were launched during preparation of this plan.

### 13.1 Suggested assignments

| Agent | Bounded assignment | Primary ownership | Do not edit concurrently |
|---|---|---|---|
| Security agent | W1 implementation, encoding matrix, call-path trace | Shared XML helper, decoding helper if needed, XML tests | Telemetry, production prompts, shared docs without coordination |
| Accounting agent | W2 usage contract and both classifier boundaries; coordinate integration seam | Usage module, root/target classifiers, dedicated usage tests | Shared pipeline/diagnostics unless assigned ownership |
| Evaluation agent | W3 offline analyzer, corpus manifest, fake-provider evaluation/reporting | Developer tooling, sanitized evaluation fixtures, tooling tests | Production classifier defaults or cache behavior |
| Integration/review lead | W0, shared propagation/rollup integration, W4, W9, final acceptance | Pipeline, diagnostics, shared docs, engine digest, focused CI additions | Agent-owned files until their changes are ready |

If fewer agents are available, implement sequentially in the same order. Parallelism is useful only for independent files and evidence gathering. Do not create multiple competing usage schemas or multiple extraction/classification implementations.

Before the accounting and integration agents code, agree on the collector/snapshot contract, completeness semantics, and the one authoritative aggregation path. Send short interface notes containing names/types/defaults and failure behavior. Reconcile changes through commits or explicit handoff, not by racing edits.

### 13.2 Per-agent completion report

Each agent must return:

1. The concrete problem and resulting behavior.
2. Confirmed evidence and any hypothesis disproved during implementation.
3. Files changed and the reason for each change.
4. Tests run, interpreter/platform, results, and relevant missing checks.
5. Invariants touched and how preservation was demonstrated.
6. Compatibility, fingerprint, dependency, schema, and documentation implications.
7. Remaining risks, uncertainties, and any conditional work deliberately deferred.
8. Commit(s) or a precise diff handoff for integration.

Do not call work complete because a function works in isolation if failure artifacts, compatibility seams, or required integration checks remain untested.

### 13.3 Independent integration review questions

The final reviewer should answer these from the diff and evidence, not from the implementation author's confidence:

- Can the XML parser interpret a different encoding/representation from the one screened for declarations?
- Are valid non-ASCII inputs preserved, and are unsupported encodings rejected consistently?
- Can a paid response occur before an exception that discards its observed usage?
- Can a worker finish after the accounting snapshot was finalized?
- Can the same request contribute through both a callback and a returned summary?
- Can a reused profile import historical usage into a new run's totals?
- Can logging verbosity remove cost information?
- Can an injected fake or custom callable cause a second real invocation during argument fallback?
- Did telemetry change prompt construction, model settings, retries, or validators unintentionally?
- Are any reported savings only character counts or estimates presented as observations?
- If payload changes were selected, were holdout quality, fallback roles, context gaps, and failed runs included?
- If caching was selected, does a same-role-list/different-template case invalidate correctly?
- Does the final documentation describe exactly what was implemented and tested?

## 14. Expected final handoff artifacts

Deliver the following for the initial implementation:

1. Reviewable implementation commits/PRs for selected work packages.
2. Updated `README.md` and `CLAUDE.md` with the new parser and accounting contracts.
3. Focused adversarial and accounting regression tests.
4. Offline workload/evaluation tooling with no-network defaults and example sanitized inputs.
5. A Markdown implementation report, suggested location `docs/IMPLEMENTATION_REPORT_2026-09-08.md`, containing actual results rather than copied baseline claims.
6. A workload/evaluation report when data is available, explicitly distinguishing measured, estimated, and unavailable results.
7. A short decision table stating which of W5-W8 were implemented, deferred, or rejected and why.

Do not deliver speculative code for every conditional work package. The desired result is a safer parser, trustworthy accounting, and an evidence-based path to optimization while preserving the formatter's document contracts.

## 15. Primary references and freshness

These references support specific technical qualifications. Recheck provider contracts before implementing live accounting or evaluation behavior; model/API details may change.

- [Expat XML security](https://libexpat.github.io/doc/xml-security/): amplification protections and the limits of interpreting a small entity-expansion reproduction. Consulted 2026-09-08.
- [Python XML processing/security](https://docs.python.org/3/library/xml.html): runtime/parser-version considerations. Consulted during the review on 2026-09-08; also check the documentation for the supported interpreter versions.
- [Anthropic prompt caching](https://platform.claude.com/docs/en/build-with-claude/prompt-caching): usage categories, minimum prefix sizes, cache availability after response start, and TTL behavior. Consulted 2026-09-08.
- [Anthropic pricing](https://platform.claude.com/docs/en/about-claude/pricing): dated rates for optional evaluation cost estimates. Consulted 2026-09-08. Do not hard-code these rates into the formatter as part of W2.

The original document `C:\Users\AbrahamBorg\Downloads\spec-template-nomalizer_REVIEW_BRIEF.md` remains contextual source material. Its instructions, historical branch name, measurements, and proposed optimizations are not independently verified authority.
