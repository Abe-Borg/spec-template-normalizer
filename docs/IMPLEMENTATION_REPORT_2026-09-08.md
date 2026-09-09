# Specification Formatter: implementation report

**Plan:** `docs/IMPLEMENTATION_PLAN_2026-09-08.md`, removed once the work
closed. The parts still load-bearing are reproduced here, so this report
stands alone; the full plan is recoverable from history when the reasoning
behind a decision is wanted:

```bash
P=docs/IMPLEMENTATION_PLAN_2026-09-08.md
git show "$(git rev-list -1 HEAD -- "$P")^:$P" > "$P"
```

(`git log -- "$P"` only lists the commits; recovering the file needs the
blob from the deletion commit's parent, which is what `git show` above
does.)
**Implemented:** 2026-09-08 to 2026-09-09
**Baseline at start:** `b156679` (suite: 1049 passed, 3 skipped)
**Final:** `49c73cc` (suite: 1115 passed, 3 skipped)
**Outcome:** the unconditional programme (W1, W2-reduced, W4) shipped; the spend
gate closed W3 and W5–W8. **Implementation is complete; acceptance is
not** — the Word inspection the plan required for W1 has not been done (see §5).

This report records what was actually measured and decided, not what the plan
predicted. Where the two differ, the difference is the point.

## 1. What shipped

| Package | PR | Result |
|---|---|---|
| W1 — encoding-independent XML declaration rejection | [#42](https://github.com/Abe-Borg/spec-template-normalizer/pull/42) | merged |
| W2 (reduced) — architect usage recorded; observed usage preserved on failure | [#43](https://github.com/Abe-Borg/spec-template-normalizer/pull/43) | merged |
| W4 — documentation corrected against the implementation | [#44](https://github.com/Abe-Borg/spec-template-normalizer/pull/44) | merged |

W0 was folded into W1 as ordinary verification, as the revised plan specified.
No evaluation tooling was built (see §4), so the plan's item about describing
that tooling is not applicable.

## 2. Measured results

Recorded because the plan asks for actual numbers rather than repeated
baseline claims.

**Runtime and parser versions.** The original bypass was reproduced on
**Python 3.11.15 / Expat 2.6.1** (the Windows CI version) — closing the
environment gap in the planning review, which had only Python 3.14.6 / Expat
2.8.1. The CI step added in W1 exercises the guard on the 3.10 floor, which
runs **Python 3.10.21 / Expat 2.8.2**. Rejection therefore holds across two
Expat majors, which matters because the fix depends on Expat's own
`StartDoctypeDeclHandler` rather than on anything in Python.

**Guard cost.** Measured over the sanitized corpus regression's real payloads
— 109 `parse_untrusted_xml` calls, 0.29 MB total, median part 943 B, p90
3.2 KB, largest 31 KB — parse time went **6.55 ms → 8.86 ms**, so **+2.32 ms
per corpus run**, about 0.5% of that run's wall clock. An earlier synthetic
58 KB microbenchmark was roughly twice the largest real part and ~60x the
median; the corpus figure supersedes it.

**Suite.** 1049 → 1115 passed, 3 skipped. 66 tests added.

**Engine digest.** `dfffda5580c73d7d` → `cdfe7148986b940e`, from W2 touching
root `llm_classifier.py`. This invalidated cached architect profiles once, at
the cost of one fresh analysis per template. W1 and W4 touched no covered
file. The conservative digest was kept, per plan invariant 7.

**Dependencies.** None added. `xml.parsers.expat` and `threading` are
standard library, so `requirements.txt` and `THIRD_PARTY_NOTICES.md` are
untouched.

## 3. Defects found

Eleven. **Five were already recorded in the revised plan's evidence table before
implementation began** — the UTF-16 bypass, the stale-declaration
mojibake, both usage-accounting gaps, and the verbosity-filtering hazard. Six
surfaced during implementation, from the plan's encoding matrix, an adversarial
encoding sweep, and automated review.

An earlier draft of this report said the plan knew about one. That was wrong,
contradicted by the table below and by this section's own narrative, and it
flattered the implementation at the planning work's expense. The planning
review found more than a third of what was ultimately fixed, and the
correction is the point of writing this down.

| # | Defect | Found by | Severity in practice |
|---|---|---|---|
| 1 | UTF-16 part smuggles `<!DOCTYPE` past the ASCII byte scan; entity expands | the plan | the reason W1 existed |
| 2 | Decoded `str` parsed through a stale declaration: `windows-1252` turns `é` into `Ã©` **with no error**; `utf-16` fails outright | review of the plan, then recorded in it | **silent document corruption** |
| 3 | Declared codec Python lacks raises bare `LookupError` — not a `ValueError`, so it escapes every caller handling `UntrustedXmlError` | adversarial sweep | unhandled crash on malformed input |
| 4 | Declared multi-byte encoding Expat refuses raises bare `ValueError` with no part name | adversarial sweep | unhandled crash on malformed input |
| 5 | `prepare_xml_text_for_utf8` rewrote declaration-shaped text **anywhere** in a part, including inside CDATA — on paths that write to disk and into a published output part | review | silent edit to published document content |
| 6 | Architect response usage never read; refusals and output-limit responses recorded as free | the plan (W2's premise) | untruthful cost reporting |
| 7 | Target usage published only on success; refusal or merge failure discarded it | the plan (W2's premise) | untruthful cost reporting |
| 8 | `prepare_template_profile` dropped `Phase1Result.usage`; init-failure path ignored `exc.observed_usage` — so `format_specifications()` published **no** architect tokens at all | review | W2 inert at the canonical entry point |
| 9 | Target usage attached only to an INFO event, so it vanished at `warning`/`error` verbosity | **the plan named it**; review caught the first W2 commit ignoring it | cost reporting depended on log level |
| 10 | `usage_complete` derived from "at least one field", so a response missing `output_tokens` reported a short total as final | review | overstated completeness |
| 11 | Deterministic-only target returned no usage snapshot, making a free target indistinguishable from unavailable telemetry | review | ambiguous accounting |

Two patterns are worth carrying forward.

**Conversion boundaries drop fields silently.** Defects 8 and 11 are the same
shape: a value exists upstream, a conversion doesn't carry it, and nothing
fails — the feature simply goes inert. Nothing in the type system or the test
suite objected. This is why usage now travels on explicit fields through
`Phase1Result` → `TemplateProfile` → `BatchResult` → `TargetFormatResult`
rather than being read back out of diagnostics events.

**A named hazard is not a handled one.** Defect 9 was written down in the plan
before any code was touched, and the first W2 commit still shipped it —
copying the warning into `CLAUDE.md` while implementing the thing it warns
against. Knowing about a failure mode in advance did not prevent it; only
someone re-checking the implementation against the stated requirement did.

**Descriptions drift faster than code, and nothing fails when they do.** W4 — a package whose entire purpose is correcting inaccurate
descriptions — introduced two inaccurate descriptions of its own, and its PR
body carried a corrected claim in its uncorrected form until that was caught
too.

## 4. The spend gate

**Decision: spend is small. W3, W5, W6, W7 and W8 are closed.**

The owner inspected the provider's existing usage reporting and judged
architect and target classification spend immaterial. That is the answer the
gate exists to obtain, and it retires the optimization programme without
building anything to measure it.

The figure itself is deliberately **not** recorded in this repository. The
plan asked for it, but the plan did not account for this repository being
source-available under PolyForm: a real API spend figure is the owner's
private financial information and does not belong in a public file. The
decision is what the record needs; the number lives with the owner.

W2-reduced shipping regardless was the correct call and is unaffected by this
result. A run that reports a refused target as free is wrong whatever the
spend turns out to be, and that is now fixed.

### Decision table for the conditional packages

| Package | Decision | Why |
|---|---|---|
| W3 — workload analyzer and evaluation harness | **Closed** | Built only to decide whether optimization is worth it. Spend says no, so the harness has no question left to answer. The corpus and gold-labelling programme it needed would have cost the owner's own adjudication time — the real constraint — for a decision already made. |
| W5 — payload reduction | **Closed** | Its acceptance gate required demonstrated token or cost savings worth the added complexity. With spend immaterial there is no saving worth the risk to classification quality. |
| W6 — target-classification cache | **Closed** | No measured repeated work to justify invalidation, corruption, concurrency and provenance obligations. Anyone revisiting it must read the correctness argument reproduced below first. |
| W7 — architect response replay cache | **Closed** | Was already conditional on evidence that repeated architect analysis is expensive. It is not. `ENGINE_SOURCE_DIGEST` stays conservative. |
| W8 — effort, model, chunking, deterministic-rule tuning | **Closed** | Each needed the adjudicated corpus from W3. Defaults are unchanged, which was the plan's position absent evidence. |

#### If a target-classification cache is ever reconsidered

Reproduced from the plan because it is the one conclusion that must not be
lost with it. The review brief that preceded this work proposed keying a
target-classification cache on the target's hash plus the available role
names.

**That key is incorrect, not merely coarse.** `role_specs` — the architect's
portable numbering patterns — flows into `build_phase2_slim_bundle` and drives
*deterministic* classification (`core/classification.py:503,569`). Two
architect templates with identical role **names** but different numbering
patterns produce different deterministic dispositions, a different unresolved
set, and a different request. Keying on target hash plus role names would
therefore serve a **wrong** cached classification, not a stale one — the
formatter would apply dispositions computed for a different template and
nothing would fail.

If it is ever built: cache model-derived dispositions for an exact
classification request plan; rebuild the target bundle and deterministic
dispositions every run; reapply every local validator, deterministic-override
check and coverage check before application; and never cache formatted DOCX
outputs or bypass source snapshots, run isolation, application policy,
package validation or publication. Cache identity must cover the target
source identity and paragraph-index universe, the exact template-derived role
definitions including numbering patterns and counter constraints, the actual
unresolved paragraph data and every piece of evidence sent to the model, the
system and user instructions with serialized role ordering and response
schema, provider and model identity with effort and output constraints, the
chunking and re-ask strategy, the preprocessing and merge semantics, and an
explicit cache contract version. The five-file architect engine digest does
**not** cover target preprocessing. Acceptance requires a two-template,
same-role-list regression that fails under the brief's proposed key.

None of these is rejected on merit. Each is closed because the evidence that
would justify it does not exist and, at this spend, is not worth generating.
If the workload changes — many more targets per run, or a much larger
template — the gate can be re-run and any of them reopened on the same terms.

## 5. Limitations of this work

Stated plainly, because the plan's acceptance checklist asks for it and
because an implementation report that only lists successes is not evidence.

**No Word visual inspection was performed — this is the one outstanding
acceptance item.** The plan required inspecting
representative output in Word for any W1 encoding behaviour that affects
resulting documents, and W1 does affect them: defects 2 and 5 both changed
what reaches a parsed tree or a published part. This work ran in a Linux
container with no Word and no representative specifications. The evidence
that exists is the corpus regression, package validation, the full suite, and
the format-only invariants — all of which are XML-level. **This is an open
acceptance item for the owner, not a completed one.** The realistic cases to
look at are a target whose `styles.xml` or `numbering.xml` carries a
non-UTF-8 declaration (defect 2) and, if one can be found, a part containing
literal `<?xml … ?>` text inside CDATA (defect 5).

**No live provider calls were made.** Every accounting test uses fakes. The
usage contract is verified against those fakes and against the real SDK's
attribute names, not against a live response. First real run should be
sanity-checked: `diagnostics.usage` in `run.json` should carry non-zero
architect counters on a fresh analysis and nothing on a reused profile.

**Defect 5's reachability is unproven.** The unanchored rewrite is a real bug
on a real write path, but Word does not normally emit CDATA in
`document.xml`, and ordinary `w:t` text is escaped and so unaffected. It was
fixed because a silent edit to published content is not worth leaving in on a
likelihood argument, not because a document was observed hitting it.

**End-to-end exploitability of defect 1 was not established**, and the plan
deliberately did not ask for it. Rejection of a prohibited declaration is a
stated contract; it was bypassable and now is not. Whether a practical denial
of service could have been mounted through it is a separate question that was
not answered, and modern Expat's amplification countermeasures argue against
assuming yes.

**One test was written loosely enough to pass against a defect.** The
original deterministic-only assertion was `assert "usage" not in result or
result["usage"].get("requests_attempted", 0) == 0`, which accepts both the
correct answer and the buggy one. It is tightened now, but the lesson
generalises: an assertion containing `or` is usually a test hedging about what
the contract is.

## 6. Recommended follow-ups

Not authorized by this report; listed so they are not lost.

1. **Do the Word inspection** described in §5 before treating W1 as fully
   accepted.
2. **Check `diagnostics.usage` on the first real run**, per §5.
3. **Consider whether the conversion-boundary pattern appears elsewhere.**
   Defects 8 and 11 were found in the accounting because it was under active
   development. The same shape — a field carried by one type and dropped by
   the next, failing nothing — may exist on paths nobody has recently
   exercised. This is a hypothesis, not a finding; it deserves a look, not a
   rewrite.
4. **Re-run the spend gate if the workload changes materially.** The closures
   in §4 are conditional on today's spend, not permanent judgements.

## 7. Provenance

Implemented across three reviewed pull requests, each verified on Windows
Python 3.11 and Linux Python 3.10/3.11 CI before merge. Automated review ran
on each and produced six of the eleven findings; every one was reproduced
against the code before being accepted, and two of its suggested remedies
were declined in favour of alternatives recorded on the review threads. Prompts, schemas, and validators were not modified: no contract in the
CSI role table, the bundle manifest, or the error-code set changed.
