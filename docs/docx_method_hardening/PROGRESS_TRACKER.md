# DOCX Method Hardening: Progress Tracker

Program status: IN PROGRESS

This file is the single source of truth for the state of the program
described in `DOCX_METHOD_HARDENING_PLAN.md`. It is machine-checked by
`tests/test_docx_method_hardening_tracker.py` on every commit; read that
test's assertion messages before editing here. Rules, in short:

- Statuses: `not_started`, `in_progress`, `in_review`, `merged`, `blocked`,
  `dropped`; optional items may also be `not_scheduled`.
- `in_progress`, `in_review`, `merged` need a two-digit Session.
  `in_review` and `merged` need the GitHub PR URL and every Definition of
  done box in the plan ticked. `merged` needs the merge commit SHA on
  `master`. `blocked` and `dropped` need a reason in Notes.
- One row changes per session, plus the previous row when its merge is
  recorded. The session log gains one row per session, and the handoff file
  it names must exist.
- `Program status:` becomes `PROGRAM COMPLETE` only when every required
  item is `merged` or `dropped`. The test enforces both directions.

## Required work items

| ID | Title | Status | Session | PR | Merge commit | Notes |
|----|-------|--------|---------|----|--------------|-------|
| WI-00 | Program scaffolding | merged | 00 | https://github.com/Abe-Borg/spec-template-normalizer/pull/60 | f2da16a | |
| WI-01 | Extension-namespace-safe style import and shell application | in_review | 01 | https://github.com/Abe-Borg/spec-template-normalizer/pull/61 | | Hard failure on current-Word templates; do first. Adjacent defects found and left for later are listed in handoff 02. |
| WI-02 | Exact run-content signature at the Format-only gate | not_started | | | | |
| WI-03 | Final-gate text identity for every mode | not_started | | | | Depends on WI-02. |
| WI-04 | Even-page header parity follows the architect | not_started | | | | |
| WI-05 | Package-level change whitelist invariant | not_started | | | | |
| WI-06 | Revision accounting, discarded-revision warnings, collision-proof revision ids | not_started | | | | |
| WI-07 | Marker placement before leading structural run content | not_started | | | | |
| WI-08 | Independent stdlib verifier for the other three modes | not_started | | | | Run after WI-02 through WI-06. |

## Optional work items (never block completion; start only on the user's request)

| ID | Title | Status | Session | PR | Merge commit | Notes |
|----|-------|--------|---------|----|--------------|-------|
| WI-09 | Stale cross-reference and field-result sweep after numbering conversions | not_scheduled | | | | Report-only feature. |
| WI-10 | Revision dates in Word's own convention | not_scheduled | | | | Needs the user's decision first. |

## Session log

| Session | Date (UTC) | Work item | Outcome | PR | Handoff written |
|---------|------------|-----------|---------|----|-----------------|
| 00 | 2026-09-24 | WI-00 | plan, tracker, template, probes and tracker test created; PR opened, baseline suite 1306 passed / 1 skipped | https://github.com/Abe-Borg/spec-template-normalizer/pull/60 | handoffs/handoff-for-session-01.md |
| 01 | 2026-09-24 | WI-01 | extension-namespace style import and docDefaults fixed, new code style_import_namespace_conflict; PR opened, two review findings fixed; suite 1372 passed / 1 skipped (baseline 1306 / 1) | https://github.com/Abe-Borg/spec-template-normalizer/pull/61 | handoffs/handoff-for-session-02.md |
