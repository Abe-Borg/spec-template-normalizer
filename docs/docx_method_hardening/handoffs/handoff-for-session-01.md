# Handoff prompt for session 01

Paste everything below the horizontal rule into the next session as its
first message. Written by session 00 on 2026-09-24 UTC.

---

You are continuing the **DOCX Method Hardening** program in the repository
`Abe-Borg/spec-template-normalizer`. Read these files completely before your
first tool call, in this order:

1. `CLAUDE.md` (engineering guide; every rule in it applies)
2. `docs/docx_method_hardening/DOCX_METHOD_HARDENING_PLAN.md`
3. `docs/docx_method_hardening/PROGRESS_TRACKER.md`
4. This prompt

## State at handoff

- Previous session: `00`, work item `WI-00: Program scaffolding`. Session 00
  wrote no engine code; it created the plan, the tracker, the handoff
  template, the two probes and the tracker test.
- Pull request: `https://github.com/Abe-Borg/spec-template-normalizer/pull/60`. Merge status when this prompt was
  written: `open`.
- Tracker rows changed by the previous session: `WI-00` created as
  `in_review`; session log row `00` added.
- Verification the previous session ran, with results:
  `python -m pytest tests/test_docx_method_hardening_tracker.py -q` passed;
  both probes ran from the repository root and reproduced their defects
  (`probe_style_import_w14.py`: `format_only` row FAIL with `unbound prefix`,
  `csi_to_canadian` row OK; `probe_format_only_gate.py`: 7 mutations
  ACCEPTED, control REJECTED). The full suite was run once on the unchanged
  tree; 1306 passed, 1 skipped (the GUI test skips without `tkinter`).
- Anything left unfinished, unexpected, or decided along the way: nothing
  unfinished. Two facts you should not have to rediscover: `master` is the
  default branch, and `tests/style_application_regression/test_engine_errors.py`
  bounds the number of error codes, so adding a code means raising that
  bound in the same PR.

## Your assignment: `WI-01: Extension-namespace-safe style import and shell application`

- Follow the **session start ritual** in the plan (section 3.2) first:
  confirm the PR above is merged; if it is not, stop and tell the user rather
  than building on a stale base. Record the merge in the tracker
  (`WI-00` -> `merged`, with the merge commit), then mark `WI-01`
  `in_progress` with session `01`.
- The plan section for `WI-01` is the specification. Its
  **Definition of done** checklist is what you must complete and tick.
- Key files: `spec_formatter/style_application/core/style_import.py`
  (`_ppr_children_by_name`, `_rpr_children_by_name`,
  `_effective_ppr_inner_in_arch`, `_effective_full_rpr_inner_in_arch`,
  `_materialize_full_rpr_for_detached_body`, `import_arch_styles_into_target`),
  `spec_formatter/style_application/core/xml_helpers.py`
  (`iter_direct_child_xml_blocks`),
  `spec_formatter/style_application/arch_env_applier.py` (`apply_doc_defaults`,
  `apply_environment_to_target`),
  `spec_formatter/style_application/core/errors.py` (`ERROR_REMEDIATIONS`),
  `docs/docx_method_hardening/probes/probe_style_import_w14.py`,
  `tests/test_unified_roundtrip.py` (`_deterministic_classifier`, the pattern
  for an end-to-end run without an API key), `CLAUDE.md`, `README.md`.
- Pitfalls already discovered that bear on this item: the portable
  stylesheet is the architect's `styles.xml` verbatim plus derived role
  styles, so anything in the architect's docDefaults reaches the importer;
  the Phase 1 side deliberately preserves extension properties
  (`tests/test_core_fix_regressions.py`), so carry them, do not drop them;
  `ET.tostring` currently normalizes fragment whitespace, so switching to raw
  fragments will change exact-string assertions in existing tests, which is
  expected.

## End of session

- One pull request for this session, containing only `WI-01` and its
  bookkeeping. Follow the **session end ritual** in the plan (section 3.3):
  run the checks, tick the boxes, set the tracker row to `in_review` with the
  PR URL, write `docs/docx_method_hardening/handoffs/handoff-for-session-02.md`
  from the template, paste it in chat, subscribe to the PR, drive CI to green,
  and paste the post-merge version of the handoff when the PR merges.
- `WI-01` is not the last required item, so no completion banner this session.
