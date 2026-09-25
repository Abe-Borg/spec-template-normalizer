# Handoff prompt for session 05

Paste everything below the horizontal rule into the next session as its
first message. Written by session 04 on 2026-09-25 UTC.

---

You are continuing the **DOCX Method Hardening** program in the repository
`Abe-Borg/spec-template-normalizer`. Read these files completely before your
first tool call, in this order:

1. `CLAUDE.md` (engineering guide; every rule in it applies)
2. `docs/docx_method_hardening/DOCX_METHOD_HARDENING_PLAN.md`
3. `docs/docx_method_hardening/PROGRESS_TRACKER.md`
4. This prompt

## State at handoff

- Previous session: `04`, work item `WI-04: Even-page header parity follows
  the architect`.
- Pull request: `https://github.com/Abe-Borg/spec-template-normalizer/pull/64`.
  Merge status when this prompt was written: `open, CI pending`.
- Tracker rows changed by the previous session: `WI-03` -> `merged` with
  merge commit `23c6a54`; `WI-04` -> `in_review` with the PR URL; session
  log row `04` added.
- Verification the previous session ran, with results:
  - `pytest` was not installed in the container. Run
    `pip install -r requirements-dev.txt` before the baseline.
  - Baseline before any change: `python -m pytest -q` gave 1506 passed,
    1 skipped. The skip is the GUI test, which needs `tkinter`.
  - The WI-04 tests were committed first (37deb55), against the unchanged
    engine:
    - `tests/test_header_parity.py`: `ImportError` (no `apply_header_parity`);
    - the extended round trips, `test_unified_formatter_round_trips_without_api`
      and `test_unified_canadian_mode_converts_typed_csi_markers_end_to_end`:
      `assert 0 == 1`, because the output had no `w:evenAndOddHeaders`.
  - Before and after, real runs in both architect modes:
    - architect has an even header with the switch on: the output switch was
      off, and is now on;
    - architect has no even header, or only a dormant one, while the target
      switch is on: the output switch stayed on (even pages blank), and is
      now off;
    - architect has no headers: the target keeps its switch, before and after.
  - After the implementation:
    - `tests/test_header_parity.py`: 62 passed;
    - `python -m pytest -q`: 1568 passed, 1 skipped;
    - `python -m pytest tests/test_sanitized_format_only_corpus.py -q`:
      2 passed;
    - `tests/test_engine_identity.py`: passes (no fingerprinted file
      touched); the tracker test: 6 passed;
    - both probes exit 0;
    - changed modules import and run on Python 3.10; pyflakes is clean on
      every changed file.
- Anything left unfinished, unexpected, or decided along the way:
  - **Decided.**
    - New module `core/header_parity.py`:
      - `CT_SETTINGS_CHILD_ORDER` is the complete 98-element sequence;
      - `even_and_odd_headers` is the strict, namespace-aware reader. It
        raises `HeaderParityError` on a repeated element or a `w:val`
        outside the six `ST_OnOff` spellings;
      - `even_and_odd_headers_as_written` returns the switch uninterpreted;
      - `architect_even_and_odd_headers` reads the switch from
        `settings.settings_xml`. `None` means off; a missing field means
        unknown and raises;
      - `set_even_and_odd_headers` is the lexical writer. It proves each edit
        afterwards.
    - **The switch follows the header set, not the shell alone.** This
      refines the plan, and the plan now records it.
      - `HeaderFooterImportResult.replaced_target_parts` is set by
        `import_headers_footers`;
      - `apply_header_parity` (`arch_env_applier.py`) runs after the import,
        only when that flag is true;
      - a target that keeps its own headers keeps its own switch;
      - `_ensure_target_settings_part` is now the one creation path for both
        compat and parity.
    - **Preflight.** `_validate_header_parity` in `core/registry.py`, under
      `applies_shell`, rejects a template that has headers or footers but a
      switch that cannot be read. A dormant `even` reference is not rejected.
      Two hand-built test registries gained `settings.settings_xml`.
    - **Gate.** `_verify_header_parity` and `_document_settings_part` in
      `phase2_invariants.py`. The settings part is resolved through the
      document relationship. The check records `header_parity_checked`,
      `header_parity_follows_architect` and `even_and_odd_headers` as it
      starts. `apply_environment` records `header_parity_follows_architect`,
      `even_and_odd_headers` and `header_parity_changed`.
    - **Insertion point.** The switch goes after the last child that must
      precede it. A part with no valid position fails closed.
    - **Schema check.** The order table was checked against two copies of
      the transitional `wml.xsd`, which agree:
      - ISO/IEC 29500-4:2012, page 922, schema line 2896;
      - python-docx `ref/xsd/wml.xsd` at commit `e454546`.

      The proxy blocks ecma-international.org and raw.githubusercontent.com,
      but `git clone` of a public GitHub repository works; a PDF parser can
      be loaded from a downloaded wheel on `PYTHONPATH`. A test holds the
      constant equal to the plan's appendix B.
    - No new error code, stage or policy field. `APPLICATION_POLICY_VERSION`
      is unchanged: no policy field was added, as with WI-01 to WI-03.
  - **Found, not fixed.** Raise these with the user; do not fold them into
    WI-05 unasked.
    1. **Non-standard settings part name.** `apply_settings` and
       `apply_header_parity` work on `word/settings.xml` by name. If a
       target relates its settings under another name and has no
       `word/settings.xml`, creating one retargets the relationship
       (`_ensure_relationship_in_document_rels`). That orphans the target's
       real settings, silently dropping its tracking and protection state.
       - This predates WI-04 on the compat path; WI-04 made it reachable when
         the architect's switch is on.
       - The gate would still catch a parity mismatch, but not the orphaned
         settings.
       - Word always writes `word/settings.xml`, so this is rare.
    2. **Settings order (handoff 04 finding 5), still open.** `apply_settings`
       inserts a missing `w:compat` just before `</w:settings>`.
       `CT_SETTINGS_CHILD_ORDER` now exists, so the fix is small, but it was
       not in WI-04's scope.
    3. **Handoff 04 findings 1 to 4 and 6** still stand:
       - tracked `w:rPrChange` on a heading run;
       - the doubled gap after `PART n`;
       - `html.unescape` in the forward edit;
       - `_TRACK_REVISIONS_RX` reads only double quotes;
       - session 02's and session 01's findings.

       The WI-04 reader does not use `_TRACK_REVISIONS_RX`; it parses, so it
       accepts either quoting.

## Your assignment: `WI-05: Package-level change whitelist invariant`

- Follow the **session start ritual** in the plan (section 3.2) first:
  confirm the PR above is merged; if it is not, stop and tell the user rather
  than building on a stale base. Record the merge in the tracker
  (`WI-04` -> `merged`, with the merge commit), then mark `WI-05`
  `in_progress` with session `05`.
  - The session log row you add makes the tracker test require
    `handoffs/handoff-for-session-06.md`.
  - Commit a short placeholder with the right title line in your first
    commit, as sessions 01-04 did, and replace it before review.
- The plan section for `WI-05` is the specification. Its
  **Definition of done** checklist is what you must complete and tick.
- Key files:
  - `spec_formatter/style_application/phase2_invariants.py`:
    `verify_phase2_invariants`, where the member census belongs. The source
    package is already opened there.
  - `spec_formatter/style_application/batch_runner.py`:
    - `_build_and_patch_output` passes `OPTIONAL_REPLACEMENT_PARTS` as
      replacements in every mode;
    - `env_result["header_footer_import"]` is the header/footer manifest to
      thread into the gate.
  - `env_result` also carries `header_parity` and `replaced_target_parts`.
  - `spec_formatter/style_application/docx_patch.py`: `patch_docx`.
  - `spec_formatter/style_application/core/application_policy.py`: derive
    the allowed set from policy fields, never from a mode name.
- Pitfalls already discovered that bear on this item:
  - **Parts can be added, not only changed.** Under the full shell,
    `word/settings.xml` may be **added**: `_ensure_target_settings_part`
    creates it when the target has none and the architect has compat or its
    even/odd switch is on. `[Content_Types].xml` and
    `word/_rels/document.xml.rels` change with it. Theme and fontTable can
    be added the same way.
  - **Headers can be replaced or kept.** Header/footer parts are replaced
    only when `replaced_target_parts` is true; otherwise they must stay
    byte-identical, which `_verify_target_header_footer_preserved` already
    checks.
  - **Architect-free modes.** `test_architect_free_modes_leave_the_target_settings_byte_identical`
    in `tests/test_header_parity.py` already proves the settings part
    unchanged in both architect-free modes; WI-05 generalizes that to every
    part.
  - **Encodings.** Fixture parts are often UTF-16. Compare member bytes and
    hashes, never decoded text.
  - **Mutating the package in tests.** To change the output after the
    engine's edits, wrap `batch_runner._build_and_patch_output` and edit
    files in `extract_dir` with `read_bytes` / `write_bytes`, as
    `tests/test_final_gate_text_identity.py` and `tests/test_header_parity.py`
    do. A member that is not in the replacement set is copied from the
    source by `patch_docx`, so an out-of-remit change must be injected into
    the packaged output (or into a replacement) to reach the gate.
  - **Review bot.** The automated Codex reviewer on this repository posts
    P1/P2 findings as review threads. Earlier sessions treated them as bug
    reports to verify and fix, not as optional.

## End of session

- One pull request for this session, containing only `WI-05` and its
  bookkeeping. Follow the **session end ritual** in the plan (section 3.3):
  run the checks, tick the boxes, set the tracker row to `in_review` with the
  PR URL, write `docs/docx_method_hardening/handoffs/handoff-for-session-06.md`
  from the template, paste it in chat, subscribe to the PR, drive CI to green,
  and paste the post-merge version of the handoff when the PR merges.
- `WI-05` is not the last required item, so no completion banner this session.
