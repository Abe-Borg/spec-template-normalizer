# Handoff prompt for session 06

Paste everything below the horizontal rule into the next session as its
first message. Written by session 05 on 2026-09-25 UTC.

---

You are continuing the **DOCX Method Hardening** program in the repository
`Abe-Borg/spec-template-normalizer`. Read these files completely before your
first tool call, in this order:

1. `CLAUDE.md` (engineering guide; every rule in it applies)
2. `docs/docx_method_hardening/DOCX_METHOD_HARDENING_PLAN.md`
3. `docs/docx_method_hardening/PROGRESS_TRACKER.md`
4. This prompt

## State at handoff

- Previous session: `05`, work item `WI-05: Package-level change whitelist
  invariant`.
- Pull request: `https://github.com/Abe-Borg/spec-template-normalizer/pull/65`.
  Merge status when this prompt was written: `open, CI pending`.
- Tracker rows changed by the previous session: `WI-04` -> `merged` with
  merge commit `f3aef0b`; `WI-05` -> `in_review` with the PR URL; session
  log row `05` added.
- Verification the previous session ran, with results:
  - `pytest` was not installed in the container. Run
    `pip install -r requirements-dev.txt` before the baseline.
  - Baseline before any change: `python -m pytest -q` gave 1570 passed,
    1 skipped. The skip is the GUI test, which needs `tkinter`.
  - The WI-05 tests were committed first (`0323b6d`), against the unchanged
    engine:
    - `tests/test_package_change_whitelist.py` failed to import (no
      `_package_member_census`).
    - With the two direct-call names stubbed, 35 end-to-end tests failed:
      - 23 outputs carrying an out-of-remit change published as success. The
        changes were: footnotes, comments or a `[trash]` item altered in each
        mode's packaged output; a member removed and another added; stray
        styles, settings (UTF-16), numbering, theme or font table writes in
        `canadian_to_csi`; and a styles change or added theme in the
        standalone mode.
      - 4 happy paths had no `package_members_*` counters.
      - `run.log` named no member.
  - After the implementation:
    - `tests/test_package_change_whitelist.py`: 47 passed;
    - `python -m pytest -q`: 1617 passed, 1 skipped;
    - `python -m pytest tests/test_sanitized_format_only_corpus.py -q`:
      2 passed;
    - `tests/test_engine_identity.py`: passes (no fingerprinted file
      touched); the tracker test: 6 passed;
    - both probes exit 0;
    - on Python 3.10 the application imports and the changed test modules
      pass (183 passed); pyflakes is clean on every changed file.
    - A 3.10 interpreter is at `/usr/bin/python3.10`. `uv venv -p
      /usr/bin/python3.10` plus `uv pip install -r requirements-dev.txt`
      gives a working environment in seconds.
- Anything left unfinished, unexpected, or decided along the way:
  - **Decided.**
    - **Census and remit.** They live in `phase2_invariants.py`:
      - `_package_member_census` hashes every member's bytes (SHA-256, never
        decoded text);
      - `_package_remit` derives the allowed sets from `applies_role_styles`,
        `import_body_numbering` and `apply_full_architect_shell`, never from
        a mode name;
      - `_enforce_package_remit` fails with `INVARIANT FAIL: package member
        outside the mode's remit <changed|added|removed>: <name>`, reporting
        the first in name order plus a count of the rest;
      - `_verify_package_member_remit` composes the census and the
        enforcement for direct use.
    - **Recorded first, enforced last.** `verify_phase2_invariants` takes and
      records the census before the body check, so
      `package_members_compared/_changed/_added/_removed` reach the
      `build_output` event on every failure path. It enforces the remit after
      every other check, so the more specific messages win.
    - **New parameters.** `verify_phase2_invariants` gained
      `header_footer_manifest` and `log`. `_build_and_patch_output` gained
      `log` and passes the manifest that chose its replacements.
    - **An omitted manifest allows no header change.** A mode with no shell
      that is handed a manifest naming anything raises `ValueError`.
    - **Shell parts.** The settings, theme and font table parts the
      *output's* document relates, **plus** the conventional names
      (`word/settings.xml`, `word/theme/theme1.xml`, `word/fontTable.xml`).
      The handoff for this session said to treat the related part, not the
      conventional name, as the settings part. The related part is admitted,
      but the conventional name had to stay too:
      `test_switch_is_applied_to_the_settings_part_the_document_relates`
      (WI-04) is exactly finding 1b below. The compat step writes a stray,
      unrelated `word/settings.xml`, and refusing that would change what such
      a run publishes. That is the owner's call. The plan's WI-05 *as
      implemented* note records this.
    - **The header set is cross-checked against the packages as well as the
      manifest:**
      - a header or footer part or its `.rels` needs the manifest **and** an
        output document relationship;
      - media may only be added, and needs the manifest and an output header
        relationship;
      - a removal needs the manifest and must have been a header or footer
        part the source's document related (or that part's `.rels`).
    - **Names go to the run log only.** Lines read `Package member <verb>:
      <name>`, and `pipeline._SAFE_OPERATIONAL_PREFIXES` gained
      `"Package member"`.
    - **Existing tests adjusted, stricter rather than looser:**
      - four tests that assert an early failure's exact `verification_out`
        now also expect the census counters, with exact values;
      - the direct header/footer gate test in `test_sectpr_tools.py` passes
        the manifest, and asserts that omitting it fails.
    - **No contract change.** No new error code, stage or policy field.
      `APPLICATION_POLICY_VERSION` is unchanged.
  - **Found, not fixed.** Raise these with the user; do not fold them into
    WI-06 unasked.
    1. **Non-standard settings or font table part names** (handoff 05
       finding 1, still open, same family):
       - `apply_settings` edits `word/settings.xml` by name;
       - `apply_font_table` merges into `word/fontTable.xml` by name, and
         relates it only when it creates it.

       What goes wrong:
       - A target whose document relates its settings under another name,
         with no `word/settings.xml`, gets a new one; the relationship is
         retargeted to it and the real settings are orphaned.
       - With a stray `word/settings.xml` beside the related part, compat
         goes into the stray and never takes effect.

       The WI-05 whitelist deliberately admits these names. Narrowing it to
       related parts belongs with the fix.
    2. **Settings order** (handoff 04 finding 5): `apply_settings` still
       inserts a missing `w:compat` just before `</w:settings>`.
    3. **Parity log lines.** WI-04's `Even/odd headers: ...` log lines have
       no prefix in `_SAFE_OPERATIONAL_PREFIXES`, so `run.log` shows them
       only as `[untrusted detail omitted; sha256=...]`.
    4. **Handoff 04 findings 1 to 4 and 6** still stand:
       - tracked `w:rPrChange` on a heading run;
       - the doubled gap after `PART n`;
       - `html.unescape` in the forward edit;
       - `_TRACK_REVISIONS_RX` reads only double quotes;
       - session 02's and session 01's findings.

## Your assignment: `WI-06: Revision accounting, discarded-revision warnings, collision-proof revision ids`

- Follow the **session start ritual** in the plan (section 3.2) first:
  confirm the PR above is merged; if it is not, stop and tell the user rather
  than building on a stale base. Record the merge in the tracker
  (`WI-05` -> `merged`, with the merge commit), then mark `WI-06`
  `in_progress` with session `06`.
  - The session log row you add makes the tracker test require
    `handoffs/handoff-for-session-07.md`.
  - Commit a short placeholder with the right title line in your first
    commit, as sessions 01-05 did, and replace it before review.
- The plan section for `WI-06` is the specification. Its
  **Definition of done** checklist is what you must complete and tick. It
  has three parts: a census, warnings and ids. The plan rates it effort M;
  if it does not fit one session, use the partial-PR rule in section 3.6.
- Key files:
  - `spec_formatter/style_application/phase2_invariants.py`:
    `verify_phase2_invariants`, where the revision census belongs. Follow
    WI-05's pattern: record the counts as the check starts (on success and
    on every failure), and let the more specific checks report first.
  - `spec_formatter/style_application/core/expected_changes.py`:
    `without_own_revisions` and `MARKER_REVISION_AUTHOR`. It projects out
    only `w:ins`, not the `w:pPrChange` the tracked reverse conversion also
    writes.
  - `spec_formatter/style_application/core/canadian_to_csi.py`:
    `_MARKER_REVISION_ID_BASE = 900000` (line 141) and the two
    `revision_id=` allocations near line 1190.
  - `spec_formatter/style_application/header_footer_importer.py`:
    - `_remove_existing_hf_files` deletes the target's parts;
    - `_write_hf_parts` writes the architect's;
    - `HeaderFooterImportResult` is the manifest.
  - `spec_formatter/style_application/arch_env_applier.py`:
    `apply_environment_to_target` copies the manifest into `hf_result` **key
    by key**. New counts must be added there, or they never reach
    `env_result`.
  - `spec_formatter/style_application/batch_runner.py`: the
    `apply_environment` diagnostics event (`phase.set(...)`) and
    `_build_and_patch_output`.
  - `tests/test_conversion_verification.py`: the independent stdlib verifier
    the plan asks you to extend for the id assertions. It must share no code
    with the engine.
- Pitfalls already discovered that bear on this item:
  - **`WARNING:` lines do not reach `run.log` verbatim.**
    - The plan asks for a `WARNING:` line naming the part and count.
    - `pipeline._SAFE_OPERATIONAL_PREFIXES` has no `WARNING` entry, so such
      a line is written as `[untrusted detail omitted; sha256=...]`.
    - WI-05 added `"Package member"` the same way. Add a specific prefix
      (and test that the line survives into `run.log`) rather than a bare
      `"WARNING"`, which would let any warning text through.
  - **The header/footer manifest is now read by the final gate.**
    `_header_footer_manifest_names` reads only its five name keys, so extra
    count keys are harmless. Changing what `part_names`, `rels_names`,
    `media_names`, `removed_part_names` or `removed_rels_names` mean changes
    the whitelist too.
  - **Exact `verification_out` assertions.** Several tests in
    `tests/style_application_regression/test_final_package_validation.py`
    assert an early failure's *exact* `verification_out` dict (four were
    extended with the package counters in WI-05). New counters recorded as
    a check starts will appear there too: extend those expectations with
    exact values; do not loosen them.
  - **Paired annotation ids.** The plan is explicit that valid OOXML repeats
    an id across paired annotations (`w:bookmarkStart`/`w:bookmarkEnd`,
    comment ranges). Assert uniqueness among revision elements only.
  - **Mutating the package in tests.** Two helpers in
    `tests/test_package_change_whitelist.py` show how to stage damage:
    - `_alter_before_packaging` wraps `batch_runner._build_and_patch_output`
      and edits files in `extract_dir` with bytes;
    - `_alter_packaged_output` wraps `batch_runner.patch_docx` and rewrites
      the built output, for members that are not replacements.
  - **Encodings.** Fixture parts are often UTF-16. Count and compare in
    parsed XML or bytes, never in text decoded with an assumed encoding.
  - **Review bot.** The automated Codex reviewer on this repository posts
    P1/P2 findings as review threads. Earlier sessions treated them as bug
    reports to verify and fix, not as optional.

## End of session

- One pull request for this session, containing only `WI-06` and its
  bookkeeping. Follow the **session end ritual** in the plan (section 3.3):
  run the checks, tick the boxes, set the tracker row to `in_review` with the
  PR URL, write `docs/docx_method_hardening/handoffs/handoff-for-session-07.md`
  from the template, paste it in chat, subscribe to the PR, drive CI to green,
  and paste the post-merge version of the handoff when the PR merges.
- `WI-06` is not the last required item, so no completion banner this session.
