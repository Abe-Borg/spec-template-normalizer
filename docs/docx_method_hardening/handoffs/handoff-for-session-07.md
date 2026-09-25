# Handoff prompt for session 07

Paste everything below the horizontal rule into the next session as its
first message. Written by session 06 on 2026-09-25 UTC.

---

You are continuing the **DOCX Method Hardening** program in the repository
`Abe-Borg/spec-template-normalizer`. Read these files completely before your
first tool call, in this order:

1. `CLAUDE.md` (engineering guide; every rule in it applies)
2. `docs/docx_method_hardening/DOCX_METHOD_HARDENING_PLAN.md`
3. `docs/docx_method_hardening/PROGRESS_TRACKER.md`
4. This prompt

## State at handoff

- Previous session: `06`, work item `WI-06: Revision accounting,
  discarded-revision warnings, collision-proof revision ids`.
- Pull request: `https://github.com/Abe-Borg/spec-template-normalizer/pull/66`.
  Merge status when this prompt was written: `open, CI pending`.
- Tracker rows changed by the previous session: `WI-05` -> `merged` with
  merge commit `e9fc64b`; `WI-06` -> `in_review` with the PR URL; session
  log row `06` added.
- Verification the previous session ran, with results:
  - `pytest` was not installed in the container. Run
    `pip install -r requirements-dev.txt` before the baseline.
  - Baseline before any change: `python -m pytest -q` gave 1617 passed,
    1 skipped. The skip is the GUI test, which needs `tkinter`.
  - The WI-06 tests were committed first (`a2122ba`), against the unchanged
    engine:
    - `tests/test_revision_accounting.py` failed to import (no
      `core/revisions.py`).
    - With its direct-call names stubbed, all 18 tests failed:
      - 5 damaged outputs published as success: a reviewer's `w:del` turned
        into a `w:ins`; a reviewer's `w:pPrChange` dropped; a reviewer's
        revision re-signed with this application's name; a tracked marker's
        `w:pPrChange` dropped; an unpredicted `w:pPrChange` in this
        application's name.
      - 6 runs recorded no revision census.
      - 2 runs recorded no header/footer revision counts.
      - 2 tracked conversions allocated ids that collide with the source's
        (900000 and 900003).
      - 3 direct tests need the new API.
    - `tests/test_conversion_verification.py` gained three id checks; two
      failed on the old engine.
  - After the implementation:
    - `tests/test_revision_accounting.py`: 18 passed;
      `tests/test_conversion_verification.py`: 15 passed;
    - `python -m pytest -q`: 1638 passed, 1 skipped;
    - `python -m pytest tests/test_sanitized_format_only_corpus.py -q`:
      2 passed;
    - `tests/test_engine_identity.py`: passes (no fingerprinted file
      touched); the tracker test: 6 passed;
    - both probes exit 0;
    - on Python 3.10 the application imports and the changed test modules
      pass (279 passed); pyflakes reports nothing new on any changed file.
    - A 3.10 interpreter is at `/usr/bin/python3.10`. `uv venv -p
      /usr/bin/python3.10` plus `uv pip install -r requirements-dev.txt`
      gives a working environment in seconds.
- Anything left unfinished, unexpected, or decided along the way:
  - **Decided.**
    - **One module, `core/revisions.py`.**
      - `revision_census` counts revision elements by `(author, kind)` from
        a *parsed* part, so it is namespace-aware and reads UTF-16.
      - `REVISION_KINDS` is the complete ECMA-376 17.13.5 family, not only
        the plan's eleven kinds: it adds move and custom-XML range markers,
        cell revisions, `w:tblPrExChange` and `w:tblGridChange`.
      - `max_annotation_id` / `highest_annotation_id_in_package` define the
        annotation-id space.
      - `revision_kinds_by_author` reads a paragraph fragment lexically,
        through the new public `xml_helpers.iter_start_tags`.
    - **The expected delta is the prediction.**
      `ExpectedParagraphChange.own_revisions` lists the kinds an edit adds
      in this application's name: `("ins", "pPrChange")` for a tracked
      automatic source, empty for every untracked edit.
      `ExpectedParagraphChanges.own_revisions_added()` sums them.
    - **Census at the gate.** It is recorded as the gate starts
      (`revisions_before`, `revisions_after`,
      `revisions_added_by_application`), beside the package census, and
      enforced after the body, section and run-property checks, before the
      package remit.
      - Other authors must keep each kind they had, per author.
      - This application's name must carry exactly its source revisions
        plus the prediction.
      - The failure message names kinds and counts, never an author.
    - **Per-paragraph check.** `prediction_mismatch` has a last check,
      `own_revision_kinds`, which the converter and the gate both run, so a
      misplaced `w:pPrChange` fails as `conversion_prediction_mismatch` with
      its location. `PREDICTION_CHECKS` gained it.
    - **`without_own_revisions` is unchanged.** Its consumers compare runs,
      and deleting a `w:pPrChange` is not the rejection it stands for.
    - **Header/footer revisions.**
      - The importer counts revisions per part before deleting a target part
        and before writing an architect part
        (`HeaderFooterImportResult.revisions_discarded` /
        `revisions_imported`).
      - It logs `WARNING: Discarded tracked revisions in replaced target
        part <name>: <n>` and `WARNING: Imported tracked revisions in
        architect part <name>: <n>`. Those two exact prefixes were added to
        `pipeline._SAFE_OPERATIONAL_PREFIXES`.
      - `arch_env_applier` copies the totals into `hf_result`.
      - The `apply_environment` event records
        `header_footer_revisions_discarded` / `_imported` (zero included).
        A warning-level `header_footer_revisions` event follows when either
        is non-zero.
      - "The audit" is `audit.json`'s `diagnostics` array, which holds every
        engine event whatever the diagnostics level. No new audit key, so no
        schema version bump.
      - A malformed target header/footer part now fails rather than being
        discarded uncounted.
    - **Ids.** `apply_canadian_to_csi` scans every `.xml` part below `word/`
      of a *tracked* target. `plan_canadian_to_csi` gained
      `highest_annotation_id` and allocates from max + 1 in document order:
      a paragraph's `w:pPrChange` first, then its `w:ins`. An id above
      2**31 - 1 fails closed (plain `ValueError`).
    - **Existing tests adjusted, stricter rather than looser:**
      - the four exact `verification_out` tests expect the three revision
        counters (0);
      - `test_markers_are_tracked_when_the_source_tracks_revisions` used to
        match only ids starting with `9`, which would have passed vacuously.
        It now matches this application's revisions by author and requires
        four distinct ids.
    - **No contract change.** No new error code, stage or policy field.
      `APPLICATION_POLICY_VERSION` is unchanged.
  - **Found, not fixed.** Raise these with the user; do not fold them into
    WI-07 unasked.
    1. **Format-only drops a `w:numberingChange`** (new, found by the
       census). When Format-only rewrites a paragraph's direct `w:numPr`, a
       `w:numberingChange` inside it is lost. On `master` that target
       published as a success with the revision silently gone (verified).
       The census now withholds it as `untrusted_error`. The fix belongs in
       the numbering materialization in `core/classification.py`. The
       revision is rare in practice (legacy Word writes it).
    2. **Non-standard settings or font table part names** (handoff 05
       finding 1, still open):
       - `apply_settings` edits `word/settings.xml` by name;
       - `apply_font_table` merges into `word/fontTable.xml` by name.

       The WI-05 whitelist admits these names; narrowing it belongs with the
       fix.
    3. **Settings order** (handoff 04 finding 5): `apply_settings` still
       inserts a missing `w:compat` just before `</w:settings>`.
    4. **Parity log lines.** WI-04's `Even/odd headers: ...` lines have no
       prefix in `_SAFE_OPERATIONAL_PREFIXES`, so `run.log` shows them only
       as `[untrusted detail omitted; sha256=...]`.
    5. **Handoff 04 findings 1 to 4 and 6** still stand:
       - tracked `w:rPrChange` on a heading run;
       - the doubled gap after `PART n`;
       - `html.unescape` in the forward edit;
       - `_TRACK_REVISIONS_RX` reads only double quotes;
       - session 02's and session 01's findings.

## Your assignment: `WI-07: Marker placement before leading structural run content`

- Follow the **session start ritual** in the plan (section 3.2) first:
  confirm the PR above is merged; if it is not, stop and tell the user rather
  than building on a stale base. Record the merge in the tracker
  (`WI-06` -> `merged`, with the merge commit), then mark `WI-07`
  `in_progress` with session `07`.
  - The session log row you add makes the tracker test require
    `handoffs/handoff-for-session-08.md`.
  - Commit a short placeholder with the right title line in your first
    commit, as sessions 01-06 did, and replace it before review.
- The plan section for `WI-07` is the specification. Its
  **Definition of done** checklist is what you must complete and tick. The
  plan rates it effort S.
- Key files:
  - `spec_formatter/style_application/core/canadian_to_csi.py`:
    - `_insert_marker` and `_FIRST_TEXT_RX`: the untracked marker is spliced
      before the first `w:t`. The tracked `w:ins` is placed before the run
      that holds it, found with `iter_element_xml_blocks(..., "w:r")`.
    - `_predict_marker_insertion`: the prediction of that placement on the
      run-content signature. It must move with the edit, or the converter's
      own prediction check refuses every paragraph it changes (session 03's
      note in the plan).
    - `_verify_marked_paragraph` and `_verify_prediction`: the plan asks
      for a structural post-check in both.
  - `tests/test_canadian_to_csi.py`: `_convert_tracked`, the geometry
    fixtures and `_TRACKING_ON`.
  - `tests/test_conversion_verification.py`: the independent verifier. It
    must share no code with the engine.
    - Its `_run_tabs` counts only `w:tab` whose parent is a direct `w:r` of
      the paragraph. A tracked marker's run sits inside `w:ins`, so it is
      not counted there.
- Pitfalls already discovered that bear on this item:
  - **WI-06's prediction now covers revisions.** A tracked paragraph's
    `ExpectedParagraphChange.own_revisions` is `("ins", "pPrChange")`, and
    `own_revision_kinds` checks it per paragraph. Moving the tracked `w:ins`
    earlier changes neither kind nor count, but a move that splits or
    duplicates the insertion will trip it.
    - Revision ids are allocated in document order: `w:pPrChange` (in
      `w:pPr`) and then `w:ins`. Placing the `w:ins` earlier in the
      paragraph keeps that order.
  - **The `_TRACKED_OR_FIELD_RX` refusal** is evaluated on the XML *before*
    the insertion point (`unprotected[: match.start()]`). If the insertion
    point moves earlier, re-check what that slice now covers. A field or
    tracked change that precedes a leading tab must still refuse the
    paragraph.
  - **`w:lastRenderedPageBreak` is inert**, and the run-content signature
    excludes it. The plan says the marker goes after `w:rPr` and after an
    inert `w:lastRenderedPageBreak`.
  - **Exact `verification_out` assertions** in
    `tests/style_application_regression/test_final_package_validation.py`
    now include the package and revision counters. A new counter recorded
    as a check starts must be added there with exact values.
  - **Mutating the package in tests.** Use `_damage_before_publication` in
    `tests/test_final_gate_text_identity.py`, and build any source that
    needs its own forward run *before* installing the damage. Otherwise the
    damage hits the forward run too (session 06 hit this).
  - **Encodings.** Fixture parts are often UTF-16. Compare in parsed XML or
    bytes, never in text decoded with an assumed encoding.
  - **Review bot.** The automated Codex reviewer on this repository posts
    P1/P2 findings as review threads. Earlier sessions treated them as bug
    reports to verify and fix, not as optional.

## End of session

- One pull request for this session, containing only `WI-07` and its
  bookkeeping. Follow the **session end ritual** in the plan (section 3.3):
  run the checks, tick the boxes, set the tracker row to `in_review` with the
  PR URL, write `docs/docx_method_hardening/handoffs/handoff-for-session-08.md`
  from the template, paste it in chat, subscribe to the PR, drive CI to green,
  and paste the post-merge version of the handoff when the PR merges.
- `WI-07` is not the last required item (WI-08 follows), so no completion
  banner this session.
