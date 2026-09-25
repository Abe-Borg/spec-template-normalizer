# Handoff prompt for session 04

Paste everything below the horizontal rule into the next session as its
first message. Written by session 03 on 2026-09-25 UTC.

---

You are continuing the **DOCX Method Hardening** program in the repository
`Abe-Borg/spec-template-normalizer`. Read these files completely before your
first tool call, in this order:

1. `CLAUDE.md` (engineering guide; every rule in it applies)
2. `docs/docx_method_hardening/DOCX_METHOD_HARDENING_PLAN.md`
3. `docs/docx_method_hardening/PROGRESS_TRACKER.md`
4. This prompt

## State at handoff

- Previous session: `03`, work item `WI-03: Final-gate text identity for
  every mode (identity except the enumerated diff)`.
- Pull request: `https://github.com/Abe-Borg/spec-template-normalizer/pull/63`.
  Merge status when this prompt was written: `open, CI running`.
- Tracker rows changed by the previous session: `WI-02` -> `merged` with
  merge commit `9f8d5b0`; `WI-03` -> `in_review` with the PR URL; session
  log row `03` added.
- Verification the previous session ran, with results:
  - `pytest` was not installed in the container. Run
    `pip install -r requirements-dev.txt` before the baseline.
  - Baseline before any change: `python -m pytest -q` gave 1463 passed,
    1 skipped. The skip is the GUI test, which needs `tkinter`.
  - Owner-approved placement fix, its own commit (e2725fa):
    - the new regression test failed before the fix (a `w:ins` found inside
      a `w:r`) and passes after it;
    - a real tracked `canadian_to_csi` run over bold runs published 10 of 10
      markers inside their runs before the fix, and 0 after it.
  - The WI-03 tests were committed first (de967b9), against the unchanged
    engine:
    - `tests/test_final_gate_text_identity.py`: 16 failed. Nine damaged runs
      published (`assert not True`), one of them a visible word change. Two
      were caught only by the run-property invariant, with no code. Five
      happy paths lacked the counters (`KeyError`).
    - `tests/test_conversion_prediction.py` and the new section of
      `test_final_package_validation.py`: `ImportError`.
    - The new `test_batch_runner.py` plumbing test: `ModuleNotFoundError`.
  - After the implementation and one review round (see below):
    - new and affected tests: 150 passed before the review round;
    - `python -m pytest -q`: 1506 passed, 1 skipped;
    - `python -m pytest tests/test_sanitized_format_only_corpus.py -q`:
      2 passed;
    - `tests/test_engine_identity.py`: 4 passed (no fingerprinted file
      touched); the tracker test: 6 passed;
    - both probes exit 0;
    - changed modules import and run on Python 3.10; pyflakes is clean apart
      from two unused imports in `core/canadian_to_csi.py` that predate the
      PR.
- Anything left unfinished, unexpected, or decided along the way:
  - **Decided.**
    - New module `core/expected_changes.py`:
      - `ExpectedParagraphChange` holds `visible_text`, `run_content` and
        `run_content_outside_own_revisions`;
      - `ExpectedParagraphChanges` holds `changes` and `roles`, both
        read-only; `NO_EXPECTED_PARAGRAPH_CHANGES` is the empty value;
      - `prediction_mismatch` / `first_prediction_mismatch` are the one
        comparison;
      - `MARKER_REVISION_AUTHOR` and `without_own_revisions` moved here
        (re-exported from `canadian_to_csi`; `phase2_invariants` imports the
        projection as `_without_own_revisions`).
    - Predictions are worked out from the *source* paragraph's run-content
      signature, never read off the edited XML:
      - `_predict_marker_removal` in `marker_tools`;
      - `_predict_marker_insertion` and `_predict_marked_paragraph` in
        `canadian_to_csi`;
      - each converter asserts its assembled document against them exactly.
    - `apply_csi_to_canadian` / `apply_canadian_to_csi` return
      `ConversionResult(report, expected_paragraph_changes)`.
      `ConversionPlan` carries the prediction too.
    - The gate (`_verify_conversion_body_invariants`) runs for the three
      conversion modes:
      - unpredicted paragraphs are compared exactly, *without* projection
        (stricter than the plan's text, which would hide an unpredicted
        marker);
      - predicted paragraphs are checked for visible text, then exact run
        content, then run content outside own revisions;
      - failures are `conversion_prediction_mismatch` with a
        SECTION/heading location;
      - an omitted prediction allows no change; Format-only refuses a
        non-empty one.
    - Counters: `body_signature_paragraphs_compared` and
      `body_paragraphs_expected_changed`. `record_body_check` writes them as
      the body check starts, and the comparison writes them again in a
      `finally`.
  - **Review (Codex, two P2s, both reproduced, fixed in 3a6de43, threads
    resolved).**
    - An unpredicted paragraph is now compared by visible text too. The
      signature records runs, not their containers, so a run wrapped in
      `w:del` / `w:moveFrom` kept its signature while its text vanished.
    - The counters are recorded from the start of the body check, in both
      branches, so a paragraph-count or Format-only text failure still shows
      the check ran.
  - **Owner decision.** The user was asked in chat and chose to fix the
    tracked-marker placement bug in WI-03's PR, in its own commit.
    `_insert_marker` searched back for `<w:r`, which matched `<w:rPr` and put
    the `w:ins` inside every formatted run (invalid OOXML that published).
  - **Plan text.**
    - WI-03 gained an *as implemented* note with the decisions above.
    - WI-07's claim that the tracked path was already correct is corrected.
      WI-07 must also move `_predict_marker_insertion` together with the
      edit, or the converter's own prediction check refuses every paragraph
      it changes.
    - WI-06 names the projection's new home.
  - **Found, not fixed.** Raise these with the user; do not fold them into
    WI-04 unasked.
    1. **Tracked formatting change on a heading run.** Tracked
       `canadian_to_csi` on a run whose `w:rPr` holds a `w:rPrChange` fails
       with an uncoded `ValueError`: `_run_properties` copies the
       protected-subtree placeholder into the marker run, and the restore
       step finds it twice.
    2. **Doubled gap after `PART n`.** The forward conversion keeps the
       delimiter tab when the `- ` after `PART n` continues in the next text
       node, so the output has a doubled gap after the automatic number. The
       round trip returns `PART 1<tab><tab>GENERAL`
       (`test_forward_prediction_keeps_a_tab_the_separator_runs_past` pins
       today's behaviour).
    3. **HTML decoding in the forward edit.** The forward edit decodes `w:t`
       with `html.unescape` and re-encodes with `html.escape`. C1 character
       references (`&#x80;` becomes a euro sign) and `&#13;` are therefore
       rewritten. WI-03's prediction now makes such a paragraph fail closed;
       the root fix, decoding as XML, is not done.
    4. **Quoting in `_TRACK_REVISIONS_RX`.** It reads only a double-quoted
       `w:val`, so `<w:trackRevisions w:val='false'/>` reads as tracking on.
       WI-04's plan says to mirror it for `w:evenAndOddHeaders`: accept either
       quoting there.
    5. **Settings order.** When the target has no `w:compat`, `apply_settings`
       inserts one just before `</w:settings>`. That breaks `CT_Settings`
       order whenever later children exist (`w:rsids`, `w:themeFontLang`,
       `w:clrSchemeMapping`, `w:shapeDefaults`, `w:decimalSymbol`,
       `w:listSeparator`, ...). This is the table WI-04 builds.
    6. **Earlier findings.** Session 02's findings still stand (handoff 03
       has the detail):
       - Format-only body failures carry no error code;
       - containers are not recorded by the signature;
       - math runs are unchecked;
       - scanning cost;
       - session 01's four findings.

       Session 02's Windows note -- `header_footer_importer` writes
       `word/document.xml` with `Path.write_text` -- now also covers the
       architect `csi_to_canadian` body, which WI-03's gate compares. A
       literal CR in a text node would fail that target on Windows only;
       Word never writes one.

## Your assignment: `WI-04: Even-page header parity follows the architect`

- Follow the **session start ritual** in the plan (section 3.2) first:
  confirm the PR above is merged; if it is not, stop and tell the user rather
  than building on a stale base. Record the merge in the tracker
  (`WI-03` -> `merged`, with the merge commit), then mark `WI-04`
  `in_progress` with session `04`.
  - The session log row you add makes the tracker test require
    `handoffs/handoff-for-session-05.md`.
  - Commit a short placeholder with the right title line in your first
    commit, as sessions 01-03 did, and replace it before review.
- The plan section for `WI-04` is the specification. Its
  **Definition of done** checklist is what you must complete and tick.
- Key files:
  - `spec_formatter/style_application/arch_env_applier.py`:
    - `apply_settings` applies only the architect's `w:compat` block, and
      returns early when the registry has none;
    - `_MINIMAL_SETTINGS_XML`, `_ensure_settings_in_content_types` and
      `_ensure_settings_in_rels` create the part when the target lacks one;
    - `apply_environment_to_target` calls it.
  - Root `arch_env_extractor.py`: `extract_settings` captures
    `settings.settings_xml`, canonicalized. This file is **fingerprinted**.
    Deriving parity from the registry, as the plan says, avoids touching it.
    If you do touch it, run `python engine_identity.py` and commit the new
    digest.
  - `spec_formatter/style_application/header_footer_importer.py`:
    `import_headers_footers` and `_rewire_document_sectpr`, which wire the
    `default`, `first` and `even` references.
  - `spec_formatter/style_application/phase2_invariants.py`:
    - `verify_phase2_invariants` step 2, the header/footer section, is where
      `header_parity_checked` belongs;
    - the WI-03 body check runs first and is unaffected by `settings.xml`.
  - `spec_formatter/style_application/core/canadian_to_csi.py`:
    `_TRACK_REVISIONS_RX` / `source_tracks_revisions`, the "off" spellings
    to mirror (see finding 4).
  - Tests:
    - `tests/test_unified_roundtrip.py`: `_write_docx(architect=True)` writes
      settings with `<w:evenAndOddHeaders/>` as UTF-16 LE, and the target's
      settings are UTF-16 BE; `test_unified_formatter_round_trips_without_api`
      is the even-header run to extend;
    - `tests/test_architect_free_modes.py`;
    - the settings tests under `tests/style_application_regression/`.
- Pitfalls already discovered that bear on this item:
  - **Parity must not depend on the compat block.** `apply_settings` returns
    early when the architect has no `w:compat`, and only that path creates a
    missing settings part.
  - **The order table must be verified against the schema text.** The
    repository holds no `wml.xsd`. The container has HTTPS through a proxy,
    so fetch the ECMA-376 or ISO/IEC 29500 schema, check appendix B against
    it, and cite the source in the PR.
  - **Encodings.** The fixtures' settings parts are UTF-16, so read and
    write them with `read_xml_text` / `write_xml_text` (`core/ooxml_text.py`).
    Never use `Path.read_text` or `Path.write_text`, which also translates
    newlines on Windows.
  - **Architect-free modes.** They must leave the target's settings
    byte-identical. WI-05 will later prove that for every part.
  - **Mutating the package in tests.** To change the output after the
    engine's edits, wrap `batch_runner._build_and_patch_output` and edit
    files in `extract_dir` with `read_bytes` / `write_bytes`, as
    `tests/test_final_gate_text_identity.py` does.
  - **Review bot.** The automated Codex reviewer on this repository posts
    P1/P2 findings as review threads. Earlier sessions treated them as bug
    reports to verify and fix, not as optional.

## End of session

- One pull request for this session, containing only `WI-04` and its
  bookkeeping. Follow the **session end ritual** in the plan (section 3.3):
  run the checks, tick the boxes, set the tracker row to `in_review` with the
  PR URL, write `docs/docx_method_hardening/handoffs/handoff-for-session-05.md`
  from the template, paste it in chat, subscribe to the PR, drive CI to green,
  and paste the post-merge version of the handoff when the PR merges.
- `WI-04` is not the last required item, so no completion banner this session.
