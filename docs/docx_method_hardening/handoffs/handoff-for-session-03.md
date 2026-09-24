# Handoff prompt for session 03

Paste everything below the horizontal rule into the next session as its
first message. Written by session 02 on 2026-09-24 UTC.

---

You are continuing the **DOCX Method Hardening** program in the repository
`Abe-Borg/spec-template-normalizer`. Read these files completely before your
first tool call, in this order:

1. `CLAUDE.md` (engineering guide; every rule in it applies)
2. `docs/docx_method_hardening/DOCX_METHOD_HARDENING_PLAN.md`
3. `docs/docx_method_hardening/PROGRESS_TRACKER.md`
4. This prompt

## State at handoff

- Previous session: `02`, work item `WI-02: Exact run-content signature at the
  Format-only gate`.
- Pull request: `https://github.com/Abe-Borg/spec-template-normalizer/pull/62`.
  Merge status when this prompt was written: `open, CI running`.
- Tracker rows changed by the previous session: `WI-01` -> `merged` with
  merge commit `4082087`; `WI-02` -> `in_review` with the PR URL; session
  log row `02` added.
- Verification the previous session ran, with results:
  - Baseline before any change: `python -m pytest -q` gave 1372 passed,
    1 skipped.
  - The WI-02 tests were committed first (5c53dc6), against the unchanged
    engine:
    - the seven probe rows, as gate tests: `DID NOT RAISE`;
    - the probe run as a test: exit 1, seven mutations accepted;
    - the counter tests: `KeyError`;
    - a real Format-only run with a tab dropped after classification
      application: it **succeeded**;
    - the signature unit module: `ImportError`, because the function did not
      exist yet.
  - After the fix: 137 new and affected tests passed.
  - `python -m pytest -q`: 1463 passed, 1 skipped. The skip is the GUI test,
    which needs `tkinter`.
  - `python -m pytest tests/test_sanitized_format_only_corpus.py -q`:
    2 passed.
  - `python -m pytest tests/test_engine_identity.py -q`: 4 passed. No
    fingerprinted file was touched.
  - `probe_format_only_gate.py`: every row REJECTED, exit 0.
    `probe_style_import_w14.py`: both rows OK, exit 0.
  - Changed modules import and run on Python 3.10; pyflakes is clean.
- Anything left unfinished, unexpected, or decided along the way:
  - **Decided.**
    - The signature lives in `core/xml_helpers.py`:
      `paragraph_run_content_signature(p_xml)` returns one tuple of items per
      run. `run_content_difference(before, after)` returns the kind of the
      first difference, or `None`. `RUN_CONTENT_DIFFERENCE_KINDS` is the
      closed set of kinds; `RunContentItem` and `RunContentSignature` are the
      types.
    - `("ref", kind, id)` uses the element's local name as `kind`.
    - Decoding is what an XML parser reports: line ends are normalized first
      (XML 1.0 section 2.11), then `xml_unescape` expands references. CDATA
      is literal, and comments and processing instructions are not content.
      Element markup inside a text node is kept verbatim.
    - Runs nested in a run's own `w:ruby` follow the run that holds them.
      Only `("other", ...)` children are searched for nested runs.
    - Kinds are reported over the flattened items:
      - `run_boundary` means the same items were regrouped into runs, or an
        empty run was added or removed;
      - `preserve_space` means only `xml:space` changed;
      - otherwise the removed or added item is named when the rest still
        lines up.
    - The gate runs the normalized-text check first (message unchanged), then
      `_verify_format_only_run_content`, then the numbering checks.
      `body_signature_paragraphs_compared` is written in a `finally`, so it
      also appears on the failure path.
    - The failure stays a `RuntimeError` with the plan's message; no error
      code. Session 02 checked a real failing run: `run.json`, `audit.json`
      and `run.log` show `untrusted_error`, the hashed detail and no
      location. The kind and index exist only in the developer detail, and
      the README makes no claim otherwise. The existing "body text changed"
      failure has always behaved the same way.
    - The probe now uses `TemporaryDirectory` and closes its `ZipFile`,
      because the suite runs it; Windows cannot delete an open file. Its
      note column is 100 characters wide, so no verdict is truncated.
  - **Plan text.** Appendix A gained an *as implemented* note recording the
    decisions above.
  - **Unexpected, not fixed (outside WI-02).**
    - `header_footer_importer.py` writes `word/document.xml` (in
      `_rewire_document_sectpr`) and header/footer parts with
      `Path.write_text`. On Windows that translates every `\n` to `\r\n`, so
      an existing `\r\n` becomes `\r\r\n`.
    - The signature normalizes line ends, so the ordinary `\n` -> `\r\n` case
      is invisible to it, as it is to any parser.
    - A literal CR inside a text node would now fail the target, on Windows
      only. That is correct: its parsed text really changes.
    - Word never writes literal CRs into text, and no fixture has one outside
      the signature's own unit test. The fix is `write_xml_text`.
  - **Adjacent findings, not scheduled** (raise them with the user; do not
    fold them into WI-03 unasked):
    1. **Format-only body failures carry no error code.** This covers text,
       run content, numbering, definitions and out-of-scope structure. Users
       see `untrusted_error` and cannot find the paragraph.
       - WI-03's plan keeps "the Format-only message for Format-only".
       - A code (for example `format_only_body_changed`) with
         `ErrorLocation(paragraph_index=N)` would fix it. It could be
         attached with `attach_engine_error`, which keeps the `RuntimeError`
         type, so `match=` tests hold.
       - That is a contract change: it needs `ERROR_REMEDIATIONS` (and the
         `15 <= len <= 22` bound in `test_engine_errors.py`), the CLAUDE.md
         code list, and the README.
    2. **Containers are not recorded.** Per appendix A, the signature records
       run content but not the container a run sits in, nor the container's
       attributes (`w:hyperlink` `r:id`/`w:anchor`, `w:fldSimple` `w:instr`,
       `w:sdt` properties).
       - Unwrapping a hyperlink or content control therefore passes.
       - WI-06's census covers revision wrappers; nothing covers the others.
    3. **Math is unchecked.** Math runs (`m:r`/`m:t`) are outside both the
       normalized-text check (it reads `w:t` only) and the signature (it
       walks `w:r`).
    4. **Scanning cost.** The cost of the signature, and of
       `iter_direct_child_xml_blocks` everywhere, is dominated by
       `_scan_markup`'s character-by-character quote scan. Measured cost for
       2,000 paragraphs of 6 runs: signature 0.53 s per side, run-property
       invariant 1.86 s. Not a problem today; noted in case a large target
       ever is.
    5. **Session 01's four findings still stand** (handoff 02 has the
       detail):
       - `numbering_importer.inject_numbering_into_xml` merges declarations
         by prefix only and never merges `mc:Ignorable`;
       - `_extract_tag_inner` truncates a style's `w:rPr`/`w:pPr` at a
         nested `w:rPrChange`/`w:pPrChange`;
       - materialized `w:` children keep first-seen order, not the schema
         sequence;
       - Phase 1's `_propagate_style_fragment_namespaces` drops
         `mc:Ignorable` tokens.

## Your assignment: `WI-03: Final-gate text identity for every mode (identity except the enumerated diff)`

- Follow the **session start ritual** in the plan (section 3.2) first:
  confirm the PR above is merged; if it is not, stop and tell the user rather
  than building on a stale base. Record the merge in the tracker
  (`WI-02` -> `merged`, with the merge commit), then mark `WI-03`
  `in_progress` with session `03`. The session log row you add makes the
  tracker test require `handoffs/handoff-for-session-04.md`; commit a short
  placeholder with the right title line in your first commit, as sessions
  01 and 02 did, and replace it before review.
- The plan section for `WI-03` is the specification. Its
  **Definition of done** checklist is what you must complete and tick.
- Key files:
  - `spec_formatter/style_application/phase2_invariants.py`:
    - `verify_phase2_invariants` runs the body checks only under
      `policy.preserve_target_numbering`; that is what WI-03 widens.
    - `_verify_format_only_body_invariants` and
      `_verify_format_only_run_content` hold the WI-02 check and its counter.
    - `_without_own_revisions` and `_OWN_REVISION_RX`.
  - `spec_formatter/style_application/core/xml_helpers.py`:
    `paragraph_run_content_signature`, `run_content_difference`.
  - `spec_formatter/style_application/batch_runner.py`:
    `_apply_classified_target_impl` (where each converter's result is
    consumed) and `_build_and_patch_output` (which calls the gate).
  - `spec_formatter/style_application/core/csi_to_canadian.py`:
    `plan_csi_to_canadian`, `apply_csi_to_canadian`, `MarkerEdit`,
    `CanadianConversionReport`, `ConversionPlan`.
  - `spec_formatter/style_application/core/marker_tools.py`:
    `_verify_changed_paragraph`.
  - `spec_formatter/style_application/core/canadian_to_csi.py`:
    `plan_canadian_to_csi`, `apply_canadian_to_csi`, its `ConversionPlan`,
    `_verify_prediction`, `MARKER_REVISION_AUTHOR`.
  - Tests:
    - `tests/test_unified_roundtrip.py`: `_deterministic_classifier`,
      `_write_canadian_pair`, and the new `_write_run_content_pair` with its
      late-corruption test.
    - `tests/test_architect_free_modes.py`, `tests/test_canadian_to_csi.py`,
      `tests/style_application_regression/test_final_package_validation.py`.
- Pitfalls already discovered that bear on this item:
  - **Tracked markers.** The tracked reverse conversion writes each marker
    as its own run inside a `w:ins` authored `MARKER_REVISION_AUTHOR`.
    - `_without_own_revisions` removes those `w:ins` blocks. Compute the
      signature *after* that projection, or every tracked paragraph gains a
      run.
    - The `w:pPrChange` the same conversion writes needs no projection for
      the signature. `strip_out_of_scope_subtrees` already removes
      `w:pPrChange`, so it never reaches the signature.
  - **Untracked markers.** An untracked reverse marker is inserted into the
    paragraph's first text run, and the forward converter edits `w:t`
    contents. Converted paragraphs therefore change by signature and must be
    in the predicted set; they are never compared for identity.
  - **Keep the counter.** Keep the `body_signature_paragraphs_compared` key
    and its "on the failure path too" behaviour when generalizing. The plan
    uses the same name and adds `body_paragraphs_expected_changed`.
  - **Error codes.** `conversion_prediction_mismatch` already exists and is
    what the Canadian modes should raise. Format-only keeps its
    `FORMAT_ONLY INVARIANT FAIL:` messages, per the plan. The
    `ERROR_REMEDIATIONS` bound (15..22, with 22 codes) should not need to
    move. Adjacent finding 1 above is the user's call, not this item's.
  - **Mutating the package in tests.** To corrupt the output after the
    converters, wrap `batch_runner._build_and_patch_output` and edit
    `extract_dir/word/document.xml` with `read_bytes`/`write_bytes` before
    calling the real one. Session 02's
    `test_format_only_withholds_output_when_run_content_changes_after_application`
    does exactly this. Use bytes, not text: `Path.write_text` translates
    newlines on Windows.
  - **Review bot.** The automated Codex reviewer on this repository posts
    P1/P2 findings as review threads. Sessions 00 and 01 treated them as bug
    reports to verify and fix, not as optional.

## End of session

- One pull request for this session, containing only `WI-03` and its
  bookkeeping. Follow the **session end ritual** in the plan (section 3.3):
  run the checks, tick the boxes, set the tracker row to `in_review` with the
  PR URL, write `docs/docx_method_hardening/handoffs/handoff-for-session-04.md`
  from the template, paste it in chat, subscribe to the PR, drive CI to green,
  and paste the post-merge version of the handoff when the PR merges.
- `WI-03` is not the last required item, so no completion banner this session.
