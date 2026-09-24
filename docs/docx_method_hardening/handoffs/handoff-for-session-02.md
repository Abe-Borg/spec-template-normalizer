# Handoff prompt for session 02

Paste everything below the horizontal rule into the next session as its
first message. Written by session 01 on 2026-09-24 UTC.

---

You are continuing the **DOCX Method Hardening** program in the repository
`Abe-Borg/spec-template-normalizer`. Read these files completely before your
first tool call, in this order:

1. `CLAUDE.md` (engineering guide; every rule in it applies)
2. `docs/docx_method_hardening/DOCX_METHOD_HARDENING_PLAN.md`
3. `docs/docx_method_hardening/PROGRESS_TRACKER.md`
4. This prompt

## State at handoff

- Previous session: `01`, work item `WI-01: Extension-namespace-safe style
  import and shell application`.
- Pull request: `https://github.com/Abe-Borg/spec-template-normalizer/pull/61`.
  Merge status when this prompt was written: `open, CI running`.
- Tracker rows changed by the previous session: `WI-00` -> `merged` with
  merge commit `f2da16a`; `WI-01` -> `in_review` with the PR URL; session
  log row `01` added.
- Verification the previous session ran, with results:
  - Baseline before any change: `python -m pytest -q` gave 1306 passed,
    1 skipped.
  - The WI-01 tests were committed first (fc2d40c). 21 of them failed, each
    for the defect's own reason: Format-only stopped at `style_import` with
    `unbound prefix` and no error code; `csi_to_canadian` stopped at
    `classification_application` because the stylesheet was already
    ill-formed, and silently accepted a target that bound `w14` to a foreign
    URI.
  - After the fix, including the two review follow-ups below: 97 new and
    affected tests passed.
  - `python -m pytest -q`: 1372 passed, 1 skipped. The skip is the GUI test,
    which needs `tkinter`.
  - `python -m pytest tests/test_sanitized_format_only_corpus.py -q`: 2
    passed.
  - `python -m pytest tests/test_engine_identity.py -q`: 4 passed. No
    fingerprinted file was touched.
  - `probe_style_import_w14.py`: both rows OK, exit 0.
  - `probe_format_only_gate.py`: unchanged, 7 mutations ACCEPTED, control
    REJECTED, exit 1. That is your defect.
- Anything left unfinished, unexpected, or decided along the way:
  - **Decided.**
    - `apply_environment_to_target` now takes a required keyword argument,
      `architect_namespaces`.
    - `import_arch_styles_into_target` takes the same keyword, but it is
      optional. When omitted, it is derived from `arch_styles_xml`, which
      gives the same answer.
    - One `RootNamespaces` value is built per target in
      `_apply_classified_target_impl` and passed to both.
    - A prefix the architect root does not declare fails with the same code
      as a conflicting binding, `style_import_namespace_conflict`.
    - A default namespace is never added or changed.
    - A Strict Open XML target or template binds `w` to a different URI, so
      it now fails closed instead of mixing namespaces.
  - **Changed after review.** The Codex reviewer raised two P2 findings on
    PR 61, both fixed in the PR, and the plan's WI-01 text and third box were
    amended to match:
    - Materialized properties are keyed by **expanded** name (namespace URI
      and local name, resolved through the architect root's declarations),
      not by the prefix as written. Two prefixes bound to one namespace name
      one property.
    - `prefixes_used` also counts prefixes named in markup-compatibility
      values: `Requires` on `mc:Choice`, `Ignorable`, `MustUnderstand`,
      `ProcessContent`, `PreserveElements` and `PreserveAttributes`. These
      are recognised by namespace, with the source root's declarations
      passed as `context`.
  - **Unexpected, fixed in the same PR.** ElementTree used to normalize
    quoting inside materialized style children. Raw fragments keep the
    architect's quoting, and the geometry invariant's `_ind_from_ppr_xml`
    read only double-quoted `w:ind` attributes. It now reads either quoting,
    with tests for both the false failure and the false pass.
  - **Also added.** The new namespace step reports what it did:
    - a `run.log` line naming the declared prefixes;
    - `StyleImportResult.declared_namespace_prefixes`;
    - a `styles_namespace_additions` count on the `apply_environment` and
      `style_import` diagnostics events.
  - **Plan text fixed.** The first WI-01 box said "both probes'
    `format_only` rows"; only the w14 probe has them.
  - **Adjacent defects found and deliberately left alone** (outside WI-01's
    remit; none is scheduled; raise them with the user rather than folding
    them into another item):
    1. `numbering_importer.inject_numbering_into_xml` merges the architect
       numbering root's declarations by prefix only. A target prefix bound
       to a different URI is silently kept, so architect `w15:` attributes
       then mean something else. `mc:Ignorable` is never merged either.
       Session 01 reproduced both.
       `ensure_root_declarations` / `declare_fragment_namespaces` in
       `core/xml_helpers.py` are the tools to fix it.
    2. `_extract_tag_inner` in `core/style_import.py` is a non-greedy regex.
       It truncates a style's `w:rPr`/`w:pPr` at a nested
       `w:rPrChange/w:rPr` or `w:pPrChange/w:pPr` (a tracked
       style-definition change), so such an architect style still fails
       Format-only. It now fails as `Architect style 'X' contains malformed
       w:rPr XML`, which session 01 confirmed. For a table style with no
       rPr of its own, the same regex picks up a `w:tblStylePr` rPr.
    3. Materialization keeps first-seen order for `w:` children (the plan
       said to), which is not the `CT_RPr`/`CT_PPr` schema sequence.
    4. Phase 1's `_propagate_style_fragment_namespaces` in root
       `docx_decomposer.py` (fingerprinted) copies declarations for derived
       role styles but not their `mc:Ignorable` tokens.

## Your assignment: `WI-02: Exact run-content signature at the Format-only gate`

- Follow the **session start ritual** in the plan (section 3.2) first:
  confirm the PR above is merged; if it is not, stop and tell the user rather
  than building on a stale base. Record the merge in the tracker
  (`WI-01` -> `merged`, with the merge commit), then mark `WI-02`
  `in_progress` with session `02`. The session log row you add makes the
  tracker test require `handoffs/handoff-for-session-03.md`. Session 01
  committed a short placeholder with the right title line in its first
  commit, and replaced it before review.
- The plan section for `WI-02` is the specification, with appendix A as the
  signature's exact contents. Its **Definition of done** checklist is what
  you must complete and tick.
- Key files:
  - `spec_formatter/style_application/phase2_invariants.py`:
    `_verify_format_only_body_invariants`, `verify_phase2_invariants`.
  - `spec_formatter/style_application/core/xml_helpers.py`:
    `paragraph_text_from_block`, `strip_out_of_scope_subtrees`,
    `xml_unescape`, `_scan_markup`.
  - `spec_formatter/style_application/core/classification.py`:
    `_normalize_paragraph_for_contract`.
  - `docs/docx_method_hardening/probes/probe_format_only_gate.py`.
  - `tests/style_application_regression/test_final_package_validation.py`.
  - `CLAUDE.md` invariant 2 and `README.md`.
- Pitfalls already discovered that bear on this item:
  - **Unescaping.** The plan asks for an XML-only unescape.
    `core/xml_helpers.xml_unescape` now exists (the five predefined entities
    plus numeric references, never HTML entities) with a test. Reuse it
    rather than writing a second one.
  - **Scanning.** `_scan_markup` is now a module-level, quote-aware tag
    scanner in `core/xml_helpers.py`, shared by `iter_direct_child_xml_blocks`
    and the namespace helpers. `iter_element_xml_blocks` keeps its own inline
    copy on purpose: it is the engine's hot path.
  - **Message contract.** `test_final_package_validation.py` matches on
    `target body text changed`. The plan requires keeping that message and
    adding the run-content one beside it.
  - **Error codes.** `tests/style_application_regression/test_engine_errors.py`
    now bounds `ERROR_REMEDIATIONS` at `15 <= len <= 22`, with 22 codes.
    WI-02's failures are `FORMAT_ONLY INVARIANT FAIL: ...` messages, not new
    codes, so the bound should not need to move.
  - **Review bot.** The automated Codex reviewer on this repository posts
    P1/P2 findings as review threads. Session 00 treated them as bug reports
    to verify and fix, not as optional.
  - **Corpus.** A corpus failure after your change most likely means the
    signature counted `w:pPr/w:tabs` or something inside `w:txbxContent`.
    The plan says so, and nothing session 01 saw contradicts it.

## End of session

- One pull request for this session, containing only `WI-02` and its
  bookkeeping. Follow the **session end ritual** in the plan (section 3.3):
  run the checks, tick the boxes, set the tracker row to `in_review` with the
  PR URL, write `docs/docx_method_hardening/handoffs/handoff-for-session-03.md`
  from the template, paste it in chat, subscribe to the PR, drive CI to green,
  and paste the post-merge version of the handoff when the PR merges.
- `WI-02` is not the last required item, so no completion banner this session.
