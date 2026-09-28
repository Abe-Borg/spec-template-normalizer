# Handoff prompt for session 08

Paste everything below the horizontal rule into the next session as its
first message. Written by session 07 on 2026-09-28 UTC.

---

You are continuing the **DOCX Method Hardening** program in the repository
`Abe-Borg/spec-template-normalizer`. Read these files completely before your
first tool call, in this order:

1. `CLAUDE.md` (engineering guide; every rule in it applies)
2. `docs/docx_method_hardening/DOCX_METHOD_HARDENING_PLAN.md`
3. `docs/docx_method_hardening/PROGRESS_TRACKER.md`
4. This prompt

## State at handoff

- Previous session: `07`, work item `WI-07: Marker placement before leading
  structural run content`.
- Pull request: `https://github.com/Abe-Borg/spec-template-normalizer/pull/67`.
  Merge status when this prompt was written: `open, CI pending`.
- Tracker rows changed by the previous session: `WI-06` -> `merged` with
  merge commit `4f55ef5`; `WI-07` -> `in_review` with the PR URL; session
  log row `07` added.
- Verification the previous session ran, with results:
  - `pytest` was not installed in the container. Run
    `pip install -r requirements-dev.txt` before the baseline.
  - Baseline before any change: `python -m pytest -q` gave 1640 passed,
    1 skipped. The skip is the GUI test, which needs `tkinter`.
  - The WI-07 tests were committed first (`486e240`). Against the unchanged
    engine, 18 failed:
    - `tests/test_canadian_to_csi.py`: 13. Seven untracked shapes put the
      marker behind a leading tab (in the text run or in an earlier run with
      no text), a line break, a positional tab, a symbol or a carriage
      return. Three tracked shapes put it behind a tab or positional tab in
      an earlier run with no text. Three tests of the new structural checks.
    - `tests/test_conversion_verification.py`: 4 (the enumerated text diff,
      the accept-all view, and two new element-level placement checks).
    - `tests/test_architect_free_modes.py`: 1. End to end, the misplaced
      marker published as a success.
  - After the implementation:
    - the four changed test modules plus `tests/test_conversion_prediction.py`:
      106 passed;
    - `python -m pytest -q`: 1670 passed, 1 skipped;
    - `python -m pytest tests/test_sanitized_format_only_corpus.py -q`:
      2 passed;
    - `tests/test_engine_identity.py`: passes (no fingerprinted file
      touched); the tracker test: 6 passed;
    - both probes exit 0;
    - on Python 3.10 the application imports and the changed and gate test
      modules pass (143 passed); pyflakes reports nothing new (its two
      warnings on `canadian_to_csi.py`, unused `Callable` and
      `ErrorLocation`, are on `master` too).
    - A 3.10 interpreter is at `/usr/bin/python3.10`. `uv venv -p
      /usr/bin/python3.10` plus `uv pip install -r requirements-dev.txt`
      gives a working environment in seconds.
- Anything left unfinished, unexpected, or decided along the way:
  - **Decided.**
    - **A plan refinement: typed markers are replaced in place.** The plan's
      rule (the marker leads the run content) is right for an automatically
      numbered paragraph. It is wrong for a typed Canadian marker being
      replaced: that marker stood where its author typed it, often behind
      indentation tabs typed in front of it (`<tab>.1<tab>Text`). So
      `_insert_marker(typed=...)` keeps the typed path in front of the first
      `w:t`. Only the automatic path moved. The plan's *as implemented* note
      and its first Definition of done box say so. Typed replacement under
      tracking still fails closed, so every tracked marker is automatic.
    - **Placement.** `_first_run_content` finds the first run, in document
      order, with any direct child except `w:rPr` and
      `w:lastRenderedPageBreak`. Untracked, the marker goes before that
      first content child. Tracked, the `w:ins` goes before that run and
      copies its `w:rPr`.
    - **Refusal unchanged.** `_TRACKED_OR_FIELD_RX` still searches everything
      before the first `w:t`, which encloses the new, earlier insertion
      point. A leading tab inside a tracked insertion or simple field, or
      after a complex field, is refused as before. That is conservative for
      the last case (the marker could be written ahead of the field), and
      was kept on purpose.
    - **Prediction.** `_predict_marker_insertion(typed=...)` restates the
      rule on the signature: position 0 of the first non-empty run, or
      before the first `t` item when typed.
    - **Structural post-check.** `_marker_stands_first` reads the signature
      and requires the marker's own text node to come first (automatic) or
      to be the first text node (typed). `_verify_marked_paragraph(typed=...)`
      and `_verify_prediction(typed_markers=...)` run it after their text
      checks. A failure is `conversion_prediction_mismatch` (an application
      defect), not `canadian_to_csi_hierarchy`.
    - **Tests adjusted, stricter rather than looser.** Three older
      `_verify_prediction` unit tests had fake output with the marker and
      its text in one `w:t` joined by a literal tab character. The converter
      never writes that shape, and the structural check refuses it. They now
      use the converter's real shape.
    - **No contract change.** No new error code, stage or policy field.
      `APPLICATION_POLICY_VERSION` is unchanged.
  - **Found, not fixed.** Raise these with the user; do not fold them into
    WI-08 unasked.
    1. **Leading inline drawings and math** (new). A run holding only a
       drawing, picture or object is not run content: the edit sees it as a
       placeholder and the signature strips it. So a marker still follows a
       leading inline drawing. Math zones (`m:oMath`, `m:r`) are not `w:r`
       either. Both are rare at the start of a numbered heading. A fix needs
       a placement signal the signature does not carry.
    2. **Forward converter keeps a tab typed before a CSI marker** (new).
       `csi_to_canadian` removes a typed marker and its delimiter tab, but
       not a tab typed *in front of* the marker. Word then renders that tab
       after the automatic Canadian number, so a round trip of
       `<tab>A.<tab>Text` returns `A.<tab><tab>Text`. The reverse converter
       now faithfully reproduces what the Canadian file shows. The question
       is whether the forward conversion should drop leading indentation it
       replaces with list geometry. That is a behaviour decision for the
       user.
    3. **Format-only drops a `w:numberingChange`** (handoff 07 finding 1).
       The census withholds such a target (`untrusted_error`). The fix
       belongs in the numbering materialization in `core/classification.py`.
    4. **Non-standard settings or font table part names** (handoff 05
       finding 1): `apply_settings` and `apply_font_table` edit
       `word/settings.xml` and `word/fontTable.xml` by name. The WI-05
       whitelist admits those names; narrowing it belongs with the fix.
    5. **Settings order** (handoff 04 finding 5): `apply_settings` still
       inserts a missing `w:compat` just before `</w:settings>`.
    6. **Parity log lines.** WI-04's `Even/odd headers: ...` lines have no
       prefix in `_SAFE_OPERATIONAL_PREFIXES`, so `run.log` shows them only
       as `[untrusted detail omitted; sha256=...]`.
    7. **Handoff 04 findings 1 to 4 and 6** still stand:
       - tracked `w:rPrChange` on a heading run (`_run_properties` copies
         the protected-subtree placeholder into the marker run; WI-07 now
         copies the first content run's `w:rPr`, same defect);
       - the doubled gap after `PART n`;
       - `html.unescape` in the forward edit;
       - `_TRACK_REVISIONS_RX` reads only double quotes;
       - session 02's and session 01's findings.

## Your assignment: `WI-08: Independent stdlib verifier for the other three modes`

- Follow the **session start ritual** in the plan (section 3.2) first:
  confirm the PR above is merged; if it is not, stop and tell the user rather
  than building on a stale base. Record the merge in the tracker
  (`WI-07` -> `merged`, with the merge commit), then mark `WI-08`
  `in_progress` with session `08`.
  - The session log row you add makes the tracker test require
    `handoffs/handoff-for-session-09.md`.
  - Commit a short placeholder with the right title line in your first
    commit, as sessions 01-07 did, and replace it before review.
- The plan section for `WI-08` is the specification. Its
  **Definition of done** checklist is what you must complete and tick. The
  plan rates it effort M.
- Key files:
  - `tests/test_conversion_verification.py`: the existing independent
    verifier for `canadian_to_csi`. It shares no code with the engine and
    reads packages with `zipfile` and `xml.etree.ElementTree` only. Its
    `_first_run_content`, `_walk_text` and `_annotations` show the style.
    Model the new module on it; do **not** import from it into the engine
    or from the engine into it.
  - `tests/test_unified_roundtrip.py`: `_deterministic_classifier`
    (line 280) and `_write_docx(path, *, architect=...)` drive `format_only`
    and `csi_to_canadian` through `format_specifications` with no API key.
    `_write_canadian_pair` and `_write_run_content_pair` build other shapes.
  - `tests/test_architect_free_modes.py`: `_write_docx`, `_run` and
    `CSI_LINES` drive `csi_to_canadian_standalone` (and `canadian_to_csi`)
    with no API key and no architect.
  - The public entry point is `spec_formatter.format_specifications`; it is
    the only engine import the new module may use (the plan allows it to
    *produce* the output). Pipeline mode constants such as
    `CSI_TO_CANADIAN_STANDALONE` live in `spec_formatter.pipeline`. Passing
    the mode string literally avoids importing them.
- Pitfalls already discovered that bear on this item:
  - **Predictions are hand-written.** The plan requires the expected text
    of every paragraph (with `\t` and breaks spelled out) and the package
    whitelist to be written in the test source before the engine runs. It
    must be computed from the mode, never from `ApplicationPolicy`, and
    never derived from output.
  - **What each mode may change.** WI-05's remit:
    - `format_only` and architect `csi_to_canadian`: the body, the styles,
      the numbering, the shell (settings, theme, font table, content types,
      relationships) and the header/footer set the importer reports;
    - `csi_to_canadian_standalone`: the body, plus its added and wired
      numbering (`word/numbering.xml`, `[Content_Types].xml`,
      `word/_rels/document.xml.rels`); never the styles.

    Write your own list for your fixture. Media and header names are
    allocated by the importer, so match header/footer parts by pattern or by
    the output's relationships, not by fixed names.
  - **Forward markers.** `csi_to_canadian` removes a `- ` or `: ` after
    `PART n` with the marker (`PART 1 - GENERAL` becomes `GENERAL`). When
    the separator continues in the next text node, the delimiter tab stays
    (handoff 04 finding 2). Hand-written expectations must reflect that, or
    avoid the shape.
  - **Revision census.** WI-06 counts every revision element by author and
    kind in `word/document.xml`. Every mode keeps it unchanged except
    tracked `canadian_to_csi`, which WI-08 does not cover. Count
    independently (`_REVISION_ELEMENTS` in the existing verifier is a model).
    Revision-id uniqueness is scoped to revision elements, because paired
    annotations share ids.
  - **Settings parity (WI-04).** With the architect's header set imported,
    the output's `w:evenAndOddHeaders` follows the architect. With the
    target keeping its own set (standalone mode), the target's switch is
    byte-identical. Find the settings part through the document's
    relationship, not by name.
  - **Header/footer text.** Format-only's guarantee is scoped to the body.
    Header and footer wording is replaced by the architect's, with section
    tokens patched, so do not hold it to identity.
  - **Encodings.** Fixture parts are often UTF-16. Parse with
    `ET.fromstring(bytes)` and compare bytes or parsed XML, never text
    decoded with an assumed encoding.
  - **"Every output XML part parses with its declared namespaces"**:
    `ET.fromstring` on each part's bytes already fails on an unbound
    prefix. Include parts with a non-`.xml` extension that
    `[Content_Types].xml` declares as XML (WI-06 learned this).
  - **Exact `verification_out` assertions** in
    `tests/style_application_regression/test_final_package_validation.py`
    list every gate counter. WI-08 should add no engine counter; if it does,
    add it there with exact values.
  - **Review bot.** The automated Codex reviewer on this repository posts
    P1/P2 findings as review threads. Earlier sessions treated them as bug
    reports to verify and fix, not as optional.

## End of session

- One pull request for this session, containing only `WI-08` and its
  bookkeeping. Follow the **session end ritual** in the plan (section 3.3):
  run the checks, tick the boxes, set the tracker row to `in_review` with the
  PR URL, write `docs/docx_method_hardening/handoffs/handoff-for-session-09.md`
  from the template, paste it in chat, subscribe to the PR, drive CI to green,
  and paste the post-merge version of the handoff when the PR merges.
- `WI-08` is the **last required item**. After its PR merges, follow plan
  section 3.5:
  - paste the post-merge handoff, which assigns session 09 the closeout
    (a bookkeeping-only PR that records the merge and flips
    `Program status:` to `PROGRAM COMPLETE`);
  - print the completion banner exactly as written there. Never print it
    while any required item is not merged.
