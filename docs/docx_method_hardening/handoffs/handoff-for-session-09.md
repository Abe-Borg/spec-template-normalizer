# Handoff prompt for session 09

Paste everything below the horizontal rule into the next session as its
first message. Written by session 08 on 2026-09-28 UTC.

---

You are continuing the **DOCX Method Hardening** program in the repository
`Abe-Borg/spec-template-normalizer`. Read these files completely before your
first tool call, in this order:

1. `CLAUDE.md` (engineering guide; every rule in it applies)
2. `docs/docx_method_hardening/DOCX_METHOD_HARDENING_PLAN.md`
3. `docs/docx_method_hardening/PROGRESS_TRACKER.md`
4. This prompt

## State at handoff

- Previous session: `08`, work item `WI-08: Independent stdlib verifier for
  the other three modes`. **WI-08 is the last required item.**
- Pull request: `https://github.com/Abe-Borg/spec-template-normalizer/pull/68`.
  Merge status when this prompt was written: `open, CI pending`.
- Tracker rows changed by the previous session: `WI-07` -> `merged` with
  merge commit `e6e30e0`; `WI-08` -> `in_review` with the PR URL; session
  log row `08` added.
- Verification the previous session ran, with results:
  - `pytest` was not installed in the container. Run
    `pip install -r requirements-dev.txt` before the baseline.
  - Baseline before any change: `python -m pytest -q` gave 1670 passed,
    1 skipped. The skip is the GUI test, which needs `tkinter`.
  - After the change:
    - `tests/test_independent_verification_all_modes.py`: 76 passed
      (3 modes x (9 checks + 15 mutation self-tests), 3 prediction-table
      checks, 1 import check);
    - `python -m pytest -q`: 1746 passed, 1 skipped;
    - `python -m pytest tests/test_sanitized_format_only_corpus.py -q`:
      2 passed;
    - `tests/test_engine_identity.py` (no fingerprinted file touched) and
      the tracker test: 10 passed;
    - both probes exit 0;
    - on Python 3.10 the new module, `tests/test_conversion_verification.py`
      and the tracker test pass (100 passed); pyflakes is clean on the new
      module.
    - A 3.10 interpreter is at `/usr/bin/python3.10`. `uv venv -p
      /usr/bin/python3.10` plus `uv pip install -r requirements-dev.txt`
      gives a working environment in seconds.
- Anything left unfinished, unexpected, or decided along the way:
  - **Decided.**
    - **Self-contained.** The new module writes its own fixtures, its own
      architect classifier and its own readers, and imports nothing from
      another test module either. A test parses its imports: only the
      standard library, `pytest`, and `from spec_formatter import
      format_specifications`.
    - **Beyond the plan's six checks:** paragraphs a mode does not edit are
      held to element identity (not only text); tables in every mode and
      section properties without a shell likewise; the imported headers and
      footers read as the architect's; both inputs are byte-identical after
      the run.
    - **The checks can fail.** Each check is a function, and fifteen
      mutations of every mode's real output must each be rejected by the
      check that exists to find it. Two mutations first exposed gaps in the
      verifier itself (a `ParseError` escaping as a non-assertion, and a
      `w:t` inside a deletion rendering like `w:delText`); both were fixed
      before the PR.
    - **Predicted first.** The two architect modes matched the hand-written
      predictions on their first run. The standalone run was withheld by the
      engine (finding 1 below), so its fixture gives the target a numbering
      part of its own, and a comment says why.
    - No engine change; no new code, stage, policy field or counter.
  - **Found, not fixed.** Put these to the user; they are theirs to
    schedule. Do not fold them into the closeout.
    1. **False positive withholds runs that keep the target's headers**
       (new, found by WI-08). `_verify_target_header_footer_preserved`
       (`phase2_invariants.py`) compares the target's header and footer
       `<Relationship>` elements as raw text
       (`_extract_hf_relationship_subset`). The engine's own writers
       re-serialize `word/_rels/document.xml.rels` with ElementTree, which
       writes `Target="header9.xml" />`. So whenever a relationship is added
       to a target that keeps its own header set, the run fails with
       `INVARIANT FAIL: relationship subset changed`. Reproduced:
       - `csi_to_canadian_standalone` on a target with a header and no
         numbering part (`_ensure_numbering_package_wiring`). That is a
         typical typed-marker spec with no Word lists, so the standalone mode
         refuses most real inputs;
       - `csi_to_canadian` when the architect has no header parts and the
         target no numbering part;
       - `format_only` when the architect has no header parts and the target
         no theme (`_ensure_relationship_in_document_rels`).

       It fails closed; nothing bad publishes. The fix is small: compare the
       relationships parsed (Id, Type, resolved Target, TargetMode), as the
       `_targets` half of the same function already does. Add the
       reproduction first: in `tests/test_independent_verification_all_modes.py`,
       give the standalone target `numbering=False` (and move `"word/numbering.xml"`
       from `changed` to `added`, with `[Content_Types].xml` and the document
       relationships in `changed`).
    2. **Imported header revision ids can collide with the body's** (new,
       found by WI-08). The header/footer importer keeps an architect part's
       revision ids as they were, so an imported pending revision can share
       its `w:id` with one in the target's body; reproduced with `w:id="11"`
       in `word/document.xml` and `word/header2.xml`. It publishes as a
       success. WI-06 counts and warns about imported revisions; it does not
       renumber them.
    3. **Leading inline drawings and math** (handoff 08 finding 1). A marker
       still follows a leading inline drawing; math zones are not `w:r`.
    4. **Forward converter keeps a tab typed before a CSI marker** (handoff
       08 finding 2). A round trip of `<tab>A.<tab>Text` returns
       `A.<tab><tab>Text`. A behaviour decision for the user.
    5. **Format-only drops a `w:numberingChange`** (handoff 07 finding 1).
       The census withholds such a target (`untrusted_error`). The fix
       belongs in the numbering materialization in `core/classification.py`.
    6. **Non-standard settings or font table part names** (handoff 05
       finding 1): `apply_settings` and `apply_font_table` edit
       `word/settings.xml` and `word/fontTable.xml` by name.
    7. **Settings order** (handoff 04 finding 5): `apply_settings` still
       inserts a missing `w:compat` just before `</w:settings>`.
    8. **Parity log lines.** WI-04's `Even/odd headers: ...` lines have no
       prefix in `_SAFE_OPERATIONAL_PREFIXES`, so `run.log` shows them only
       as `[untrusted detail omitted; sha256=...]`.
    9. **Handoff 04 findings 1 to 4 and 6** still stand: a tracked
       `w:rPrChange` on a heading run (`_run_properties` copies the
       protected-subtree placeholder into the marker run); the doubled gap
       after `PART n`; `html.unescape` in the forward edit;
       `_TRACK_REVISIONS_RX` reads only double quotes; session 02's and
       session 01's findings.

## Your assignment: program closeout (bookkeeping only)

Every required work item is merged once PR 68 is. Your session is the
**closeout** described in plan section 3.5, not a work item.

- Follow the **session start ritual** in the plan (section 3.2) first:
  confirm PR 68 is **merged**; if it is not, stop and tell the user rather
  than building on a stale base.
- Your pull request is bookkeeping only, in one commit if you like:
  - `PROGRESS_TRACKER.md`: `WI-08` -> `merged` with its merge commit (the
    SHA on `master`), and `Program status:` -> `PROGRAM COMPLETE`. Do both
    in the same commit: the tracker test fails in either direction when the
    status line and the required table disagree.
  - Add session log row `09` with the work item `closeout` (the tracker test
    accepts exactly that word) and the handoff `handoffs/handoff-for-session-10.md`.
  - Write `handoffs/handoff-for-session-10.md`, starting with the line
    `# Handoff prompt for session 10`. It says the program is complete and
    lists what the user may still choose: the optional items `WI-09` (stale
    cross-reference and field-result sweep) and `WI-10` (revision dates in
    Word's convention, which needs the user's decision first), and the
    found-not-fixed list above, finding 1 first.
  - No code, test or documentation change beyond these. Run the tracker test
    and the full suite anyway, and record the results.
- Key files: `docs/docx_method_hardening/PROGRESS_TRACKER.md`,
  `docs/docx_method_hardening/handoffs/`,
  `tests/test_docx_method_hardening_tracker.py` (read its assertion
  messages before editing).
- Pitfalls already discovered that bear on this item:
  - **The status line is enforced both ways.** `PROGRAM COMPLETE` while any
    required row is not `merged` or `dropped` fails, and so does
    `IN PROGRESS` once all are.
  - **Review bot.** The automated Codex reviewer on this repository posts
    P1/P2 findings as review threads. Earlier sessions treated them as bug
    reports to verify and fix, not as optional.

## End of session

- One pull request for this session, containing only the closeout
  bookkeeping. Follow the **session end ritual** in the plan (section 3.3)
  where it applies: run the checks, open the PR against `master`, paste the
  prompt from `handoff-for-session-10.md` in chat, subscribe to the PR,
  drive CI to green.
- When the closeout PR merges, paste the post-merge version of the session
  10 handoff, and print the completion banner from plan section 3.5 again,
  exactly as written there.
