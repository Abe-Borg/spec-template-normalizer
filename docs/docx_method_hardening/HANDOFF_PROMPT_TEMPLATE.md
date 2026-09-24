# Handoff prompt template

Every session of the DOCX Method Hardening program ends by writing
`docs/docx_method_hardening/handoffs/handoff-for-session-NN.md`, where `NN` is
the **next** session's two-digit number, and by pasting that file's prompt
into chat verbatim. The same prompt is pasted again, with the merge line
updated, when the session's PR merges. `tests/test_docx_method_hardening_tracker.py`
fails if the file is missing or its title does not match its filename.

Copy everything below the rule into the new file. Replace every `<...>`.
Keep the headings: the next agent reads them in order and the tracker test
checks the title line.

---

# Handoff prompt for session <NN>

Paste everything below the horizontal rule into the next session as its
first message. Written by session <NN-1> on <YYYY-MM-DD UTC>.

---

You are continuing the **DOCX Method Hardening** program in the repository
`Abe-Borg/spec-template-normalizer`. Read these files completely before your
first tool call, in this order:

1. `CLAUDE.md` (engineering guide; every rule in it applies)
2. `docs/docx_method_hardening/DOCX_METHOD_HARDENING_PLAN.md`
3. `docs/docx_method_hardening/PROGRESS_TRACKER.md`
4. This prompt

## State at handoff

- Previous session: `<NN-1>`, work item `<WI-xx: title>`.
- Pull request: `<url>`. Merge status when this prompt was written:
  `<open, CI green | open, CI red because ... | merged as <sha>>`.
- Tracker rows changed by the previous session: `<list>`.
- Verification the previous session ran, with results:
  `<commands and outcomes, including the full suite and the corpus regression>`.
- Anything left unfinished, unexpected, or decided along the way:
  `<be specific; "nothing" is an acceptable answer>`.

## Your assignment: `<WI-yy: title>`

- Follow the **session start ritual** in the plan (section 3.2) first:
  confirm the PR above is merged; if it is not, stop and tell the user rather
  than building on a stale base. Record the merge in the tracker
  (`<WI-xx>` -> `merged`, with the merge commit), then mark `<WI-yy>`
  `in_progress` with session `<NN>`.
- The plan section for `<WI-yy>` is the specification. Its
  **Definition of done** checklist is what you must complete and tick.
- Key files: `<paths>`.
- Pitfalls already discovered that bear on this item: `<list or "none">`.

## End of session

- One pull request for this session, containing only `<WI-yy>` and its
  bookkeeping. Follow the **session end ritual** in the plan (section 3.3):
  run the checks, tick the boxes, set the tracker row to `in_review` with the
  PR URL, write `docs/docx_method_hardening/handoffs/handoff-for-session-<NN+1>.md`
  from the template, paste it in chat, subscribe to the PR, drive CI to green,
  and paste the post-merge version of the handoff when the PR merges.
- If `<WI-yy>` is the last required item, print the completion banner from
  plan section 3.5 after the merge, exactly as written there.
