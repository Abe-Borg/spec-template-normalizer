# Style-application regression tests

These tests exercise the target-side engine in
`spec_formatter.style_application`. They began as a copy of the separate
`spec-style-applier` repository's suite, taken at commit
`545aa8f69a3bfd397155879a4819b144ecdd5f0b` when the two applications were
merged, and have been maintained here since. There is no upstream to refresh
them from: re-copying would discard every change made in this repository.

Import engine modules through `spec_formatter.style_application`, and do not
add compatibility aliases to `sys.path` or `sys.modules`. The engine and the
root Phase 1 code each have a `docx_decomposer` and an `llm_classifier`
module, so an alias lets a bare import resolve to whichever one loaded first
during a combined test run.

`test_phase1_to_phase2_roundtrip.py` runs Phase 1 from this checkout in a
child interpreter and feeds the published bundle to the engine. Its architect
fixture puts the first section break in a dedicated empty paragraph, which
Phase 1 skips as empty, so exactly two content paragraphs are classifiable; a
section break on a paragraph with visible text would stay classifiable.
