# Vendored Phase 2 regression tests

These tests were copied from `spec-style-applier/tests` at commit
`545aa8f69a3bfd397155879a4819b144ecdd5f0b`.

Imports and string-based `monkeypatch` targets were mechanically rewritten from
the source repository's top-level module layout to
`spec_formatter.style_application`. This keeps the suite self-contained and
prevents aliases such as `core`, `batch_runner`, or `docx_decomposer` from
overwriting the Phase 1 modules in `sys.modules` during a combined test run.

The two legacy standalone-GUI tests remain as explicit skips because that GUI was
intentionally replaced by the unified application and is not part of the
vendored engine. All other upstream tests, including the real Phase 1-to-Phase 2
round trip, execute against the namespaced implementation.

The round-trip fixture places its first section break in a dedicated empty
paragraph. Current Phase 1 correctly excludes section-break paragraphs from
classification, so this preserves the upstream fixture's intent of presenting
two classifiable content paragraphs.

When refreshing this directory from upstream, copy `test_*.py` and apply these
import-prefix rewrites again; do not add compatibility aliases to `sys.path` or
`sys.modules`.
