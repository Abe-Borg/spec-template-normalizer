"""The committed engine fingerprint must match the analysis engine's source."""

from __future__ import annotations

from pathlib import Path

import engine_identity
from engine_identity import (
    ENGINE_SOURCE_DIGEST,
    ENGINE_SOURCE_FILES,
    compute_engine_source_digest,
)


REPO_ROOT = Path(__file__).resolve().parents[1]


def test_committed_engine_digest_matches_the_checkout():
    recomputed = compute_engine_source_digest(REPO_ROOT)
    assert recomputed == ENGINE_SOURCE_DIGEST, (
        "One of the architect-analysis engine files changed "
        f"({', '.join(ENGINE_SOURCE_FILES)}) without updating "
        "engine_identity.ENGINE_SOURCE_DIGEST. Cached template profiles depend on "
        f"it; run `python engine_identity.py` and set the constant to {recomputed!r}."
    )


def test_engine_digest_is_line_ending_independent(tmp_path: Path):
    for name in ENGINE_SOURCE_FILES:
        (tmp_path / name).write_bytes((REPO_ROOT / name).read_bytes().replace(b"\n", b"\r\n"))
    assert compute_engine_source_digest(tmp_path) == ENGINE_SOURCE_DIGEST


def test_engine_digest_changes_when_an_engine_file_changes(tmp_path: Path):
    for name in ENGINE_SOURCE_FILES:
        (tmp_path / name).write_bytes((REPO_ROOT / name).read_bytes())
    (tmp_path / "paragraph_rules.py").write_bytes(
        (REPO_ROOT / "paragraph_rules.py").read_bytes() + b"\n# a behavioural change\n"
    )
    assert compute_engine_source_digest(tmp_path) != ENGINE_SOURCE_DIGEST


def test_engine_digest_is_exposed_through_the_facade_and_manifest_producer():
    from phase1_bundle import ProducerIdentity
    from spec_formatter import template_analysis

    assert template_analysis.ENGINE_SOURCE_DIGEST == ENGINE_SOURCE_DIGEST
    producer = ProducerIdentity(name="spec-template-normalizer", version="x", run_id="r")
    assert producer.to_dict()["engine_fingerprint"] == ENGINE_SOURCE_DIGEST
    assert engine_identity.ENGINE_SOURCE_DIGEST_LENGTH == len(ENGINE_SOURCE_DIGEST)
