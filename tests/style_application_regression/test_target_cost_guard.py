"""Target cost preflight must refuse before constructing a model client."""

from types import SimpleNamespace

import anthropic
import pytest

from spec_formatter.llm_usage import UsageCollector
from spec_formatter.style_application import batch_runner
from spec_formatter.style_application.core.errors import ERROR_REMEDIATIONS


ENV_NAME = "SPEC_FORMATTER_MAX_TARGET_PARAGRAPHS"


@pytest.fixture
def target(monkeypatch, tmp_path):
    source = tmp_path / "source.docx"
    source.write_bytes(b"unchanged source")
    monkeypatch.delenv(ENV_NAME, raising=False)
    monkeypatch.setattr(batch_runner, "DocxDecomposer", lambda _path: SimpleNamespace(
        extract=lambda **_kwargs: tmp_path,
    ))

    def unexpected_client(*_args, **_kwargs):
        pytest.fail("Target preflight constructed an Anthropic client")

    monkeypatch.setattr(anthropic, "Anthropic", unexpected_client)
    return source


def _bundle(unresolved):
    # The free workload already exceeds the default cap. Only unresolved
    # paragraphs should count, regardless of how much local work was done.
    free_count = batch_runner.DEFAULT_MAX_TARGET_PARAGRAPHS + 1
    return {
        "paragraphs": [
            {"paragraph_index": i, "text": "Synthetic requirement."}
            for i in range(unresolved)
        ],
        "deterministic_classifications": [
            {"paragraph_index": i, "csi_role": "PARAGRAPH"}
            for i in range(unresolved, unresolved + free_count)
        ],
        "deterministic_ignored_paragraphs": [
            {"paragraph_index": unresolved + free_count, "reason": "boilerplate"},
        ],
        "filter_report": {"paragraphs_out_of_scope": [
            {"paragraph_index": i, "reason": "table"}
            for i in range(unresolved + free_count + 1, unresolved + 2 * free_count + 1)
        ]},
    }


def _process(source, *, api_key="offline-test-key"):
    return batch_runner.process_single_file(
        source, {"PARAGRAPH": "Body"}, {}, "<w:styles/>", ["PARAGRAPH"],
        api_key, source.parent / "output",
    )


def _assert_refused(result):
    assert result.success is False
    assert result.output_path is None
    assert result.stage == "classification_preflight"
    assert result.error_code == "target_too_large"
    assert result.safe_error == ERROR_REMEDIATIONS["target_too_large"]
    assert result.usage == UsageCollector().snapshot()
    assert result.usage["requests_attempted"] == 0
    assert result.usage["usage_complete"] is True
    assert not any(event["event"] == "classify" for event in result.diagnostics)


@pytest.mark.parametrize("override,limit", [(None, 2_000), ("3", 3), ("2500", 2_500)])
@pytest.mark.parametrize("offset", [-1, 0, 1])
def test_target_guard_boundary_and_positive_overrides(target, monkeypatch, override, limit, offset):
    if override is not None:
        monkeypatch.setenv(ENV_NAME, override)
    bundle = _bundle(limit + offset)
    monkeypatch.setattr(batch_runner, "build_phase2_slim_bundle", lambda *_a, **_k: bundle)
    classifications = []
    applications = []

    def classify(**kwargs):
        classifications.append(kwargs["slim_bundle"])
        return {"classifications": [], "ignored_paragraphs": []}

    def apply(**kwargs):
        applications.append(kwargs)
        return target.parent / "out.docx", None, {}, {}, {}

    # Refused cases retain the real classifier; the fixture's client trap
    # proves the cost guard prevents any SDK construction (and hence calls).
    if offset <= 0:
        monkeypatch.setattr(batch_runner, "classify_target_document", classify)
    monkeypatch.setattr(batch_runner, "_apply_classified_target", apply)

    result = _process(target)

    if offset > 0:
        _assert_refused(result)
        assert not applications
    else:
        assert result.success is True
        assert result.stage == "complete"
        assert classifications == [bundle]
        assert len(applications) == 1
    assert target.read_bytes() == b"unchanged source"


@pytest.mark.parametrize("override", ["", "invalid", "0", "-1", "2.5"])
def test_invalid_override_keeps_default_guard(target, monkeypatch, override):
    monkeypatch.setenv(ENV_NAME, override)
    bundle = _bundle(batch_runner.DEFAULT_MAX_TARGET_PARAGRAPHS + 1)
    monkeypatch.setattr(batch_runner, "build_phase2_slim_bundle", lambda *_a, **_k: bundle)

    _assert_refused(_process(target))


def test_oversized_target_fails_before_api_key_check(target, monkeypatch):
    monkeypatch.setenv(ENV_NAME, "3")
    monkeypatch.setattr(batch_runner, "build_phase2_slim_bundle", lambda *_a, **_k: _bundle(4))

    _assert_refused(_process(target, api_key=""))


def test_large_deterministic_target_is_unaffected(target, monkeypatch):
    monkeypatch.setenv(ENV_NAME, "1")
    bundle = _bundle(0)
    monkeypatch.setattr(batch_runner, "build_phase2_slim_bundle", lambda *_a, **_k: bundle)
    applications = []

    def apply(**kwargs):
        applications.append(kwargs["classifications"])
        return target.parent / "out.docx", None, {}, {}, {}

    monkeypatch.setattr(batch_runner, "_apply_classified_target", apply)

    result = _process(target, api_key="")

    assert result.success is True
    assert result.usage == UsageCollector().snapshot()
    assert applications[0]["classifications"] == bundle["deterministic_classifications"]
    assert applications[0]["ignored_paragraphs"] == bundle["deterministic_ignored_paragraphs"]
