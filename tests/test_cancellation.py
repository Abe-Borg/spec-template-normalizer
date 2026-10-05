"""Cancellation stops paid work and preserves truthful, complete run artifacts."""

from __future__ import annotations

import hashlib
import json
import threading
from pathlib import Path
from types import SimpleNamespace

import anthropic
import httpx
import pytest

import llm_classifier as architect_classifier
from spec_formatter import pipeline
from spec_formatter.cancellation import RunCancelled, wait_for_retry
from spec_formatter.llm_usage import UsageCollector, usage_from_exception
from spec_formatter.style_application import batch_runner
from spec_formatter.style_application.core import llm_classifier as target_classifier
from spec_formatter.style_application.core.errors import ERROR_REMEDIATIONS, RUN_STATUSES
from tests.test_architect_free_modes import _document_xml, _write_docx


def _run(tmp_path, targets, event, **kwargs):
    return pipeline.format_specifications(
        None, targets, tmp_path / "output", "offline-key",
        conversion_mode="csi_to_canadian_standalone",
        cancel_event=event,
        **kwargs,
    )


def _assert_artifacts(result, statuses):
    assert result.cancelled and not result.success
    manifest = json.loads(result.manifest_path.read_text())
    assert manifest["schema_version"] == 4
    assert manifest["status"] == "cancelled"
    assert manifest["status"] in RUN_STATUSES
    assert len(manifest["targets"]) == len(statuses) == len(result.targets)
    assert (result.run_dir / "run.log").is_file()
    assert result.diagnostics_path.is_file()
    for line in result.diagnostics_path.read_text().splitlines():
        json.loads(line)
    assert not (result.run_dir / ".staging").exists()
    for item, record, expected in zip(result.targets, manifest["targets"], statuses):
        audit = json.loads(item.audit_path.read_text())
        assert audit["schema_version"] == 4
        assert record["success"] == audit["success"] == item.success == (expected == "succeeded")
        assert record["usage"] == audit["usage"] == item.usage
        assert record["audit_path"] == str(item.audit_path)
        if expected == "cancelled":
            assert record["stage"] == audit["stage"] == item.stage == "cancelled"
            assert record["error_code"] == audit["error_code"] == "run_cancelled"
            assert audit["error"] == ERROR_REMEDIATIONS["run_cancelled"]
            assert audit["output"] is None
            assert record["output_path"] is None and item.output_path is None
        else:
            assert item.output_path.is_file()
            assert record["output_sha256"] == hashlib.sha256(item.output_path.read_bytes()).hexdigest()
    assert len(list(result.run_dir.glob("*.docx"))) == statuses.count("succeeded")
    return manifest


@pytest.mark.parametrize("mode", ["format_only", "csi_to_canadian", "csi_to_canadian_standalone", "canadian_to_csi"])
def test_cancel_before_start_writes_every_audit_without_starting_work(tmp_path, mode):
    event = threading.Event()
    event.set()
    architect = _write_docx(tmp_path / "architect.docx") if mode in ("format_only", "csi_to_canadian") else None
    targets = [_write_docx(tmp_path / f"target-{i}.docx") for i in range(3)]

    def never_called(**_kwargs):
        pytest.fail("Cancelled run started work")

    result = pipeline.format_specifications(
        architect, targets, tmp_path / "output", "",
        conversion_mode=mode, cancel_event=event,
        _template_analyzer=never_called, _target_processor=never_called,
    )
    manifest = _assert_artifacts(result, ["cancelled"] * 3)
    usage = manifest["diagnostics"]["usage"]["target"]
    assert usage["requests_attempted"] == 0 and usage["usage_complete"]


def test_cancel_between_targets_keeps_published_docx_and_does_not_start_next(tmp_path):
    event = threading.Event()
    targets = [_write_docx(tmp_path / f"target-{i}.docx") for i in range(3)]
    originals = [p.read_bytes() for p in targets]
    started = []

    def progress(message):
        if message.startswith("Processing target"):
            started.append(message)
        if message.startswith("Formatted 1 of"):
            event.set()

    result = _run(tmp_path, targets, event, max_workers=1, progress=progress)
    _assert_artifacts(result, ["succeeded", "cancelled", "cancelled"])
    assert len(started) == 1
    assert [p.read_bytes() for p in targets] == originals


class _Stream:
    def __init__(self, *, text="{}", stop_reason="end_turn", on_event=None, on_final=None, iterable=True):
        self.text = text
        self.stop_reason = stop_reason
        self.on_event = on_event
        self.on_final = on_final
        self.iterable = iterable
        self.closed = False

    def __enter__(self):
        return self

    def __exit__(self, *_args):
        self.closed = True

    def __iter__(self):
        if not self.iterable:
            raise TypeError("final-message-only stream")
        return self._events()

    def _events(self):
        if self.on_event:
            self.on_event()
        yield SimpleNamespace(type="message_start")

    def get_final_message(self):
        if self.on_final:
            self.on_final()
        return SimpleNamespace(
            stop_reason=self.stop_reason,
            usage=SimpleNamespace(input_tokens=20, output_tokens=3),
        )

    def get_final_text(self):
        return self.text


def _install_client(monkeypatch, stream_factory):
    requests = []

    def stream(**kwargs):
        requests.append(kwargs)
        return stream_factory(len(requests), kwargs)

    client = SimpleNamespace(messages=SimpleNamespace(stream=stream))
    monkeypatch.setattr(anthropic, "Anthropic", lambda **_kwargs: client)
    return client, requests


@pytest.mark.parametrize("diagnostics_level", ["info", "error"])
def test_cancel_during_chunked_classification_stops_queued_requests_and_records_unknown_usage(tmp_path, monkeypatch, diagnostics_level):
    event = threading.Event()
    paragraphs = "".join(
        f"<w:p><w:r><w:t>Requirement without a list marker number {i}.</w:t></w:r></w:p>"
        for i in range(12)
    )
    document = _document_xml().replace("<w:sectPr>", paragraphs + "<w:sectPr>")
    target = _write_docx(tmp_path / "target.docx", document=document)
    original = target.read_bytes()
    streams = []

    def split(bundle):
        assert len(bundle["paragraphs"]) > 6
        return [{"paragraphs": [p]} for p in bundle["paragraphs"]]

    def make_stream(number, kwargs):
        assert not event.is_set(), "Sent another paid request after cancellation"
        # The first completed response is accounted for; cancellation closes
        # the second before final usage. Other chunks are waiting for a slot.
        if number == 1:
            content = kwargs["messages"][0]["content"]
            index = json.loads(content.split("\n")[-1])["paragraphs"][0]["paragraph_index"]
            text = json.dumps({"classifications": [], "ignored_paragraphs": [{"paragraph_index": index, "reason": "non_csi_content"}]})
            stream = _Stream(text=text)
        else:
            assert number == 2
            stream = _Stream(on_event=event.set)
        streams.append(stream)
        return stream

    monkeypatch.setattr(target_classifier, "_split_bundle_into_chunks", split)
    monkeypatch.setattr(target_classifier, "_REQUEST_LIMITER", threading.BoundedSemaphore(1))
    _client, requests = _install_client(monkeypatch, make_stream)
    result = _run(tmp_path, [target], event, diagnostics_level=diagnostics_level)
    manifest = _assert_artifacts(result, ["cancelled"])
    assert len(requests) == 2 and all(s.closed for s in streams)
    usage = result.targets[0].usage
    assert usage["requests_attempted"] == 2
    assert usage["responses_completed"] == 1
    assert usage["requests_with_unknown_usage"] == 1
    assert usage["usage_complete"] is False
    assert usage["input_tokens"] == 20 and usage["output_tokens"] == 3
    assert manifest["diagnostics"]["usage"]["target"] == usage
    assert result.targets[0].audit_summary["unresolved"] >= 12
    assert target.read_bytes() == original


def test_cancel_during_architect_stream_writes_cancelled_run_artifacts(tmp_path, monkeypatch):
    event = threading.Event()
    architect = _write_docx(tmp_path / "architect.docx")
    target = _write_docx(tmp_path / "target.docx")
    stream = _Stream(on_event=event.set)
    _client, requests = _install_client(monkeypatch, lambda *_args: stream)
    result = pipeline.format_specifications(
        architect, [target], tmp_path / "output", "offline-key",
        cache_dir=tmp_path / "cache", force_template_analysis=True,
        cancel_event=event,
    )
    manifest = _assert_artifacts(result, ["cancelled"])
    assert len(requests) == 1 and stream.closed
    usage = manifest["diagnostics"]["usage"]["architect"]
    assert usage["requests_attempted"] == usage["requests_with_unknown_usage"] == 1
    assert not usage["usage_complete"] and "input_tokens" not in usage
    assert not list((tmp_path / "cache").rglob(".phase1-work-*"))


def test_cancel_after_finished_stream_preserves_known_usage(tmp_path, monkeypatch):
    event = threading.Event()
    stream = _Stream(on_final=event.set, iterable=False)
    _client, requests = _install_client(monkeypatch, lambda *_args: stream)
    with pytest.raises(RunCancelled) as raised:
        target_classifier.classify_target_document(
            {"paragraphs": [{"paragraph_index": 0, "text": "unresolved"}]},
            ["BODY"], "offline-key", cancel_event=event,
        )
    usage = usage_from_exception(raised.value)
    assert len(requests) == 1 and stream.closed
    assert usage["requests_attempted"] == usage["responses_completed"] == 1
    assert usage["usage_complete"] and usage["requests_with_unknown_usage"] == 0
    assert usage["input_tokens"] == 20


@pytest.mark.parametrize("classifier", ["architect", "target"])
@pytest.mark.parametrize("failure", ["invalid_json", "max_tokens", "transport"])
def test_cancel_during_backoff_prevents_regeneration_and_transport_retry(monkeypatch, classifier, failure):
    event = threading.Event()
    waits = []

    def wait(seconds):
        waits.append(seconds)
        event.set()
        return True

    monkeypatch.setattr(event, "wait", wait)

    def make_stream(*_args):
        if failure == "transport":
            raise anthropic.RateLimitError(
                "retry", response=httpx.Response(429, request=httpx.Request("POST", "https://api.anthropic.com"), headers={"retry-after": "120"}), body={},
            )
        return _Stream(text="bad json", stop_reason="max_tokens" if failure == "max_tokens" else "end_turn")

    client, requests = _install_client(monkeypatch, make_stream)
    usage = UsageCollector()
    with pytest.raises(RunCancelled):
        if classifier == "target":
            target_classifier.classify_target_document(
                {"paragraphs": [{"paragraph_index": 0}]}, ["BODY"], "offline-key", cancel_event=event,
            )
        elif failure == "transport":
            architect_classifier._call_api(client, "system", "user", "model", usage=usage, cancel_event=event)
        else:
            # Architect regeneration has no backoff. Set cancellation when
            # the first invalid response is parsed instead.
            original = architect_classifier._parse_response

            def parse(raw):
                event.set()
                return original(raw)

            if failure == "max_tokens":
                original = architect_classifier._call_api

                def call(*args, **kwargs):
                    try:
                        return original(*args, **kwargs)
                    finally:
                        event.set()

                monkeypatch.setattr(architect_classifier, "_call_api", call)
            else:
                monkeypatch.setattr(architect_classifier, "_parse_response", parse)
            architect_classifier._request_json_response(
                client, "system", "user", "model", response_schema={}, max_attempts=2,
                usage=usage, cancel_event=event,
            )
    assert len(requests) == 1
    if classifier == "target" and failure == "transport":
        assert waits == [120]


def test_real_backoff_wait_is_interrupted_promptly():
    event = threading.Event()
    waiting = threading.Event()
    outcome = []

    def backoff():
        waiting.set()
        try:
            wait_for_retry(event, 120)
        except RunCancelled:
            outcome.append("cancelled")

    worker = threading.Thread(target=backoff)
    worker.start()
    assert waiting.wait(2)
    event.set()
    worker.join(2)
    assert not worker.is_alive() and outcome == ["cancelled"]


def test_old_injected_processor_is_called_once_without_new_keyword(tmp_path):
    event = threading.Event()
    target = _write_docx(tmp_path / "target.docx")
    calls = []

    def old_processor(docx_path, arch_registry, env_registry, arch_styles_xml, available_roles, api_key, output_dir, source_tokens, arch_root, model, role_specs, conversion_mode):
        calls.append(docx_path)
        event.set()
        raise TypeError("TypeError inside the processor")

    result = _run(tmp_path, [target], event, _target_processor=old_processor)
    assert result.cancelled and len(calls) == 1
    assert result.targets[0].error_code == "untrusted_error"
    assert "inside the processor" in result.targets[0].error
    assert result.manifest_path.is_file()


def test_cancellation_withholds_success_returned_by_old_processor(tmp_path):
    event = threading.Event()
    target = _write_docx(tmp_path / "target.docx")

    def processor(**kwargs):
        output = _write_docx(kwargs["output_dir"] / "output.docx")
        event.set()
        return batch_runner.BatchResult("target.docx", True, output, [], None, 0.1, usage=UsageCollector().snapshot())

    result = _run(tmp_path, [target], event, _target_processor=processor)
    _assert_artifacts(result, ["cancelled"])


def test_run_status_set_and_cancelled_stages_are_documented():
    from spec_formatter.style_application.core.errors import PIPELINE_STAGES, RUNNER_STAGES
    guide = (Path(pipeline.__file__).resolve().parents[1] / "CLAUDE.md").read_text()
    assert set(RUN_STATUSES) == {"succeeded", "partial_failure", "failed", "cancelled"}
    assert "cancelled" in PIPELINE_STAGES and "cancelled" in RUNNER_STAGES
    assert "`run_cancelled`" in guide


def test_cancel_before_architect_coverage_patch_sends_no_followup(monkeypatch):
    import docx_decomposer

    event = threading.Event()
    stream = _Stream(text=json.dumps({"roles": [], "create_styles": [], "apply_pStyle": [], "ignored_paragraphs": [], "notes": []}))
    _client, requests = _install_client(monkeypatch, lambda *_args: stream)

    def validate(*_args, **_kwargs):
        event.set()
        raise ValueError("classification coverage mismatch; missing=[0], unexpected=[]")

    monkeypatch.setattr(docx_decomposer, "validate_instructions", validate)
    with pytest.raises(RunCancelled):
        architect_classifier.classify_document(
            {"paragraphs": [{"paragraph_index": 0, "text": "ambiguous prose"}]},
            "system", "user", "offline-key", cancel_event=event,
        )
    assert len(requests) == 1


def test_cancel_before_overlap_reask_sends_no_followup(monkeypatch):
    event = threading.Event()
    chunks = [{"paragraphs": [{"paragraph_index": 0}]}] * 2
    text = json.dumps({"classifications": [{"paragraph_index": 0, "csi_role": "BODY"}], "ignored_paragraphs": []})
    _client, requests = _install_client(monkeypatch, lambda *_args: _Stream(text=text))
    monkeypatch.setattr(target_classifier, "_split_bundle_into_chunks", lambda _bundle: chunks)

    def conflicts(_results):
        event.set()
        return [{"paragraph_index": 0}]

    monkeypatch.setattr(target_classifier, "_chunk_conflicts", conflicts)
    with pytest.raises(RunCancelled) as raised:
        target_classifier.classify_target_document(
            {"paragraphs": [{"paragraph_index": 0}]}, ["BODY"], "offline-key", cancel_event=event,
        )
    assert len(requests) == 2
    assert usage_from_exception(raised.value)["requests_attempted"] == 2


def test_cancel_before_structured_output_fallback_sends_no_followup(monkeypatch):
    event = threading.Event()

    def fail(*_args):
        event.set()
        raise anthropic.BadRequestError(
            "grammar compilation timed out",
            response=httpx.Response(400, request=httpx.Request("POST", "https://api.anthropic.com")),
            body={"error": {"message": "grammar compilation timed out"}},
        )

    client, requests = _install_client(monkeypatch, fail)
    usage = UsageCollector()
    with pytest.raises(RunCancelled):
        architect_classifier._call_api(client, "system", "user", "model", cancel_event=event, usage=usage)
    assert len(requests) == 1
    assert usage.snapshot()["requests_with_unknown_usage"] == 1


def test_cancel_during_final_copy_removes_partial_docx(tmp_path, monkeypatch):
    event = threading.Event()
    source = _write_docx(tmp_path / "source.docx")
    final = tmp_path / "output" / "final.docx"
    original_fsync = pipeline.os.fsync

    def fsync(fd):
        original_fsync(fd)
        event.set()

    monkeypatch.setattr(pipeline.os, "fsync", fsync)
    with pytest.raises(RunCancelled):
        pipeline._publish_output(source, final, cancel_event=event)
    assert not final.exists()
    assert not list((final.parent / ".staging").iterdir())


def test_injected_usage_cannot_publish_document_text(tmp_path):
    event = threading.Event()
    target = _write_docx(tmp_path / "target.docx")

    def processor(**_kwargs):
        event.set()
        return batch_runner.BatchResult(
            "target.docx", False, None, [], "cancelled", 0.1,
            stage="cancelled", error_code="run_cancelled",
            usage={"requests_attempted": 1, "requests_with_unknown_usage": 1,
                   "usage_complete": False, "document_text": "PRIVATE CONTENT",
                   "model": "PRIVATE CONTENT"},
        )

    result = _run(tmp_path, [target], event, _target_processor=processor)
    _assert_artifacts(result, ["cancelled"])
    assert "document_text" not in result.targets[0].usage
    for path in result.run_dir.iterdir():
        if path.is_file():
            assert b"PRIVATE CONTENT" not in path.read_bytes()
