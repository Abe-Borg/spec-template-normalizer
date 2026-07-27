from __future__ import annotations

import queue
from dataclasses import FrozenInstanceError
from datetime import datetime, timezone
from pathlib import Path
from types import SimpleNamespace

import gui
import pytest
from spec_formatter.pipeline import CSI_TO_CANADIAN, FORMAT_ONLY
from spec_formatter.style_application.core.csi_to_canadian import (
    CanadianConversionReport,
    ConversionIssue,
)


def _worker(events: queue.Queue, mode: str) -> gui.FormatWorker:
    return gui.FormatWorker(
        architect_template=Path("architect.docx"),
        target_inputs=(Path("target.docx"),),
        output_dir=Path("formatted"),
        api_key="test-key",
        reuse_template_analysis=True,
        max_workers=2,
        conversion_mode=mode,
        events=events,
    )


def test_format_worker_forwards_each_output_mode(monkeypatch):
    received = []
    sentinel = object()

    def fake_format_specifications(**kwargs):
        received.append(kwargs)
        return sentinel

    monkeypatch.setattr(gui, "format_specifications", fake_format_specifications)
    monkeypatch.setattr(gui, "default_template_cache_dir", lambda: Path("cache"))

    for mode in (FORMAT_ONLY, CSI_TO_CANADIAN):
        events: queue.Queue = queue.Queue()
        _worker(events, mode).run()
        assert events.get_nowait() == ("complete", sentinel)

    assert [item["conversion_mode"] for item in received] == [
        FORMAT_ONLY,
        CSI_TO_CANADIAN,
    ]


def test_format_worker_reports_pipeline_errors(monkeypatch):
    def fail(**_kwargs):
        raise ValueError("invalid Canadian template")

    monkeypatch.setattr(gui, "format_specifications", fail)
    monkeypatch.setattr(gui, "default_template_cache_dir", lambda: Path("cache"))
    events: queue.Queue = queue.Queue()

    _worker(events, CSI_TO_CANADIAN).run()

    kind, payload = events.get_nowait()
    assert kind == "error"
    assert payload["message"].startswith("[untrusted detail omitted; sha256=")
    assert "invalid Canadian template" not in payload["message"]
    assert payload["error_code"] == "untrusted_error"
    assert "ValueError" in payload["traceback"]
    assert payload["run_dir"] is None


def test_format_worker_reports_input_failure_without_exposing_path(monkeypatch):
    private_path = r"C:\Private\Client A\secret-template.docx"

    def fail(**_kwargs):
        raise FileNotFoundError(
            f"Architect template does not exist: {private_path}"
        )

    monkeypatch.setattr(gui, "format_specifications", fail)
    monkeypatch.setattr(gui, "default_template_cache_dir", lambda: Path("cache"))
    events: queue.Queue = queue.Queue()

    _worker(events, FORMAT_ONLY).run()

    kind, payload = events.get_nowait()
    assert kind == "error"
    assert payload["error_code"] == "input_architect_missing"
    assert payload["message"] == "Architect template does not exist."
    assert private_path not in payload["message"]
    assert private_path not in payload["traceback"]


def test_format_worker_preserves_pipeline_progress_event_timestamp(monkeypatch):
    occurred_at = datetime(2026, 7, 21, 21, 35, 24, tzinfo=timezone.utc)
    sentinel = object()

    def fake_format_specifications(**kwargs):
        kwargs["progress_event"]("Processing target", occurred_at)
        return sentinel

    monkeypatch.setattr(gui, "format_specifications", fake_format_specifications)
    monkeypatch.setattr(gui, "default_template_cache_dir", lambda: Path("cache"))
    events: queue.Queue = queue.Queue()

    _worker(events, FORMAT_ONLY).run()

    assert events.get_nowait() == (
        "progress",
        {"message": "Processing target", "occurred_at": occurred_at},
    )
    assert events.get_nowait() == ("complete", sentinel)


def test_format_worker_reports_wrapped_known_internal_error_actionably(monkeypatch):
    canonical = (
        "Architect template has conflicting section shells; use one canonical "
        "page layout and default/even/first header-footer mapping."
    )

    def fail(**_kwargs):
        raise ValueError(
            "Preflight validation failed (page_layout semantic validation failed: "
            f"{canonical}) PRIVATE DETAIL"
        )

    monkeypatch.setattr(gui, "format_specifications", fail)
    monkeypatch.setattr(gui, "default_template_cache_dir", lambda: Path("cache"))
    events: queue.Queue = queue.Queue()

    _worker(events, CSI_TO_CANADIAN).run()

    kind, payload = events.get_nowait()
    assert kind == "error"
    assert payload["error_code"] == "template_section_shell_conflict"
    assert payload["message"] == canonical
    assert "PRIVATE DETAIL" not in payload["traceback"]


def test_format_worker_forwards_failed_run_artifact_paths(monkeypatch):
    run_dir = Path("formatted/20260721_format_only_deadbeef")
    manifest = run_dir / "run.json"

    def fail(**_kwargs):
        error = RuntimeError("template analysis failed")
        error.run_dir = run_dir
        error.manifest_path = manifest
        raise error

    monkeypatch.setattr(gui, "format_specifications", fail)
    monkeypatch.setattr(gui, "default_template_cache_dir", lambda: Path("cache"))
    events: queue.Queue = queue.Queue()

    _worker(events, FORMAT_ONLY).run()

    kind, payload = events.get_nowait()
    assert kind == "error"
    assert payload["run_dir"] == run_dir
    assert payload["manifest_path"] == manifest


def test_failed_target_conversion_report_still_reaches_activity_log():
    report = CanadianConversionReport(
        paragraphs_examined=2,
        paragraphs_converted=2,
        literal_markers_removed=1,
        automatic_numbering_retargeted=1,
        unnumbered_paragraphs_numbered=0,
        edits=(),
        warnings=(
            ConversionIssue(4, "example", "Review this converted paragraph.", "Scope"),
        ),
    )
    item = SimpleNamespace(
        success=False,
        source_path=Path("mechanical.docx"),
        conversion_report=report,
    )

    lines = gui.conversion_report_log_lines(item)

    assert "processed 2 numbered paragraphs" in lines[0]
    assert "removed 1 typed markers" in lines[0]
    assert "warning p[4]" in lines[1]


def test_active_run_summary_is_frozen_and_contains_no_api_key():
    summary = gui.ActiveRunSummary(
        architect_template=Path("architect.docx"),
        target_inputs=(Path("one.docx"), Path("two.docx")),
        output_root=Path("formatted"),
        conversion_mode=FORMAT_ONLY,
        reuse_template_analysis=True,
        max_workers=2,
    )

    text = gui.active_run_summary_text(summary)

    assert "ACTIVE RUN (inputs locked)" in text
    assert "Format only" in text
    assert "2 target selection(s)" in text
    assert "architect.docx" in text
    assert "api" not in text.lower()
    with pytest.raises(FrozenInstanceError):
        summary.max_workers = 4  # type: ignore[misc]


def test_target_result_log_lines_include_every_line_and_audit_details():
    item = SimpleNamespace(
        source_path=Path("payment.docx"),
        log=("first", "second\nthird"),
        audit_summary={
            "styled": 120,
            "ignored": 2,
            "out_of_scope": 3,
            "unresolved": 0,
        },
        audit_path=Path("run/payment.audit.json"),
        conversion_report=None,
    )

    lines = gui.target_result_log_lines(item)

    assert lines[:3] == ("first", "second", "third")
    assert lines[3] == (
        "payment.docx: audit counts: styled=120, ignored=2, "
        "out_of_scope=3, unresolved=0"
    )
    assert lines[4] == f"payment.docx: audit: {Path('run/payment.audit.json')}"


def test_result_run_directory_prefers_concrete_run_dir_and_falls_back():
    current = SimpleNamespace(
        run_dir=Path("output/20260721_format_only_abcd1234"),
        output_dir=Path("legacy-output"),
        output_root=Path("output"),
    )
    legacy = SimpleNamespace(output_dir=Path("legacy-output"))

    assert gui.result_run_directory(current) == current.run_dir
    assert gui.result_run_directory(legacy) == legacy.output_dir
    assert gui.result_run_directory(SimpleNamespace()) is None


class _FakeControl:
    def __init__(self, state: str) -> None:
        self.state = state

    def cget(self, option: str) -> str:
        assert option == "state"
        return self.state

    def configure(self, **kwargs) -> None:
        self.state = kwargs["state"]


def test_run_controls_are_all_locked_and_exact_states_are_restored():
    normal = _FakeControl("normal")
    already_disabled = _FakeControl("disabled")
    app = SimpleNamespace(
        run_affecting_controls=[normal, already_disabled],
        _locked_run_control_states=[],
    )

    gui.App._lock_run_controls(app)

    assert normal.state == "disabled"
    assert already_disabled.state == "disabled"

    gui.App._unlock_run_controls(app)

    assert normal.state == "normal"
    assert already_disabled.state == "disabled"
    assert app._locked_run_control_states == []


class _FakeWidget:
    def __init__(self) -> None:
        self.configurations: list[dict] = []

    def configure(self, **kwargs) -> None:
        self.configurations.append(kwargs)


class _FakeLogBox(_FakeWidget):
    def __init__(self) -> None:
        super().__init__()
        self.insertions: list[tuple[str, str]] = []

    def insert(self, index: str, value: str) -> None:
        self.insertions.append((index, value))

    def see(self, _index: str) -> None:
        return None


def test_append_log_converts_utc_progress_time_to_local_time():
    occurred_at = datetime(2026, 7, 21, 21, 35, 24, tzinfo=timezone.utc)
    log_box = _FakeLogBox()
    app = SimpleNamespace(log_box=log_box)

    gui.App._append_log(app, "Processing target", occurred_at=occurred_at)

    expected_time = occurred_at.astimezone().strftime("%H:%M:%S")
    assert log_box.insertions == [
        ("end", f"[{expected_time}] Processing target\n")
    ]


def test_completion_does_not_replay_processor_details_and_saves_summaries(monkeypatch):
    run_dir = Path("output/20260721_format_only_abcd1234")
    target = SimpleNamespace(
        success=True,
        source_path=Path("payment.docx"),
        output_path=run_dir / "payment_FORMATTED.docx",
        error=None,
        log=("processor detail one", "processor detail two"),
        audit_summary={"styled": 121, "ignored": 1},
        audit_path=run_dir / "payment.audit.json",
        conversion_report=None,
    )
    result = SimpleNamespace(
        targets=(target,),
        run_id="abcd1234",
        run_dir=run_dir,
        output_dir=Path("legacy-output"),
        output_root=Path("output"),
        manifest_path=run_dir / "run.json",
    )
    messages: list[str] = []
    app = SimpleNamespace(
        _finish_busy_state=lambda: None,
        last_result=None,
        active_output_dir=Path("output"),
        status_label=_FakeWidget(),
        open_button=_FakeWidget(),
        _append_log=messages.append,
    )
    monkeypatch.setattr(gui.messagebox, "showinfo", lambda *_args, **_kwargs: None)

    gui.App._handle_complete(app, result)

    assert "processor detail one" not in messages
    assert "processor detail two" not in messages
    assert any("styled=121, ignored=1" in line for line in messages)
    assert f"Output: {target.output_path}" in messages
    assert f"Run folder: {run_dir}" in messages
    assert f"Persisted run log: {run_dir / 'run.log'}" in messages
    assert app.last_result is result


def test_poll_events_uses_captured_progress_timestamp_and_accepts_legacy_string():
    occurred_at = datetime(2026, 7, 21, 14, 35, 24)
    events: queue.Queue = queue.Queue()
    events.put(
        (
            "progress",
            {"message": "Processing target", "occurred_at": occurred_at},
        )
    )
    events.put(("progress", "legacy progress"))
    logged: list[tuple[str, datetime | None]] = []
    app = SimpleNamespace(
        events=events,
        status_label=_FakeWidget(),
        _append_log=lambda message, occurred_at=None: logged.append(
            (message, occurred_at)
        ),
        _handle_complete=lambda _payload: None,
        _handle_error=lambda _payload: None,
        _poll_events=lambda: None,
        after=lambda *_args: None,
    )

    gui.App._poll_events(app)

    assert logged == [
        ("Processing target", occurred_at),
        ("legacy progress", None),
    ]
