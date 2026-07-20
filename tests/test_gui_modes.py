from __future__ import annotations

import queue
from pathlib import Path
from types import SimpleNamespace

import gui
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
    assert payload["message"] == "invalid Canadian template"
    assert "ValueError" in payload["traceback"]


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
