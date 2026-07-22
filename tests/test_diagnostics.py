from __future__ import annotations

from datetime import datetime, timezone

import pytest

from spec_formatter import diagnostics as diag


def test_levels_round_trip_by_name_and_number() -> None:
    assert diag.level_from_name("debug") == diag.DEBUG
    assert diag.level_from_name("WARNING") == diag.WARNING
    assert diag.level_from_name("nonsense", default=diag.ERROR) == diag.ERROR
    assert diag.level_name(diag.INFO) == "INFO"
    assert diag.level_name(12345) == "INFO"


def test_sanitize_fields_keeps_scalars_and_identifier_strings() -> None:
    cleaned = diag.sanitize_fields(
        {
            "count": 12,
            "ratio": 3.5,
            "ok": True,
            "nothing": None,
            "mode": "format_only",
            "style_id": "CSI_Part__ARCH",
        }
    )
    assert cleaned == {
        "count": 12,
        "ratio": 3.5,
        "ok": True,
        "nothing": None,
        "mode": "format_only",
        "style_id": "CSI_Part__ARCH",
    }


def test_sanitize_fields_drops_prose_and_document_text_shaped_values() -> None:
    cleaned = diag.sanitize_fields(
        {
            "duration_ms": 4.2,
            "note": "this has spaces and is prose",  # spaces -> dropped
            "text": "single",  # doc-text-shaped KEY -> dropped even if value is safe
            "preview": "SECRET",  # doc-text-shaped key -> dropped
            "paragraph_xml": "<w:p/>",  # dropped
        }
    )
    assert cleaned == {"duration_ms": 4.2}


def test_sanitize_fields_recurses_and_bounds_depth() -> None:
    cleaned = diag.sanitize_fields(
        {
            # Nested scalars survive; a doc-text key is dropped; a string under
            # a non-whitelisted key ("b") is dropped even though it is clean.
            "nested": {"a": 1, "content": "drop-me", "b": "safe"},
            # Numbers survive in a list; strings do not, because the containing
            # key "items" is not whitelisted for string values.
            "items": [1, 2, "with space", "safe_token", {"x": 3}],
        }
    )
    assert cleaned["nested"] == {"a": 1}
    assert cleaned["items"] == [1, 2, {"x": 3}]


def test_sanitize_fields_keeps_strings_only_under_whitelisted_keys() -> None:
    cleaned = diag.sanitize_fields(
        {
            "mode": "format_only",  # whitelisted -> kept
            "model": "claude-sonnet-5",  # whitelisted -> kept
            "detail": "csi_to_canadian",  # clean token, non-whitelisted -> dropped
        }
    )
    assert cleaned == {"mode": "format_only", "model": "claude-sonnet-5"}


def test_sanitize_fields_rejects_secret_shaped_and_document_token_keys() -> None:
    api_key = "sk-ant-api03-" + "A1b2C3d4" * 12  # charset-clean, > _NAME limits
    cleaned = diag.sanitize_fields(
        {
            # Regression for the confirmed leak: an API key (or any document
            # token) used as a KEY must never survive, because JSON object keys
            # bypass value redaction downstream.
            "num_id_remaps": {api_key: 2, "CONFIDENTIAL": 4, "valid_child": 7},
            api_key: 1,
            "Uppercase": 3,  # keys must be lowercase identifiers
        }
    )
    assert cleaned == {"num_id_remaps": {"valid_child": 7}}


def test_sanitize_fields_blocks_paragraph_tokenisation_exfiltration() -> None:
    # A hostile processor tries to smuggle a confidential heading out one
    # whitespace-free token at a time, across arbitrary keys and a list.
    cleaned = diag.sanitize_fields(
        {
            "w0": "ACME",
            "w1": "MERGER",
            "w2": "PRICING",
            "tokens": ["ACME", "MERGER", "PRICING"],
            "modified": 12,
        }
    )
    # No token survives anywhere: the arbitrary word keys are gone and the list
    # collapses to empty, so nothing of the heading can be reconstructed.
    assert cleaned == {"modified": 12, "tokens": []}
    assert "ACME" not in str(cleaned)
    assert "MERGER" not in str(cleaned)


def test_recorder_respects_min_level() -> None:
    rec = diag.DiagnosticsRecorder(min_level=diag.WARNING)
    assert rec.info("pipeline", "quiet", count=1) is None
    assert rec.warning("pipeline", "loud", count=1) is not None
    events = rec.snapshot()
    assert [event.event for event in events] == ["loud"]


def test_recorder_assigns_monotonic_seq_and_utc_timestamps() -> None:
    ticks = iter(
        [
            datetime(2026, 7, 22, 12, 0, 0, tzinfo=timezone.utc),
            datetime(2026, 7, 22, 12, 0, 1, tzinfo=timezone.utc),
        ]
    )
    rec = diag.DiagnosticsRecorder(min_level=diag.DEBUG, clock=lambda: next(ticks))
    rec.info("pipeline", "first")
    rec.info("pipeline", "second")
    dumped = rec.iter_dicts()
    assert [entry["seq"] for entry in dumped] == [1, 2]
    assert dumped[0]["ts"] == "2026-07-22T12:00:00Z"
    assert dumped[1]["ts"] == "2026-07-22T12:00:01Z"


def test_timer_records_duration_and_error_type_without_message() -> None:
    rec = diag.DiagnosticsRecorder(min_level=diag.DEBUG)
    with rec.timer("target", "style_import") as phase:
        phase.set(requested_styles=4)
    with pytest.raises(RuntimeError):
        with rec.timer("target", "build_output"):
            raise RuntimeError("CONFIDENTIAL failure detail")
    events = rec.snapshot()
    ok = next(e for e in events if e.event == "style_import")
    assert ok.level == diag.INFO
    assert ok.fields["requested_styles"] == 4
    assert "duration_ms" in ok.fields
    bad = next(e for e in events if e.event == "build_output")
    assert bad.level == diag.ERROR
    assert bad.fields["failed"] is True
    assert bad.fields["error_type"] == "runtimeerror"
    dumped = "".join(str(e.as_dict()) for e in events)
    assert "CONFIDENTIAL" not in dumped


def test_ingest_resanitizes_untrusted_engine_events() -> None:
    rec = diag.DiagnosticsRecorder(min_level=diag.DEBUG)
    rec.ingest(
        [
            {
                "level": "info",
                "component": "target",
                "event": "apply_classifications",
                "fields": {"modified": 9, "leak": "SECRET DOC TEXT WITH SPACES"},
            }
        ],
        target=2,
    )
    (event,) = rec.snapshot()
    assert event.target == 2
    assert event.fields == {"modified": 9}


def test_engine_helpers_emit_and_time_sanitized_dicts() -> None:
    collector: list = []
    diag.emit(collector, "info", "target", "slim_bundle", paragraphs=10, junk="a b c")
    with diag.timed(collector, "target", "extract") as phase:
        phase.set(bytes_read=4096)
    assert collector[0]["fields"] == {"paragraphs": 10}
    assert collector[1]["event"] == "extract"
    assert "duration_ms" in collector[1]["fields"]
    assert collector[1]["fields"]["bytes_read"] == 4096


def test_summary_rolls_up_levels_and_phase_durations() -> None:
    rec = diag.DiagnosticsRecorder(min_level=diag.DEBUG)
    rec.info("pipeline", "config_load", duration_ms=2.0)
    rec.info("target", "extract", target=1, duration_ms=5.0)
    rec.info("target", "extract", target=2, duration_ms=7.0)
    rec.warning("target", "coverage", target=1, classification_pct=80.0)
    summary = rec.summary()
    assert summary["event_count"] == 4
    assert summary["counts_by_level"] == {"INFO": 3, "WARNING": 1}
    assert summary["warnings"] == 1
    assert summary["errors"] == 0
    assert summary["targets_observed"] == 2
    assert summary["phase_durations_ms"]["target.extract"] == 12.0
    assert summary["phase_counts"]["target.extract"] == 2
    assert summary["log"] == "diagnostics.jsonl"


def test_sanitize_event_cleans_component_event_and_fields() -> None:
    event = diag.sanitize_event(
        {
            "level": "warning",
            "component": "Bad Component",  # invalid name -> "invalid"
            "event": "coverage",
            "target": 3,
            "fields": {"pct": 90.0, "secret": "leak me now"},
        }
    )
    assert event == {
        "level": "WARNING",
        "component": "invalid",
        "event": "coverage",
        "fields": {"pct": 90.0},
        "target": 3,
    }
