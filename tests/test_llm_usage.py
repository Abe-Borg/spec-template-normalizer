"""The shared observed-usage contract, including work that failed."""

from __future__ import annotations

import threading
import types

import pytest

from spec_formatter.llm_usage import (
    UsageCollector,
    attach_usage,
    usage_from_exception,
    usage_numbers,
)


def _message(**fields):
    return types.SimpleNamespace(stop_reason="end_turn", usage=types.SimpleNamespace(**fields))


def test_recognized_counters_are_read_and_summed():
    collector = UsageCollector()
    collector.record_attempt()
    collector.record_response(_message(input_tokens=10, output_tokens=2))
    collector.record_attempt()
    collector.record_response(_message(input_tokens=5, output_tokens=1))
    snapshot = collector.snapshot()
    assert snapshot["input_tokens"] == 15
    assert snapshot["output_tokens"] == 3
    assert snapshot["requests_attempted"] == 2
    assert snapshot["responses_completed"] == 2
    assert snapshot["usage_complete"] is True


def test_booleans_and_non_integers_are_not_counters():
    """A bool is an int in Python; it is not a token count."""
    collector = UsageCollector()
    collector.record_attempt()
    collector.record_response(
        _message(input_tokens=True, output_tokens="12", cache_read_input_tokens=7)
    )
    snapshot = collector.snapshot()
    assert "input_tokens" not in snapshot
    assert "output_tokens" not in snapshot
    assert snapshot["cache_read_input_tokens"] == 7


def test_missing_usage_is_unknown_rather_than_zero():
    collector = UsageCollector()
    collector.record_attempt()
    collector.record_response(types.SimpleNamespace(stop_reason="end_turn"))
    snapshot = collector.snapshot()
    assert snapshot["responses_completed"] == 1
    assert snapshot["responses_with_usage"] == 0
    assert snapshot["requests_with_unknown_usage"] == 1
    assert snapshot["usage_complete"] is False
    assert "input_tokens" not in snapshot


def test_an_attempt_that_never_answered_is_incomplete_not_free():
    """A stream that dies in transit still cost something unknown."""
    collector = UsageCollector()
    collector.record_attempt()
    snapshot = collector.snapshot()
    assert snapshot["requests_attempted"] == 1
    assert snapshot["responses_completed"] == 0
    assert snapshot["requests_with_unknown_usage"] == 1
    assert snapshot["usage_complete"] is False


def test_unknown_usage_never_reports_a_negative_count():
    collector = UsageCollector()
    collector.record_response(_message(input_tokens=3))  # response without a recorded attempt
    assert collector.snapshot()["requests_with_unknown_usage"] == 0


def test_snapshots_are_independent_copies():
    collector = UsageCollector()
    collector.record_attempt()
    collector.record_response(_message(input_tokens=1))
    first = collector.snapshot()
    collector.record_attempt()
    collector.record_response(_message(input_tokens=1))
    assert first["input_tokens"] == 1
    assert collector.snapshot()["input_tokens"] == 2


def test_concurrent_recording_loses_nothing():
    collector = UsageCollector()

    def worker():
        for _ in range(200):
            collector.record_attempt()
            collector.record_response(_message(input_tokens=1, output_tokens=1))

    threads = [threading.Thread(target=worker) for _ in range(8)]
    for thread in threads:
        thread.start()
    for thread in threads:
        thread.join()
    snapshot = collector.snapshot()
    assert snapshot["requests_attempted"] == 1600
    assert snapshot["input_tokens"] == 1600
    assert snapshot["usage_complete"] is True


def test_separate_collectors_do_not_share_counts():
    a, b = UsageCollector(), UsageCollector()
    a.record_attempt()
    a.record_response(_message(input_tokens=5))
    assert b.snapshot()["requests_attempted"] == 0
    assert "input_tokens" not in b.snapshot()


def test_usage_travels_out_on_the_exception():
    collector = UsageCollector()
    collector.record_attempt()
    collector.record_response(_message(input_tokens=42))
    error = ValueError("classification failed")
    assert attach_usage(error, collector) is error
    assert usage_from_exception(error)["input_tokens"] == 42


def test_attaching_usage_never_replaces_the_real_failure():
    error = ValueError("the real problem")
    assert attach_usage(error, None) is error
    assert usage_from_exception(error) == {}
    assert str(error) == "the real problem"


def test_usage_numbers_reads_a_bare_object_without_usage():
    assert usage_numbers(object()) == {}
    assert usage_numbers(None) == {}
