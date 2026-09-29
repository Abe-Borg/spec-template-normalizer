"""The shared observed-usage contract, including work that failed."""

from __future__ import annotations

import json
import threading
import types

import pytest

from spec_formatter.llm_usage import (
    UsageCollector,
    attach_usage,
    refusal_category,
    streamed_stop_details,
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


def test_a_partly_reported_response_is_not_complete():
    """One recognized field is not full accounting.

    A response carrying input_tokens but no output_tokens leaves a billed
    counter unknown; reporting the total as complete would present a short
    number as final.
    """
    collector = UsageCollector()
    collector.record_attempt()
    collector.record_response(_message(input_tokens=500))
    snapshot = collector.snapshot()
    assert snapshot["responses_with_usage"] == 1
    assert snapshot["requests_with_unknown_usage"] == 1
    assert snapshot["usage_complete"] is False
    assert snapshot["input_tokens"] == 500


def test_cache_counters_are_not_required_for_completeness():
    """A response that neither read nor wrote cache omits those fields."""
    collector = UsageCollector()
    collector.record_attempt()
    collector.record_response(_message(input_tokens=10, output_tokens=2))
    assert collector.snapshot()["usage_complete"] is True


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


def test_refusal_category_is_accepted_by_shape_and_nothing_else():
    def refused(details):
        return types.SimpleNamespace(stop_reason="refusal", stop_details=details)

    # The category set is open, so a category added later is still read.
    assert refusal_category(refused(types.SimpleNamespace(category="general_harms"))) == "general_harms"
    assert refusal_category(refused({"category": "reasoning_extraction"})) == "reasoning_extraction"
    assert refusal_category(refused(types.SimpleNamespace(category="a_future_category"))) == "a_future_category"
    # No category, no details, or anything that is not a short identifier.
    assert refusal_category(refused(types.SimpleNamespace(category=None))) is None
    assert refusal_category(refused(None)) is None
    assert refusal_category(types.SimpleNamespace(stop_reason="end_turn")) is None
    assert refusal_category(None) is None
    for unsafe in ("Cyber", "two words", "cyber\n", "x" * 49, "", 7, True):
        assert refusal_category(refused({"category": unsafe})) is None


def _sse_client(*bodies):
    """A real pinned-SDK client whose requests replay recorded SSE bodies.

    Fabricated final messages cannot show what the SDK's stream accumulator
    drops; replaying the wire events through the real client can.
    """
    import httpx
    import anthropic

    pending = list(bodies)

    def handler(_request):
        return httpx.Response(
            200,
            headers={"content-type": "text/event-stream"},
            content=pending.pop(0),
        )

    return anthropic.Anthropic(
        api_key="k",
        max_retries=0,
        http_client=httpx.Client(transport=httpx.MockTransport(handler)),
    )


def _sse_body(stop_reason, *, text=None, stop_details=None):
    """The wire events of one streamed response, as the API sends them."""
    events = [
        {
            "type": "message_start",
            "message": {
                "id": "msg_1",
                "type": "message",
                "role": "assistant",
                "model": "m",
                "content": [],
                "stop_reason": None,
                "stop_sequence": None,
                "usage": {"input_tokens": 10, "output_tokens": 0},
            },
        }
    ]
    if text is not None:
        events += [
            {"type": "content_block_start", "index": 0, "content_block": {"type": "text", "text": ""}},
            {"type": "content_block_delta", "index": 0, "delta": {"type": "text_delta", "text": text}},
            {"type": "content_block_stop", "index": 0},
        ]
    delta = {"stop_reason": stop_reason, "stop_sequence": None}
    if stop_details is not None:
        delta["stop_details"] = stop_details
    events += [
        {"type": "message_delta", "delta": delta, "usage": {"output_tokens": 3}},
        {"type": "message_stop"},
    ]
    return "".join(
        f"event: {event['type']}\ndata: {json.dumps(event)}\n\n" for event in events
    ).encode("utf-8")


def test_a_streamed_refusal_category_survives_the_pinned_sdk():
    """The category is on the raw message_delta, not on the final message.

    The pinned SDK's accumulator copies stop_reason, stop_sequence and usage
    off message_delta but not stop_details, so reading the final message
    alone always found no category.
    """
    client = _sse_client(
        _sse_body(
            "refusal",
            stop_details={
                "type": "refusal",
                "category": "general_harms",
                "explanation": "Free text the provider wrote.",
            },
        )
    )
    with client.messages.stream(
        model="m", max_tokens=16, messages=[{"role": "user", "content": "x"}]
    ) as stream:
        details = streamed_stop_details(stream)
        final_message = stream.get_final_message()

    assert final_message.stop_reason == "refusal"
    assert final_message.content == []
    assert refusal_category(final_message, details) == "general_harms"


def test_streamed_stop_details_tolerates_a_stream_it_cannot_iterate():
    class FinalOnly:
        def get_final_message(self):
            return types.SimpleNamespace(stop_reason="refusal")

    assert streamed_stop_details(FinalOnly()) is None
    # The final message wins when an SDK carries the details there.
    message = types.SimpleNamespace(stop_details={"category": "cyber"})
    assert refusal_category(message, {"category": "bio"}) == "cyber"
