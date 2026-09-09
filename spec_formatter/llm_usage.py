"""Observed model usage for one run, including work that failed.

Both classifiers used to lose token counts. The architect path never read
``usage`` off a response at all; the target path collected counts correctly
but published them only after a successful merge, so a refusal, an exhausted
regeneration, or a merge failure discarded everything the run had already
paid for.

This module owns the one counting contract for both. It records what the
provider actually reported and, just as importantly, records when it could
not: a request whose final usage is unknown is counted as unknown, never as
zero, so a total is never quietly understated. ``usage_complete`` says
whether the totals account for every request.

Scope note: this is observation, not billing reconciliation. The provider's
invoice remains authoritative, and a snapshot with ``usage_complete`` false
is a lower bound.
"""

from __future__ import annotations

import threading
from typing import Any, Dict, Optional

#: Provider usage counters this project understands. A field the provider
#: adds later is ignored rather than guessed at.
USAGE_FIELDS = (
    "input_tokens",
    "output_tokens",
    "cache_read_input_tokens",
    "cache_creation_input_tokens",
)

#: Counters every completed response must report for its usage to be counted
#: as fully known. The cache fields are deliberately absent: a response that
#: neither read nor wrote cache legitimately omits them, so requiring them
#: would mark ordinary responses incomplete. Input and output are always
#: billed, so a response missing either is only partly accounted for.
REQUIRED_USAGE_FIELDS = ("input_tokens", "output_tokens")


def usage_numbers(final_message: Any) -> Dict[str, int]:
    """Return the recognized integer counters on ``final_message.usage``.

    A ``bool`` is not accepted as a counter even though it is an ``int``,
    and a missing or non-integer field is left out rather than read as zero.
    """

    usage = getattr(final_message, "usage", None)
    numbers: Dict[str, int] = {}
    for name in USAGE_FIELDS:
        value = getattr(usage, name, None)
        if isinstance(value, bool) or not isinstance(value, int):
            continue
        numbers[name] = value
    return numbers


class UsageCollector:
    """Thread-safe observed-usage counters for one classification scope.

    Target chunks run on a worker pool and the architect path is
    single-threaded; one lock covers both. Every method is safe to call from
    an exception path, because losing counts is the failure this exists to
    prevent.
    """

    def __init__(self) -> None:
        self._lock = threading.Lock()
        self._requests_attempted = 0
        self._responses_completed = 0
        self._responses_with_usage = 0
        self._responses_fully_accounted = 0
        self._totals: Dict[str, int] = {}

    def record_attempt(self) -> None:
        """Count a request that is about to be sent.

        Called before the stream opens, so a request that dies in transit
        still appears in the totals as attempted with unknown usage.
        """

        with self._lock:
            self._requests_attempted += 1

    def record_response(self, final_message: Any) -> None:
        """Count a completed response and add whatever usage it reported.

        Call this the moment the final message is in hand and before acting
        on its stop reason: a refusal and an output-limit response are both
        paid for, and both used to raise before anything read their usage.
        """

        numbers = usage_numbers(final_message)
        complete = all(name in numbers for name in REQUIRED_USAGE_FIELDS)
        with self._lock:
            self._responses_completed += 1
            if numbers:
                self._responses_with_usage += 1
            if complete:
                self._responses_fully_accounted += 1
            for name, value in numbers.items():
                self._totals[name] = self._totals.get(name, 0) + value

    def snapshot(self) -> Dict[str, Any]:
        """Return the counters so far as a plain, JSON-safe dict.

        Safe to call at any time, including while work is still running and
        from an exception handler.
        """

        with self._lock:
            # Completeness is measured against responses that reported every
            # required counter, not merely one. A response carrying
            # input_tokens but no output_tokens is partly unknown, and
            # counting it as known would present a short total as final.
            unknown = self._requests_attempted - self._responses_fully_accounted
            snapshot: Dict[str, Any] = {
                "requests_attempted": self._requests_attempted,
                "responses_completed": self._responses_completed,
                "responses_with_usage": self._responses_with_usage,
                "requests_with_unknown_usage": max(unknown, 0),
                "usage_complete": unknown <= 0,
            }
            snapshot.update(self._totals)
            return snapshot


def attach_usage(exc: BaseException, usage: Optional[UsageCollector]) -> BaseException:
    """Record a usage snapshot on ``exc`` and return it for re-raising.

    Classification failures propagate as ordinary exceptions, and the counts
    observed before the failure would go with them. Carrying a snapshot on
    the exception keeps the failure type and message exactly as they were
    while letting the caller record what the attempt cost.
    """

    if usage is not None:
        try:
            setattr(exc, "observed_usage", usage.snapshot())
        except Exception:
            # Never let accounting replace the real failure.
            pass
    return exc


def usage_from_exception(exc: BaseException) -> Dict[str, Any]:
    """Return the snapshot :func:`attach_usage` left on ``exc``, if any."""

    observed = getattr(exc, "observed_usage", None)
    return observed if isinstance(observed, dict) else {}
