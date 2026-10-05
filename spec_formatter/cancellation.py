"""Cooperative cancellation shared by orchestration and both classifiers."""

from __future__ import annotations

import inspect
import threading
import time
from contextlib import contextmanager
from typing import Any, Callable, Optional

from .style_application.core.errors import attach_engine_error


class RunCancelled(RuntimeError):
    """Terminal control flow with a fixed, publishable engine identity."""

    def __init__(self) -> None:
        super().__init__("The run was cancelled.")
        attach_engine_error(self, "run_cancelled")


def check_cancelled(cancel_event: Optional[threading.Event]) -> None:
    if cancel_event is not None and cancel_event.is_set():
        raise RunCancelled()


def wait_for_retry(cancel_event: Optional[threading.Event], seconds: float) -> None:
    """Wake a backoff immediately when cancellation is requested."""

    if cancel_event is None:
        time.sleep(seconds)
    elif cancel_event.wait(seconds):
        raise RunCancelled()


@contextmanager
def request_slot(limiter: threading.BoundedSemaphore, cancel_event: Optional[threading.Event]):
    """Wait for a request slot without trapping a cancelled chunk in the queue."""

    if cancel_event is None:
        limiter.acquire()
    else:
        while not limiter.acquire(timeout=0.1):
            check_cancelled(cancel_event)
    try:
        check_cancelled(cancel_event)
        yield
    finally:
        limiter.release()


def cancellation_kwargs(callback: Callable[..., Any], cancel_event: Optional[threading.Event]) -> dict:
    """Pass the additive keyword only to callbacks whose signature accepts it.

    Old injected processors and classifiers keep their signatures. Never
    discover support by executing a callback and retrying on TypeError: that
    can duplicate work or conceal a TypeError from inside the callback.
    """

    if cancel_event is None:
        return {}
    try:
        parameters = inspect.signature(callback).parameters
    except (TypeError, ValueError):
        return {}
    parameter = parameters.get("cancel_event")
    if (
        parameter is not None
        and parameter.kind in (inspect.Parameter.POSITIONAL_OR_KEYWORD, inspect.Parameter.KEYWORD_ONLY)
    ) or any(p.kind == inspect.Parameter.VAR_KEYWORD for p in parameters.values()):
        return {"cancel_event": cancel_event}
    return {}
