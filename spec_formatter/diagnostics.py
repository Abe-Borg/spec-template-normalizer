"""Structured, redaction-safe diagnostics for a formatting run.

This module is the detailed-diagnostics channel that complements the human
``run.log``.  Where ``run.log`` records product-language progress, the
diagnostics recorder captures **structured** events -- phase timings and
quantitative counts -- that are cheap to reason about after the fact
(``diagnostics.jsonl``) and to roll up (the ``diagnostics`` block in
``run.json``).

Everything here obeys the same non-negotiable boundary as the rest of the
pipeline: **no secrets and no document text may ever reach a run artifact.**
The recorder never accepts free-form prose.  Every event field is passed
through :func:`sanitize_fields`, which keeps only JSON scalars and short
identifier-shaped strings (the same character class the pipeline already
trusts for ``run.json`` structural fields).  A field value that could carry a
paragraph of a customer's document -- anything with a space, punctuation
outside the safe class, or excess length -- is dropped, not truncated.  The
pipeline additionally redacts the API key from every serialized event as
defense in depth.

The recorder is thread-safe: targets are formatted on a ``ThreadPoolExecutor``
and each records into the same shared recorder, so every event is assigned its
sequence number and timestamp under a lock.  The engine layer, which must stay
free of any recorder object so :class:`BatchResult` remains a plain data
carrier, produces events as ordinary dicts via :func:`timed`/:func:`emit`; the
pipeline folds them in with :meth:`DiagnosticsRecorder.ingest`.
"""

from __future__ import annotations

import re
import threading
import time
from contextlib import contextmanager
from dataclasses import dataclass, field
from datetime import datetime, timezone
from typing import Any, Callable, Dict, Iterable, Iterator, List, Mapping, Optional, Tuple


# Severity levels, deliberately mirroring the stdlib ``logging`` numeric order
# without importing it (the pipeline avoids process-global logging state).
DEBUG = 10
INFO = 20
WARNING = 30
ERROR = 40

_LEVEL_TO_NAME = {DEBUG: "DEBUG", INFO: "INFO", WARNING: "WARNING", ERROR: "ERROR"}
_NAME_TO_LEVEL = {name: value for value, name in _LEVEL_TO_NAME.items()}

# The one character class we trust to reach disk in a string field.  Identical
# in spirit to ``pipeline._normalize_audit_details``: identifiers, numbers,
# dotted/pathy tokens and a few symbols -- but never whitespace, which is the
# cheapest tell that a value is a sentence from a document rather than a code.
_SAFE_STRING_RX = re.compile(r"[A-Za-z0-9_.:/#@+\-]{1,160}")
_NAME_RX = re.compile(r"[a-z][a-z0-9_]{0,47}")

_MAX_FIELD_KEYS = 64
_MAX_LIST_ITEMS = 256
_MAX_DEPTH = 6

# A dropped-in-place sentinel so ``sanitize_fields`` can distinguish "value was
# rejected" from "value was legitimately ``None``".
_DROP = object()


def level_from_name(name: str, *, default: int = INFO) -> int:
    """Return the numeric level for ``name`` (case-insensitive), or *default*."""

    if not isinstance(name, str):
        return default
    return _NAME_TO_LEVEL.get(name.strip().upper(), default)


def level_name(level: int) -> str:
    """Return the canonical name for a numeric *level*."""

    return _LEVEL_TO_NAME.get(level, "INFO")


def _key_may_carry_document_text(key: str) -> bool:
    """Reject field keys whose name invites free document text.

    Mirrors ``pipeline._audit_key_may_contain_document_text`` so the two
    redaction surfaces cannot drift apart.
    """

    folded = key.casefold()
    return any(
        token in folded
        for token in ("text", "preview", "excerpt", "content", "paragraph_xml")
    )


def _sanitize_value(value: Any, *, depth: int) -> Any:
    if value is None or isinstance(value, (bool, int, float)):
        return value
    if isinstance(value, str):
        return value if _SAFE_STRING_RX.fullmatch(value) else _DROP
    if depth >= _MAX_DEPTH:
        return _DROP
    if isinstance(value, Mapping):
        nested: Dict[str, Any] = {}
        for raw_key, raw_value in value.items():
            if not isinstance(raw_key, str) or _key_may_carry_document_text(raw_key):
                continue
            if not _NAME_RX.fullmatch(raw_key) and not _SAFE_STRING_RX.fullmatch(raw_key):
                continue
            cleaned = _sanitize_value(raw_value, depth=depth + 1)
            if cleaned is not _DROP:
                nested[raw_key] = cleaned
            if len(nested) >= _MAX_FIELD_KEYS:
                break
        return nested
    if isinstance(value, (list, tuple)):
        items: List[Any] = []
        for element in value:
            cleaned = _sanitize_value(element, depth=depth + 1)
            if cleaned is not _DROP:
                items.append(cleaned)
            if len(items) >= _MAX_LIST_ITEMS:
                break
        return items
    return _DROP


def sanitize_fields(fields: Optional[Mapping[str, Any]]) -> Dict[str, Any]:
    """Return only the provably safe scalar/identifier subset of *fields*.

    Unsafe values are omitted entirely.  This is intentionally strict: a
    diagnostics field exists to be aggregated, so anything that is not a
    number, bool, ``None`` or an identifier-shaped string has no business in
    it and is far more likely to be leaked document text than a useful datum.
    """

    if not isinstance(fields, Mapping):
        return {}
    safe: Dict[str, Any] = {}
    for key, value in fields.items():
        if not isinstance(key, str) or _key_may_carry_document_text(key):
            continue
        if not _NAME_RX.fullmatch(key) and not _SAFE_STRING_RX.fullmatch(key):
            continue
        cleaned = _sanitize_value(value, depth=1)
        if cleaned is not _DROP:
            safe[key] = cleaned
        if len(safe) >= _MAX_FIELD_KEYS:
            break
    return safe


def _clean_name(value: Any) -> str:
    if isinstance(value, str) and _NAME_RX.fullmatch(value):
        return value
    return "invalid"


def sanitize_event(raw: Any) -> Dict[str, Any]:
    """Return a fully sanitized plain-dict form of one engine event.

    Used at every write boundary that serializes engine-produced events
    directly (for example the per-target ``audit.json``), so a buggy or
    hostile processor cannot smuggle document text through a diagnostics
    field even if it bypasses the recorder's ``ingest`` path.
    """

    if not isinstance(raw, Mapping):
        return {}
    event: Dict[str, Any] = {
        "level": level_name(level_from_name(str(raw.get("level", "INFO")))),
        "component": _clean_name(str(raw.get("component", "engine"))),
        "event": _clean_name(str(raw.get("event", "event"))),
        "fields": sanitize_fields(
            raw.get("fields") if isinstance(raw.get("fields"), Mapping) else {}
        ),
    }
    if isinstance(raw.get("target"), int) and not isinstance(raw.get("target"), bool):
        event["target"] = raw["target"]
    return event


@dataclass(frozen=True)
class DiagnosticEvent:
    """One structured diagnostic record."""

    seq: int
    ts: str
    level: int
    component: str
    event: str
    fields: Dict[str, Any] = field(default_factory=dict)
    target: Optional[int] = None

    def as_dict(self) -> Dict[str, Any]:
        record: Dict[str, Any] = {
            "seq": self.seq,
            "ts": self.ts,
            "level": level_name(self.level),
            "component": self.component,
            "event": self.event,
        }
        if self.target is not None:
            record["target"] = self.target
        record["fields"] = dict(self.fields)
        return record


# ---------------------------------------------------------------------------
# Engine-layer helpers: produce plain event dicts without a recorder object so
# ``BatchResult`` stays a serializable data carrier.  The pipeline re-validates
# everything on ``ingest``.
# ---------------------------------------------------------------------------


def emit(
    collector: List[Dict[str, Any]],
    level: str,
    component: str,
    event: str,
    **fields: Any,
) -> None:
    """Append one structured event dict to *collector*."""

    collector.append(
        {
            "level": level.upper() if isinstance(level, str) else "INFO",
            "component": _clean_name(component),
            "event": _clean_name(event),
            "fields": sanitize_fields(fields),
        }
    )


class _PhaseHandle:
    """Mutable handle yielded by :func:`timed` to attach result counts."""

    __slots__ = ("fields",)

    def __init__(self) -> None:
        self.fields: Dict[str, Any] = {}

    def set(self, **fields: Any) -> None:
        self.fields.update(fields)


@contextmanager
def timed(
    collector: List[Dict[str, Any]],
    component: str,
    event: str,
    *,
    level: str = "INFO",
    **fields: Any,
) -> Iterator[_PhaseHandle]:
    """Time a phase and append a structured event when it exits.

    On success the event carries ``duration_ms`` plus any fields supplied up
    front or via the yielded handle.  On failure it is recorded at ``ERROR``
    with ``failed=True`` and the exception *type name* only (never its
    message, which can echo document text), then the exception re-raises.
    """

    handle = _PhaseHandle()
    start = time.monotonic()
    try:
        yield handle
    except BaseException as exc:  # noqa: BLE001 - re-raised after recording
        duration_ms = round((time.monotonic() - start) * 1000.0, 3)
        merged = {**fields, **handle.fields, "duration_ms": duration_ms, "failed": True,
                  "error_type": _clean_name(type(exc).__name__.lower())}
        emit(collector, "ERROR", component, event, **merged)
        raise
    duration_ms = round((time.monotonic() - start) * 1000.0, 3)
    emit(collector, level, component, event, **fields, **handle.fields,
         duration_ms=duration_ms)


class DiagnosticsRecorder:
    """Thread-safe collector of structured diagnostic events for one run."""

    def __init__(
        self,
        *,
        min_level: int = INFO,
        clock: Optional[Callable[[], datetime]] = None,
    ) -> None:
        self._min_level = min_level
        self._clock = clock or (lambda: datetime.now(timezone.utc))
        self._lock = threading.Lock()
        self._events: List[DiagnosticEvent] = []
        self._seq = 0

    @property
    def min_level(self) -> int:
        return self._min_level

    def _now_iso(self) -> str:
        value = self._clock()
        if value.tzinfo is None:
            value = value.replace(tzinfo=timezone.utc)
        return value.astimezone(timezone.utc).isoformat().replace("+00:00", "Z")

    def _store(
        self,
        level: int,
        component: str,
        event: str,
        target: Optional[int],
        fields: Dict[str, Any],
    ) -> Optional[DiagnosticEvent]:
        if level < self._min_level:
            return None
        with self._lock:
            self._seq += 1
            record = DiagnosticEvent(
                seq=self._seq,
                ts=self._now_iso(),
                level=level,
                component=_clean_name(component),
                event=_clean_name(event),
                fields=fields,
                target=target if isinstance(target, int) else None,
            )
            self._events.append(record)
        return record

    def record(
        self,
        level: int,
        component: str,
        event: str,
        *,
        target: Optional[int] = None,
        **fields: Any,
    ) -> Optional[DiagnosticEvent]:
        return self._store(level, component, event, target, sanitize_fields(fields))

    def debug(self, component: str, event: str, **kwargs: Any) -> Optional[DiagnosticEvent]:
        return self.record(DEBUG, component, event, **kwargs)

    def info(self, component: str, event: str, **kwargs: Any) -> Optional[DiagnosticEvent]:
        return self.record(INFO, component, event, **kwargs)

    def warning(self, component: str, event: str, **kwargs: Any) -> Optional[DiagnosticEvent]:
        return self.record(WARNING, component, event, **kwargs)

    def error(self, component: str, event: str, **kwargs: Any) -> Optional[DiagnosticEvent]:
        return self.record(ERROR, component, event, **kwargs)

    @contextmanager
    def timer(
        self,
        component: str,
        event: str,
        *,
        target: Optional[int] = None,
        level: int = INFO,
        **fields: Any,
    ) -> Iterator[_PhaseHandle]:
        """Record *event* with ``duration_ms`` when the block exits.

        Failures record at ``ERROR`` with ``failed=True`` and the exception
        type name only, then re-raise.
        """

        handle = _PhaseHandle()
        start = time.monotonic()
        try:
            yield handle
        except BaseException as exc:  # noqa: BLE001 - re-raised after recording
            duration_ms = round((time.monotonic() - start) * 1000.0, 3)
            merged = {
                **fields,
                **handle.fields,
                "duration_ms": duration_ms,
                "failed": True,
                "error_type": type(exc).__name__.lower(),
            }
            self._store(ERROR, component, event, target, sanitize_fields(merged))
            raise
        duration_ms = round((time.monotonic() - start) * 1000.0, 3)
        merged = {**fields, **handle.fields, "duration_ms": duration_ms}
        self._store(level, component, event, target, sanitize_fields(merged))

    def ingest(
        self,
        raw_events: Optional[Iterable[Mapping[str, Any]]],
        *,
        target: Optional[int] = None,
    ) -> None:
        """Fold engine-produced event dicts into this recorder.

        Each incoming dict is treated as untrusted: its level, component,
        event name and fields are all re-validated so a buggy or hostile
        processor cannot smuggle document text through the diagnostics
        channel.  A shared *target* index is stamped onto every event.
        """

        if not raw_events:
            return
        for raw in raw_events:
            if not isinstance(raw, Mapping):
                continue
            level = level_from_name(str(raw.get("level", "INFO")))
            component = str(raw.get("component", "engine"))
            event = str(raw.get("event", "event"))
            fields = sanitize_fields(raw.get("fields") if isinstance(raw.get("fields"), Mapping) else {})
            self._store(level, component, event, target, fields)

    def snapshot(self) -> Tuple[DiagnosticEvent, ...]:
        with self._lock:
            return tuple(self._events)

    def iter_dicts(self) -> List[Dict[str, Any]]:
        return [event.as_dict() for event in self.snapshot()]

    def summary(self) -> Dict[str, Any]:
        """Return a compact, JSON-safe rollup for ``run.json``."""

        events = self.snapshot()
        counts_by_level: Dict[str, int] = {}
        phase_durations: Dict[str, float] = {}
        phase_counts: Dict[str, int] = {}
        targets: set = set()
        for record in events:
            name = level_name(record.level)
            counts_by_level[name] = counts_by_level.get(name, 0) + 1
            if record.target is not None:
                targets.add(record.target)
            duration = record.fields.get("duration_ms")
            if isinstance(duration, (int, float)) and not isinstance(duration, bool):
                key = f"{record.component}.{record.event}"
                phase_durations[key] = round(phase_durations.get(key, 0.0) + float(duration), 3)
                phase_counts[key] = phase_counts.get(key, 0) + 1
        return {
            "level": level_name(self._min_level),
            "event_count": len(events),
            "counts_by_level": counts_by_level,
            "warnings": counts_by_level.get("WARNING", 0),
            "errors": counts_by_level.get("ERROR", 0),
            "targets_observed": len(targets),
            "phase_durations_ms": phase_durations,
            "phase_counts": phase_counts,
            "log": "diagnostics.jsonl",
        }


__all__ = [
    "DEBUG",
    "INFO",
    "WARNING",
    "ERROR",
    "DiagnosticEvent",
    "DiagnosticsRecorder",
    "emit",
    "level_from_name",
    "level_name",
    "sanitize_event",
    "sanitize_fields",
    "timed",
]
