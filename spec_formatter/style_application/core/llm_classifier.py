"""
LLM-based classification for Phase 2.

Sends paragraph bundles to the Anthropic API for CSI role classification,
with retry logic, chunking for large documents, and coverage reporting.
"""

import json
import os
import threading
import time
import re
from concurrent.futures import ThreadPoolExecutor, as_completed
from typing import Any, Dict, List, Optional, Set

from .classification import (
    PHASE2_MASTER_PROMPT,
    PHASE2_RUN_INSTRUCTION,
    coerce_to_final_classifications,
)


# Sonnet 5's tokenizer produces ~30% more tokens for the same text than the
# pre-4.7 tokenizers, so estimate conservatively at ~3 chars/token.
_CHARS_PER_TOKEN = 3
_MAX_BUNDLE_TOKENS = 80_000
_MAX_BUNDLE_CHARS = _MAX_BUNDLE_TOKENS * _CHARS_PER_TOKEN
_CHUNK_OVERLAP = 20


# Transport policy -----------------------------------------------------------
#
# The SDK's hidden retries are disabled so the bounded policy below owns every
# attempt; a bad key used to cost up to 54 doomed requests per target. A
# process-wide limiter bounds concurrent requests across all targets and all
# of their chunks (six targets times six chunk workers used to mean 36 open
# streams).

_MAX_CONCURRENT_REQUESTS_ENV = "SPEC_FORMATTER_MAX_CONCURRENT_REQUESTS"
_DEFAULT_MAX_CONCURRENT_REQUESTS = 4
_MAX_CONCURRENT_REQUESTS_CEILING = 64
_TRANSPORT_RETRIES = 2  # initial request + 2 retries for transient failures
_MAX_RETRY_AFTER_SECONDS = 120.0


class ClassificationRefused(RuntimeError):
    """The model refused the request (``stop_reason == "refusal"``).

    Terminal: a refusal is not transient and regenerating the same request
    would not change the outcome. No server-side fallback is attempted
    because the run manifest records the model that produced the result.
    """


class _NeverRaised(Exception):
    """Placeholder for SDK exception types missing from a stubbed module."""


def _max_concurrent_requests() -> int:
    raw = os.environ.get(_MAX_CONCURRENT_REQUESTS_ENV, "").strip()
    try:
        value = int(raw) if raw else _DEFAULT_MAX_CONCURRENT_REQUESTS
    except ValueError:
        value = _DEFAULT_MAX_CONCURRENT_REQUESTS
    return max(1, min(value, _MAX_CONCURRENT_REQUESTS_CEILING))


_REQUEST_LIMITER = threading.BoundedSemaphore(_max_concurrent_requests())


def _configure_request_limit(limit: int) -> None:
    """Replace the process-wide limiter (tests and embedding applications)."""

    global _REQUEST_LIMITER
    _REQUEST_LIMITER = threading.BoundedSemaphore(max(1, int(limit)))


def _sdk_error_classes(anthropic_module: Any) -> Dict[str, type]:
    return {
        name: getattr(anthropic_module, name, _NeverRaised)
        for name in ("RateLimitError", "APIConnectionError", "APIStatusError")
    }


def _retry_after_seconds(error: Exception) -> Optional[float]:
    """Honour a numeric ``retry-after`` header when the SDK exposes one."""

    response = getattr(error, "response", None)
    headers = getattr(response, "headers", None)
    if headers is None:
        return None
    try:
        raw = headers.get("retry-after")
    except Exception:
        return None
    if raw is None:
        return None
    try:
        seconds = float(str(raw).strip())
    except ValueError:
        return None
    if seconds < 0:
        return None
    return min(seconds, _MAX_RETRY_AFTER_SECONDS)


def _transport_retry_delay(
    error: Exception,
    attempt: int,
    errors: Dict[str, type],
) -> Optional[float]:
    """Seconds to wait before retrying ``error``, or ``None`` to fail now.

    Rate limits wait for ``retry-after`` when present, connection failures
    and 5xx responses back off exponentially, and everything else (bad key,
    bad request, refusal, programming errors) is re-raised at once.
    """

    backoff = float(2 ** (attempt + 1))
    if isinstance(error, errors["RateLimitError"]):
        retry_after = _retry_after_seconds(error)
        return retry_after if retry_after is not None else backoff
    if isinstance(error, errors["APIConnectionError"]):
        return backoff
    if isinstance(error, errors["APIStatusError"]):
        status = getattr(error, "status_code", 0) or 0
        return backoff if status >= 500 else None
    return None


def _final_stop_reason(stream: Any) -> Optional[str]:
    get_final_message = getattr(stream, "get_final_message", None)
    if get_final_message is None:
        return None
    return getattr(get_final_message(), "stop_reason", None)


def _build_user_message(slim_bundle: dict, available_roles: list) -> str:
    # Only unresolved paragraphs belong in the model request.  The complete
    # bundle also contains deterministic classifications and filtered paragraph
    # indices for local audit/reassembly; exposing those indices invites the
    # model to echo entries that validation correctly rejects.
    prompt_bundle = {"paragraphs": slim_bundle.get("paragraphs", [])}
    if "_chunk_info" in slim_bundle:
        prompt_bundle["_chunk_info"] = slim_bundle["_chunk_info"]
    return (
        PHASE2_RUN_INSTRUCTION.strip()
        + "\n\navailable_roles: " + json.dumps(available_roles)
        + "\n\n" + json.dumps(prompt_bundle, indent=2)
    )


def _retry_requirement(error: Exception, allowed_indices: Set[int]) -> str:
    return (
        "\n\nRETRY REQUIREMENT: The prior response failed validation: "
        f"{error}. Return exactly one complete JSON object with no prose, "
        "Markdown fence, or text before or after it. Across classifications and "
        "ignored_paragraphs, return exactly one disposition for each of these "
        "paragraph_index values and no other indices: "
        f"{json.dumps(sorted(allowed_indices))}."
    )


def _classification_output_config(available_roles: list) -> dict:
    """Return the shared Anthropic structured-output contract."""

    return {
        "effort": "high",
        "format": {
            "type": "json_schema",
            "schema": {
                "type": "object",
                "additionalProperties": False,
                "properties": {
                    "classifications": {
                        "type": "array",
                        "items": {
                            "type": "object",
                            "additionalProperties": False,
                            "properties": {
                                "paragraph_index": {"type": "integer"},
                                "csi_role": {
                                    "type": "string",
                                    "enum": list(available_roles),
                                },
                            },
                            "required": ["paragraph_index", "csi_role"],
                        },
                    },
                    "ignored_paragraphs": {
                        "type": "array",
                        "items": {
                            "type": "object",
                            "additionalProperties": False,
                            "properties": {
                                "paragraph_index": {"type": "integer"},
                                "reason": {
                                    "type": "string",
                                    "enum": ["non_csi_content"],
                                },
                            },
                            "required": ["paragraph_index", "reason"],
                        },
                    },
                    "notes": {
                        "type": "array",
                        "items": {"type": "string"},
                    },
                },
                "required": ["classifications", "ignored_paragraphs", "notes"],
            },
        },
    }


def _parse_classification_response(response_text: str) -> dict:
    text = response_text.lstrip("\ufeff").strip()
    try:
        return json.loads(text)
    except json.JSONDecodeError as direct_error:
        # Models occasionally wrap an otherwise valid object in prose, XML-ish
        # tags, or a Markdown fence.  Recover only when there is exactly one
        # distinct object containing the expected top-level field; ambiguous
        # or partial output remains a hard parse failure and is retried.
        decoder = json.JSONDecoder()
        candidates: Dict[str, dict] = {}
        for match in re.finditer(r"\{", text):
            try:
                value, _end = decoder.raw_decode(text, match.start())
            except json.JSONDecodeError:
                continue
            if not isinstance(value, dict) or "classifications" not in value:
                continue
            canonical = json.dumps(
                value,
                sort_keys=True,
                separators=(",", ":"),
                ensure_ascii=False,
            )
            candidates[canonical] = value

        if len(candidates) == 1:
            return next(iter(candidates.values()))
        if len(candidates) > 1:
            raise json.JSONDecodeError(
                "multiple distinct classification JSON objects",
                text,
                0,
            ) from direct_error
        raise


def _validate_classifications(classifications: dict, available_roles: list, allowed_indices: Set[int]) -> dict:
    if not isinstance(classifications, dict):
        raise ValueError("LLM response is not a JSON object")
    items = classifications.get("classifications", [])
    if not isinstance(items, list):
        raise ValueError("LLM response missing 'classifications' array")

    valid_roles = set(available_roles)
    validated = []
    seen_indices = set()
    for item in items:
        if not isinstance(item, dict):
            raise ValueError("all classification entries must be objects")
        idx = item.get("paragraph_index")
        role = item.get("csi_role")
        if not isinstance(idx, int):
            raise ValueError(f"invalid paragraph_index: {idx!r}")
        if idx in seen_indices:
            raise ValueError(f"duplicate classification for paragraph_index={idx}")
        if idx not in allowed_indices:
            raise ValueError(f"classification index not allowed: {idx}")
        if not isinstance(role, str) or role not in valid_roles:
            raise ValueError(f"invalid csi_role for paragraph_index={idx}: {role!r}")
        seen_indices.add(idx)
        validated.append({"paragraph_index": idx, "csi_role": role})

    raw_ignored = classifications.get("ignored_paragraphs", [])
    if not isinstance(raw_ignored, list):
        raise ValueError("LLM response 'ignored_paragraphs' must be an array")
    validated_ignored = []
    for item in raw_ignored:
        if not isinstance(item, dict):
            raise ValueError("all ignored paragraph entries must be objects")
        idx = item.get("paragraph_index")
        reason = item.get("reason")
        if not isinstance(idx, int):
            raise ValueError(f"invalid ignored paragraph_index: {idx!r}")
        if idx in seen_indices:
            raise ValueError(
                f"duplicate/conflicting disposition for paragraph_index={idx}"
            )
        if idx not in allowed_indices:
            raise ValueError(f"ignored index not allowed: {idx}")
        if reason != "non_csi_content":
            raise ValueError(
                f"invalid ignored reason for paragraph_index={idx}: {reason!r}"
            )
        seen_indices.add(idx)
        validated_ignored.append({
            "paragraph_index": idx,
            "reason": reason,
        })

    missing = sorted(allowed_indices - seen_indices)
    if missing:
        raise ValueError(f"missing coverage for paragraph indices: {missing[:20]}")

    return {
        "classifications": validated,
        "ignored_paragraphs": validated_ignored,
        "notes": classifications.get("notes", []),
    }


def _split_bundle_into_chunks(slim_bundle: dict, max_chars: int = _MAX_BUNDLE_CHARS) -> List[dict]:
    paragraphs = slim_bundle.get("paragraphs", [])
    roles = slim_bundle.get("available_roles", [])
    filter_report = slim_bundle.get("filter_report", {})

    full_json = json.dumps({"paragraphs": paragraphs})
    if len(full_json) <= max_chars and len(paragraphs) <= 300:
        return [slim_bundle]

    overhead = len(json.dumps({
        "available_roles": roles,
        "filter_report": {"paragraphs_removed_entirely": [], "paragraphs_stripped": []},
        "paragraphs": []
    }))
    # Chunks carry an emptied filter_report, so size them from the paragraph
    # payload alone (a filter_report-dominated bundle would wildly inflate the
    # per-paragraph average). Clamp the chunk size above _CHUNK_OVERLAP so the
    # window always advances — otherwise `start = end - _CHUNK_OVERLAP` can
    # move backwards and loop forever.
    avg_para_size = len(json.dumps(paragraphs)) / max(len(paragraphs), 1)
    paras_per_chunk = max(_CHUNK_OVERLAP + 10, int((max_chars - overhead) / max(avg_para_size, 1)))

    chunks = []
    start = 0
    while start < len(paragraphs):
        end = min(start + paras_per_chunk, len(paragraphs))
        chunk_paras = paragraphs[start:end]
        chunk = {
            "available_roles": roles,
            "filter_report": {"paragraphs_removed_entirely": [], "paragraphs_stripped": []},
            "paragraphs": chunk_paras,
            "_chunk_info": {
                "chunk_index": len(chunks),
                "paragraph_range": [chunk_paras[0]["paragraph_index"], chunk_paras[-1]["paragraph_index"]] if chunk_paras else [0, 0]
            }
        }
        chunks.append(chunk)
        start = end - _CHUNK_OVERLAP if end < len(paragraphs) else end
    return chunks


def _merge_chunk_results(chunk_results: List[dict]) -> dict:
    seen: Dict[int, tuple[str, str]] = {}
    conflicts: List[Dict[str, Any]] = []
    all_notes: List[Any] = []

    for result in chunk_results:
        for item in result.get("classifications", []):
            idx = item.get("paragraph_index")
            role = item.get("csi_role")
            if idx is None or role is None:
                continue
            disposition = ("classified", role)
            prior = seen.get(idx)
            if prior is not None and prior != disposition:
                conflicts.append({
                    "paragraph_index": idx,
                    "existing": prior,
                    "conflicting": disposition,
                })
            seen[idx] = disposition
        for item in result.get("ignored_paragraphs", []):
            idx = item.get("paragraph_index")
            reason = item.get("reason")
            if idx is None or reason is None:
                continue
            disposition = ("ignored", reason)
            prior = seen.get(idx)
            if prior is not None and prior != disposition:
                conflicts.append({
                    "paragraph_index": idx,
                    "existing": prior,
                    "conflicting": disposition,
                })
            seen[idx] = disposition
        all_notes.extend(result.get("notes", []))

    if conflicts:
        examples = ", ".join(
            f"{c['paragraph_index']}:{c['existing']}|{c['conflicting']}"
            for c in conflicts[:10]
        )
        raise ValueError(f"Chunk merge conflicts detected ({len(conflicts)} total): {examples}")

    return {
        "classifications": [
            {"paragraph_index": idx, "csi_role": value}
            for idx, (kind, value) in sorted(seen.items())
            if kind == "classified"
        ],
        "ignored_paragraphs": [
            {"paragraph_index": idx, "reason": value}
            for idx, (kind, value) in sorted(seen.items())
            if kind == "ignored"
        ],
        "notes": all_notes,
    }


def classify_target_document(slim_bundle: dict, available_roles: list, api_key: str, model: str = "claude-sonnet-5") -> dict:
    unresolved_paragraphs = slim_bundle.get("paragraphs", [])
    if not unresolved_paragraphs:
        deterministic_only = coerce_to_final_classifications(
            slim_bundle,
            {
                "classifications": [],
                "ignored_paragraphs": [],
                "notes": ["LLM skipped: all paragraphs classified deterministically."],
            },
            available_roles,
        )
        total_expected = (
            len(deterministic_only.get("classifications", []))
            + len(deterministic_only.get("ignored_paragraphs", []))
        )
        print("LLM skipped: all paragraphs resolved deterministically.")
        print(f"Disposition coverage: {total_expected}/{total_expected} (100.0%)")
        return deterministic_only

    import anthropic
    import httpx

    # The classifier owns the retry policy: no SDK retries, a short connect
    # timeout so an unreachable endpoint fails fast, and the SDK-default
    # 10-minute read window for long adaptive-thinking turns. This mirrors the
    # architect-side client in llm_classifier._call_api.
    client = anthropic.Anthropic(
        api_key=api_key,
        timeout=httpx.Timeout(600.0, connect=5.0),
        max_retries=0,
    )
    sdk_errors = _sdk_error_classes(anthropic)
    chunks = _split_bundle_into_chunks(slim_bundle)
    chunk_results: List[dict] = [None] * len(chunks)

    def _classify_chunk(i: int, chunk: dict) -> dict:
        if len(chunks) > 1:
            print(f"  Processing chunk {i + 1}/{len(chunks)}...")

        user_message = _build_user_message(chunk, available_roles)
        max_regenerations = 2
        allowed_indices = {
            p.get("paragraph_index")
            for p in chunk.get("paragraphs", [])
            if isinstance(p, dict) and isinstance(p.get("paragraph_index"), int)
        }
        retry_error: Exception | None = None
        response_text = ""
        regeneration = 0
        transport_attempt = 0

        while True:
            try:
                # No sampling params (temperature/top_p/top_k): Sonnet 5 and
                # Opus 4.7+ reject non-default values with a 400.
                with _REQUEST_LIMITER:
                    with client.messages.stream(
                        model=model,
                        max_tokens=128000,
                        thinking={"type": "adaptive"},
                        output_config=_classification_output_config(available_roles),
                        system=PHASE2_MASTER_PROMPT.strip(),
                        messages=[{
                            "role": "user",
                            "content": (
                                user_message
                                if regeneration == 0
                                else user_message + _retry_requirement(
                                    retry_error or ValueError("prior response was not usable"),
                                    allowed_indices,
                                )
                            ),
                        }],
                    ) as stream:
                        response_text = stream.get_final_text()
                        stop_reason = _final_stop_reason(stream)
                if stop_reason == "refusal":
                    raise ClassificationRefused(
                        "LLM refused the target-classification request "
                        "(stop_reason=refusal)"
                    )
                if stop_reason == "max_tokens":
                    raise ValueError(
                        "LLM response reached the output-token limit before "
                        "completing its JSON (stop_reason=max_tokens)"
                    )
                parsed = _parse_classification_response(response_text)
                return _validate_classifications(parsed, available_roles, allowed_indices)
            except json.JSONDecodeError as e:
                # Malformed output: regenerate with the stricter instruction.
                retry_error = e
                if regeneration < max_regenerations:
                    print(f"  JSON parse error, retrying ({regeneration + 1}/{max_regenerations})...")
                    time.sleep(2 ** regeneration)
                    regeneration += 1
                else:
                    response_length = len(response_text.strip())
                    raise ValueError(
                        "Failed to parse LLM response as JSON after "
                        f"{max_regenerations + 1} attempts (last response: "
                        f"{response_length} characters): {e}"
                    )
            except ValueError as e:
                # Validation failure or a truncated response: regenerate with
                # the exact allowed indices restated.
                retry_error = e
                if regeneration < max_regenerations:
                    wait = 2 ** (regeneration + 1)
                    print(f"  Classification attempt failed: {e}, retrying in {wait}s ({regeneration + 1}/{max_regenerations})...")
                    time.sleep(wait)
                    regeneration += 1
                else:
                    raise RuntimeError(
                        f"LLM classification failed after {max_regenerations + 1} attempts: {e}"
                    )
            except ClassificationRefused:
                raise
            except Exception as e:
                # Transport failures: retry only what can heal. A bad key, a
                # bad request, or an unexpected error is re-raised at once.
                delay = _transport_retry_delay(e, transport_attempt, sdk_errors)
                if delay is None or transport_attempt >= _TRANSPORT_RETRIES:
                    raise
                print(
                    f"  Transient API failure: {e}; retrying in {delay:g}s "
                    f"({transport_attempt + 1}/{_TRANSPORT_RETRIES})..."
                )
                time.sleep(delay)
                transport_attempt += 1

    if len(chunks) == 1:
        chunk_results[0] = _classify_chunk(0, chunks[0])
    else:
        max_workers = min(len(chunks), 6)
        with ThreadPoolExecutor(max_workers=max_workers) as executor:
            futures = {
                executor.submit(_classify_chunk, i, chunk): i
                for i, chunk in enumerate(chunks)
            }
            for future in as_completed(futures):
                i = futures[future]
                chunk_results[i] = future.result()

    llm_only = chunk_results[0] if len(chunk_results) == 1 else _merge_chunk_results(chunk_results)
    result = coerce_to_final_classifications(slim_bundle, llm_only, available_roles)

    total_expected = (
        len(slim_bundle.get("paragraphs", []))
        + len(slim_bundle.get("deterministic_classifications", []))
        + len(slim_bundle.get("deterministic_ignored_paragraphs", []))
    )
    disposition_count = (
        len(result.get("classifications", []))
        + len(result.get("ignored_paragraphs", []))
    )
    if total_expected > 0 and disposition_count != total_expected:
        raise ValueError(
            f"Disposition coverage incomplete: {disposition_count}/{total_expected}. "
            "All classifiable paragraphs must be classified or explicitly ignored."
        )

    print(f"Disposition coverage: {disposition_count}/{total_expected} (100.0%)")
    return result
