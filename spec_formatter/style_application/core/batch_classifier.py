"""Batch API helpers for Phase 2 LLM classification."""

from __future__ import annotations

import json
import time
from typing import Callable, Dict, List

from .classification import PHASE2_MASTER_PROMPT, coerce_to_final_classifications
from .llm_classifier import (
    _build_user_message,
    _classification_output_config,
    _merge_chunk_results,
    _parse_classification_response,
    _split_bundle_into_chunks,
    _validate_classifications,
)

_CHUNK_DELIMITER = "__chunk"
_MAX_BATCH_REQUESTS = 10_000


class BatchClassificationError(RuntimeError):
    """Raised when one or more files cannot be classified in batch mode."""


def _make_custom_id(filename: str, chunk_index: int) -> str:
    return f"{filename}{_CHUNK_DELIMITER}{chunk_index}"


def _parse_custom_id(custom_id: str) -> tuple[str, int]:
    if _CHUNK_DELIMITER not in custom_id:
        raise ValueError(f"Invalid custom_id '{custom_id}'")
    stem, chunk_part = custom_id.rsplit(_CHUNK_DELIMITER, 1)
    return stem, int(chunk_part)


def build_batch_requests(
    file_bundles: Dict[str, dict],
    available_roles: List[str],
    model: str,
) -> List[dict]:
    requests: List[dict] = []
    for filename in sorted(file_bundles.keys()):
        slim_bundle = file_bundles[filename]
        if not slim_bundle.get("paragraphs"):
            # Deterministic-only files are reassembled locally.  Do not create
            # an empty LLM chunk (or require credentials) for them.
            continue
        chunks = _split_bundle_into_chunks(slim_bundle)
        for i, chunk in enumerate(chunks):
            user_message = _build_user_message(chunk, available_roles)
            requests.append(
                {
                    "custom_id": _make_custom_id(filename, i),
                    "params": {
                        "model": model,
                        "max_tokens": 128000,
                        "thinking": {"type": "adaptive"},
                        "output_config": _classification_output_config(
                            available_roles
                        ),
                        "system": PHASE2_MASTER_PROMPT.strip(),
                        "messages": [{"role": "user", "content": user_message}],
                    },
                }
            )

    if len(requests) > _MAX_BATCH_REQUESTS:
        raise ValueError(
            f"Batch request count {len(requests)} exceeds API limit of {_MAX_BATCH_REQUESTS}. "
            "Split into multiple runs."
        )

    return requests


def _extract_text_from_blocks(content_blocks: list) -> str:
    parts: List[str] = []
    for block in content_blocks:
        if getattr(block, "type", None) == "text":
            parts.append(getattr(block, "text", ""))
    return "".join(parts)


def submit_and_poll(
    requests: List[dict],
    api_key: str,
    poll_interval: int = 30,
    on_poll: Callable | None = None,
) -> Dict[str, dict]:
    if not requests:
        return {}

    import anthropic

    client = anthropic.Anthropic(api_key=api_key)
    batch = client.messages.batches.create(requests=requests)
    batch_id = batch.id

    while True:
        batch = client.messages.batches.retrieve(batch_id)
        if on_poll:
            on_poll(batch_id, batch.processing_status, batch.request_counts)
        if batch.processing_status == "ended":
            break
        if batch.processing_status == "canceling":
            time.sleep(max(1, poll_interval))
            continue
        time.sleep(max(1, poll_interval))

    parsed_results: Dict[str, dict] = {}
    failed_entries: Dict[str, str] = {}

    for entry in client.messages.batches.results(batch_id):
        custom_id = entry.custom_id
        result_type = entry.result.type

        if result_type != "succeeded":
            failed_entries[custom_id] = result_type
            continue

        text = _extract_text_from_blocks(entry.result.message.content)
        parsed = _parse_classification_response(text)
        parsed_results[custom_id] = parsed

    if failed_entries:
        details = ", ".join(f"{cid}: {status}" for cid, status in sorted(failed_entries.items()))
        raise BatchClassificationError(f"Batch completed with failed requests: {details}")

    return parsed_results


def reassemble_file_classifications(
    results: Dict[str, dict],
    file_bundles: Dict[str, dict],
    available_roles: List[str],
) -> Dict[str, dict]:
    chunk_map: Dict[str, List[tuple[int, dict]]] = {}
    for custom_id, parsed in results.items():
        file_stem, chunk_index = _parse_custom_id(custom_id)
        chunk_map.setdefault(file_stem, []).append((chunk_index, parsed))

    output: Dict[str, dict] = {}
    failed_files: Dict[str, str] = {}

    for filename, slim_bundle in file_bundles.items():
        unresolved_paragraphs = slim_bundle.get("paragraphs", [])
        if not unresolved_paragraphs:
            output[filename] = coerce_to_final_classifications(
                slim_bundle,
                {"classifications": [], "notes": ["LLM skipped: all paragraphs classified deterministically."]},
                available_roles,
            )
            continue

        expected_chunks = len(_split_bundle_into_chunks(slim_bundle))
        entries = sorted(chunk_map.get(filename, []), key=lambda item: item[0])

        if len(entries) != expected_chunks:
            failed_files[filename] = f"expected {expected_chunks} chunks, got {len(entries)}"
            continue

        validated_chunks: List[dict] = []
        split_chunks = _split_bundle_into_chunks(slim_bundle)
        for chunk_index, parsed in entries:
            chunk = split_chunks[chunk_index]
            allowed_indices = {
                p.get("paragraph_index")
                for p in chunk.get("paragraphs", [])
                if isinstance(p, dict) and isinstance(p.get("paragraph_index"), int)
            }
            validated = _validate_classifications(parsed, available_roles, allowed_indices)
            validated_chunks.append(validated)

        llm_only = validated_chunks[0] if len(validated_chunks) == 1 else _merge_chunk_results(validated_chunks)
        final = coerce_to_final_classifications(slim_bundle, llm_only, available_roles)

        total_expected = len(slim_bundle.get("paragraphs", [])) + len(slim_bundle.get("deterministic_classifications", []))
        classified_count = len(final.get("classifications", []))
        if total_expected > 0 and classified_count != total_expected:
            failed_files[filename] = f"coverage incomplete ({classified_count}/{total_expected})"
            continue

        output[filename] = final

    if failed_files:
        details = ", ".join(f"{name}: {reason}" for name, reason in sorted(failed_files.items()))
        raise BatchClassificationError(f"Failed to assemble classifications for files: {details}")

    return output
