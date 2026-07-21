"""
LLM-based classification for Phase 2.

Sends paragraph bundles to the Anthropic API for CSI role classification,
with retry logic, chunking for large documents, and coverage reporting.
"""

import json
import time
import re
from concurrent.futures import ThreadPoolExecutor, as_completed
from typing import Dict, Any, List, Set

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


def _build_user_message(slim_bundle: dict, available_roles: list) -> str:
    return (
        PHASE2_RUN_INSTRUCTION.strip()
        + "\n\navailable_roles: " + json.dumps(available_roles)
        + "\n\n" + json.dumps(slim_bundle, indent=2)
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
                    "notes": {
                        "type": "array",
                        "items": {"type": "string"},
                    },
                },
                "required": ["classifications", "notes"],
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

    missing = sorted(allowed_indices - seen_indices)
    if missing:
        raise ValueError(f"missing coverage for paragraph indices: {missing[:20]}")

    return {"classifications": validated, "notes": classifications.get("notes", [])}


def _split_bundle_into_chunks(slim_bundle: dict, max_chars: int = _MAX_BUNDLE_CHARS) -> List[dict]:
    paragraphs = slim_bundle.get("paragraphs", [])
    roles = slim_bundle.get("available_roles", [])
    filter_report = slim_bundle.get("filter_report", {})

    full_json = json.dumps(slim_bundle)
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
    seen: Dict[int, str] = {}
    conflicts: List[Dict[str, Any]] = []
    all_notes: List[Any] = []

    for result in chunk_results:
        for item in result.get("classifications", []):
            idx = item.get("paragraph_index")
            role = item.get("csi_role")
            if idx is None or role is None:
                continue
            prior = seen.get(idx)
            if prior is not None and prior != role:
                conflicts.append({"paragraph_index": idx, "existing_role": prior, "conflicting_role": role})
            seen[idx] = role
        all_notes.extend(result.get("notes", []))

    if conflicts:
        examples = ", ".join([f"{c['paragraph_index']}:{c['existing_role']}|{c['conflicting_role']}" for c in conflicts[:10]])
        raise ValueError(f"Chunk merge conflicts detected ({len(conflicts)} total): {examples}")

    return {
        "classifications": [{"paragraph_index": idx, "csi_role": role} for idx, role in sorted(seen.items())],
        "notes": all_notes,
    }


def classify_target_document(slim_bundle: dict, available_roles: list, api_key: str, model: str = "claude-sonnet-5") -> dict:
    unresolved_paragraphs = slim_bundle.get("paragraphs", [])
    if not unresolved_paragraphs:
        deterministic_only = coerce_to_final_classifications(
            slim_bundle,
            {"classifications": [], "notes": ["LLM skipped: all paragraphs classified deterministically."]},
            available_roles,
        )
        total_expected = len(deterministic_only.get("classifications", []))
        print("LLM skipped: all paragraphs classified deterministically.")
        print(f"Classification coverage: {total_expected}/{total_expected} (100.0%)")
        return deterministic_only

    import anthropic

    client = anthropic.Anthropic(api_key=api_key)
    chunks = _split_bundle_into_chunks(slim_bundle)
    chunk_results: List[dict] = [None] * len(chunks)

    def _classify_chunk(i: int, chunk: dict) -> dict:
        if len(chunks) > 1:
            print(f"  Processing chunk {i + 1}/{len(chunks)}...")

        user_message = _build_user_message(chunk, available_roles)
        max_retries = 2

        for attempt in range(max_retries + 1):
            try:
                # No sampling params (temperature/top_p/top_k): Sonnet 5 and
                # Opus 4.7+ reject non-default values with a 400.
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
                            if attempt == 0
                            else user_message
                            + "\n\nRETRY REQUIREMENT: The prior response was not usable. "
                            "Return exactly one complete JSON object with no prose, "
                            "Markdown fence, or text before or after it."
                        ),
                    }],
                ) as stream:
                    response_text = stream.get_final_text()
                parsed = _parse_classification_response(response_text)
                allowed_indices = {
                    p.get("paragraph_index")
                    for p in chunk.get("paragraphs", [])
                    if isinstance(p, dict) and isinstance(p.get("paragraph_index"), int)
                }
                return _validate_classifications(parsed, available_roles, allowed_indices)
            except json.JSONDecodeError as e:
                if attempt < max_retries:
                    print(f"  JSON parse error, retrying ({attempt + 1}/{max_retries})...")
                    time.sleep(2 ** attempt)
                else:
                    response_length = len(response_text.strip())
                    raise ValueError(
                        "Failed to parse LLM response as JSON after "
                        f"{max_retries + 1} attempts (last response: "
                        f"{response_length} characters): {e}"
                    )
            except Exception as e:
                if attempt < max_retries:
                    wait = 2 ** (attempt + 1)
                    print(f"  API error: {e}, retrying in {wait}s ({attempt + 1}/{max_retries})...")
                    time.sleep(wait)
                else:
                    raise RuntimeError(f"LLM classification failed after {max_retries + 1} attempts: {e}")

        raise RuntimeError("Unexpected chunk classification exit without result")

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

    total_expected = len(slim_bundle.get("paragraphs", [])) + len(slim_bundle.get("deterministic_classifications", []))
    classified_count = len(result.get("classifications", []))
    if total_expected > 0 and classified_count != total_expected:
        raise ValueError(
            f"Classification coverage incomplete: {classified_count}/{total_expected}. "
            "All classifiable paragraphs must be classified."
        )

    print(f"Classification coverage: {classified_count}/{total_expected} (100.0%)")
    return result
