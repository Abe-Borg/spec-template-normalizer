"""
LLM classifier module for automated CSI paragraph classification.

Calls the Anthropic API with the master prompt + slim bundle to produce
classification instructions (same schema as instructions.json).

Design constraint: pure module with no CLI — called by docx_decomposer.py
and gui.py.
"""
from __future__ import annotations

import json
import re
import time
from typing import Any, Dict, List, Optional


def estimate_tokens(text: str) -> int:
    """Rough token estimate (1 token ≈ 4 chars)."""
    return len(text) // 4


def _strip_code_fences(text: str) -> str:
    """Remove markdown code fences if present."""
    text = text.strip()
    if text.startswith("```"):
        # Remove opening fence (```json or ```)
        text = re.sub(r"^```\w*\s*\n?", "", text)
        # Remove closing fence
        text = re.sub(r"\n?```\s*$", "", text)
    return text.strip()


def _call_api(
    client: Any,
    system: str,
    user_message: str,
    model: str,
    max_tokens: int = 16384,
) -> str:
    """Single API call with retry logic. Returns raw response text."""
    import anthropic

    last_error: Optional[Exception] = None
    for attempt in range(3):  # initial + 2 retries
        try:
            response = client.messages.create(
                model=model,
                max_tokens=max_tokens,
                temperature=0,
                system=system,
                messages=[{"role": "user", "content": user_message}],
            )
            return response.content[0].text
        except (anthropic.APIError, anthropic.APIConnectionError, anthropic.RateLimitError) as e:
            last_error = e
            if attempt < 2:
                time.sleep(2 ** (attempt + 1))  # 2s, 4s
    raise last_error  # type: ignore[misc]


def _parse_response(raw: str) -> dict:
    """Parse JSON from LLM response, handling code fences."""
    cleaned = _strip_code_fences(raw)
    try:
        return json.loads(cleaned)
    except json.JSONDecodeError as e:
        raise ValueError(
            f"LLM response is not valid JSON: {e}\n\nRaw response (first 2000 chars):\n{raw[:2000]}"
        ) from e


def classify_document(
    slim_bundle: dict,
    master_prompt: str,
    run_instruction: str,
    api_key: str,
    model: str = "claude-sonnet-4-20250514",
) -> dict:
    """
    Classify all paragraphs in a slim bundle using the Anthropic API.

    Args:
        slim_bundle: The slim bundle dict from build_slim_bundle().
        master_prompt: Content of master_prompt.txt (system prompt).
        run_instruction: Content of run_instruction_prompt.txt (task prompt).
        api_key: Anthropic API key.
        model: Model ID to use.

    Returns:
        Parsed instructions dict (same schema as instructions.json).

    Raises:
        ValueError: If the LLM response is not valid JSON or fails validation.
    """
    import anthropic
    from docx_decomposer import validate_instructions

    client = anthropic.Anthropic(api_key=api_key)

    bundle_json = json.dumps(slim_bundle, indent=2)
    user_message = f"{run_instruction}\n\nSlim bundle:\n{bundle_json}"

    # Check if chunking is needed
    total_text = master_prompt + user_message
    token_est = estimate_tokens(total_text)

    if token_est > 80_000:
        return _classify_chunked(
            slim_bundle, master_prompt, run_instruction, client, model
        )

    raw = _call_api(client, master_prompt, user_message, model)
    instructions = _parse_response(raw)
    validate_instructions(instructions)
    return instructions


def _classify_chunked(
    slim_bundle: dict,
    master_prompt: str,
    run_instruction: str,
    client: Any,
    model: str,
    chunk_size: int = 200,
    overlap: int = 20,
) -> dict:
    """
    Chunked classification for large documents.

    First chunk returns full instructions (create_styles + roles + apply_pStyle).
    Subsequent chunks receive the already-determined styles/roles as context
    and return only apply_pStyle for their paragraph range.
    """
    from docx_decomposer import validate_instructions

    paragraphs = slim_bundle.get("paragraphs", [])
    total = len(paragraphs)

    # Build chunk boundaries
    chunks: List[tuple] = []  # (start, end) indices into paragraphs list
    start = 0
    while start < total:
        end = min(start + chunk_size, total)
        chunks.append((start, end))
        start = end - overlap if end < total else end

    # First chunk: full classification
    first_bundle = dict(slim_bundle)
    first_bundle["paragraphs"] = paragraphs[chunks[0][0]:chunks[0][1]]
    bundle_json = json.dumps(first_bundle, indent=2)
    user_msg = f"{run_instruction}\n\nSlim bundle:\n{bundle_json}"

    raw = _call_api(client, master_prompt, user_msg, model)
    merged = _parse_response(raw)
    validate_instructions(merged)

    if len(chunks) <= 1:
        return merged

    # Subsequent chunks: only apply_pStyle
    context_info = json.dumps({
        "create_styles": merged.get("create_styles", []),
        "roles": merged.get("roles", {}),
    }, indent=2)

    for chunk_start, chunk_end in chunks[1:]:
        chunk_bundle = dict(slim_bundle)
        chunk_bundle["paragraphs"] = paragraphs[chunk_start:chunk_end]
        chunk_json = json.dumps(chunk_bundle, indent=2)

        chunk_prompt = (
            f"{run_instruction}\n\n"
            f"The following styles and roles have already been determined:\n{context_info}\n\n"
            f"You MUST use these exact styles. Return ONLY the apply_pStyle array for the "
            f"paragraphs in this chunk (paragraph indices {chunk_start} to {chunk_end - 1}).\n"
            f"Output format: {{\"apply_pStyle\": [...]}}\n\n"
            f"Slim bundle chunk:\n{chunk_json}"
        )

        raw = _call_api(client, master_prompt, chunk_prompt, model)
        chunk_result = _parse_response(raw)
        chunk_apply = chunk_result.get("apply_pStyle", [])

        # Merge: deduplicate by paragraph_index (later chunk wins for overlaps)
        existing_indices = {item["paragraph_index"] for item in merged.get("apply_pStyle", [])}
        for item in chunk_apply:
            idx = item["paragraph_index"]
            if idx not in existing_indices:
                merged.setdefault("apply_pStyle", []).append(item)
                existing_indices.add(idx)

    # Sort apply_pStyle by paragraph_index
    merged["apply_pStyle"] = sorted(
        merged.get("apply_pStyle", []),
        key=lambda x: x["paragraph_index"],
    )

    validate_instructions(merged)
    return merged


def compute_coverage(slim_bundle: dict, instructions: dict) -> tuple:
    """
    Compute what percentage of classifiable paragraphs received a style.

    Returns:
        (coverage_fraction, styled_count, classifiable_count)
    """
    paragraphs = slim_bundle.get("paragraphs", [])
    classifiable = 0
    classifiable_indices = set()

    for p in paragraphs:
        text = (p.get("text") or "").strip()
        if not text:
            continue
        if p.get("contains_sectPr", False):
            continue
        if text.upper() == "END OF SECTION":
            continue
        # Editor/specifier notes in brackets
        if text.startswith("[") and text.endswith("]"):
            continue
        classifiable += 1
        classifiable_indices.add(p["paragraph_index"])

    styled_indices = {item["paragraph_index"] for item in instructions.get("apply_pStyle", [])}
    styled_count = len(styled_indices & classifiable_indices)

    coverage = styled_count / classifiable if classifiable > 0 else 1.0
    return coverage, styled_count, classifiable
