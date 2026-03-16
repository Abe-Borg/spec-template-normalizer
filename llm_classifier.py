"""
LLM classifier module for automated CSI paragraph classification.

Calls the Anthropic API with the master prompt + slim bundle to produce
classification instructions (same schema as instructions.json).

Design constraint: pure module with no CLI — imported by gui.py.
"""
from __future__ import annotations

import json
import re
import time
from typing import Any, Dict, List, Optional, Set, Tuple

from paragraph_rules import is_classifiable_paragraph


MAX_SINGLE_PASS_PARAGRAPHS = 500


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
    max_tokens: int = 128000,
) -> str:
    """Single API call with retry logic. Returns raw response text."""
    import anthropic

    last_error: Optional[Exception] = None
    for attempt in range(3):  # initial + 2 retries
        try:
            with client.messages.stream(
                model=model,
                max_tokens=max_tokens,
                temperature=1,
                thinking={"type": "adaptive"},
                output_config={"effort":"max"},
                system=system,
                messages=[{"role": "user", "content": user_message}],
            ) as stream:
                return stream.get_final_text()
        except (anthropic.APIError, anthropic.APIConnectionError, anthropic.RateLimitError) as e:
            last_error = e
            if attempt < 2:
                time.sleep(2 ** (attempt + 1))
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


def _extract_missing_indices(error: ValueError) -> Optional[List[int]]:
    """Parse missing paragraph indices from a coverage-mismatch ValueError.

    Expected format: 'apply_pStyle coverage mismatch; missing=[35, 59], unexpected=[]'
    Returns the list of missing indices, or None if this isn't a coverage error.
    """
    msg = str(error)
    if "apply_pStyle coverage mismatch" not in msg:
        return None

    m = re.search(r"missing=\[([^\]]*)\]", msg)
    if not m:
        return None

    raw = m.group(1).strip()
    if not raw:
        return None

    # Handle truncated lists (trailing '...')
    raw = raw.rstrip(".").rstrip(",").strip()
    try:
        return [int(x.strip()) for x in raw.split(",") if x.strip().isdigit()]
    except ValueError:
        return None


def _build_patch_prompt(
    slim_bundle: dict,
    instructions: dict,
    missing_indices: List[int],
) -> str:
    """Build a targeted prompt asking the LLM to classify only the missing paragraphs."""
    paragraphs = slim_bundle.get("paragraphs", [])

    # Collect the missing paragraphs with surrounding context
    context_paragraphs = []
    for idx in missing_indices:
        # Include 2 paragraphs before and after for context
        start = max(0, idx - 2)
        end = min(len(paragraphs), idx + 3)
        for i in range(start, end):
            if i not in [p["paragraph_index"] for p in context_paragraphs]:
                context_paragraphs.append(paragraphs[i])

    context_paragraphs.sort(key=lambda p: p["paragraph_index"])

    # Build the existing role→style mapping
    roles_info = json.dumps(instructions.get("roles", {}), indent=2)

    # Build the list of already-classified neighbors for context
    apply_map = {item["paragraph_index"]: item["styleId"]
                 for item in instructions.get("apply_pStyle", [])}

    neighbor_info = []
    for p in context_paragraphs:
        idx = p["paragraph_index"]
        if idx in apply_map:
            neighbor_info.append(
                f"  paragraph_index={idx} text={p.get('text', '')!r} → styleId={apply_map[idx]!r}"
            )

    prompt = (
        f"The following paragraphs were missed during initial classification.\n"
        f"The document's roles and styles have already been determined:\n{roles_info}\n\n"
        f"Already-classified neighbors for context:\n"
        + "\n".join(neighbor_info) + "\n\n"
        f"Classify ONLY the following paragraph indices: {missing_indices}\n"
        f"Use ONLY the styleIds from the roles above.\n\n"
        f"Context paragraphs (surrounding the missed ones):\n"
        f"{json.dumps(context_paragraphs, indent=2)}\n\n"
        f'Return ONLY a JSON object: {{"apply_pStyle": [{{"paragraph_index": N, "styleId": "..."}}]}}\n'
        f"No prose, no markdown, no extra keys."
    )
    return prompt


def classify_document(
    slim_bundle: dict,
    master_prompt: str,
    run_instruction: str,
    api_key: str,
    model: str = "claude-opus-4-6",
    max_patch_attempts: int = 2,
) -> dict:
    """
    Classify all paragraphs in a slim bundle using the Anthropic API.

    If the initial classification misses a small number of paragraphs
    (coverage mismatch), automatically makes a targeted follow-up call
    to classify just the missing indices before failing.

    Args:
        slim_bundle: The slim bundle dict from build_slim_bundle().
        master_prompt: Content of master_prompt.txt (system prompt).
        run_instruction: Content of run_instruction_prompt.txt (task prompt).
        api_key: Anthropic API key.
        model: Model ID to use.
        max_patch_attempts: Max number of targeted patch calls (default 2).

    Returns:
        Parsed instructions dict (same schema as instructions.json).

    Raises:
        ValueError: If the LLM response is not valid JSON or fails validation
                    after all patch attempts are exhausted.
    """
    paragraphs = slim_bundle.get("paragraphs", [])
    if len(paragraphs) > MAX_SINGLE_PASS_PARAGRAPHS:
        raise ValueError(
            "Document too large for current single-pass classification; chunked mode requires redesign"
        )

    import anthropic
    from docx_decomposer import validate_instructions

    client = anthropic.Anthropic(api_key=api_key)

    bundle_json = json.dumps(slim_bundle, indent=2)
    user_message = f"{run_instruction}\n\nSlim bundle:\n{bundle_json}"

    raw = _call_api(client, master_prompt, user_message, model)
    instructions = _parse_response(raw)

    # Attempt validation; if coverage mismatch, try targeted patching
    for patch_attempt in range(max_patch_attempts + 1):
        try:
            validate_instructions(instructions, slim_bundle=slim_bundle)
            return instructions  # Clean pass
        except ValueError as exc:
            missing = _extract_missing_indices(exc)
            if missing is None or patch_attempt >= max_patch_attempts:
                raise  # Not a coverage error, or out of patch attempts

            print(
                f"Coverage patch attempt {patch_attempt + 1}/{max_patch_attempts}: "
                f"classifying {len(missing)} missing paragraph(s): {missing}"
            )

            patch_prompt = _build_patch_prompt(slim_bundle, instructions, missing)
            patch_raw = _call_api(client, master_prompt, patch_prompt, model)
            patch_result = _parse_response(patch_raw)

            # Merge patch results into instructions
            existing_indices = {
                item["paragraph_index"]
                for item in instructions.get("apply_pStyle", [])
            }
            for item in patch_result.get("apply_pStyle", []):
                idx = item["paragraph_index"]
                if idx not in existing_indices:
                    instructions.setdefault("apply_pStyle", []).append(item)
                    existing_indices.add(idx)

            # Re-sort for deterministic output
            instructions["apply_pStyle"] = sorted(
                instructions.get("apply_pStyle", []),
                key=lambda x: x["paragraph_index"],
            )

    # Final validation (should not reach here, but safety net)
    validate_instructions(instructions, slim_bundle=slim_bundle)
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
    validate_instructions(merged, slim_bundle=slim_bundle)

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

    validate_instructions(merged, slim_bundle=slim_bundle)
    return merged


def compute_coverage(slim_bundle: dict, instructions: dict) -> tuple:
    """
    Compute what percentage of classifiable paragraphs received a style.

    Returns:
        (coverage_fraction, styled_count, classifiable_count)
    """
    paragraphs = slim_bundle.get("paragraphs", [])
    classifiable_indices = {p["paragraph_index"] for p in paragraphs if is_classifiable_paragraph(p)}
    classifiable = len(classifiable_indices)

    styled_indices = {item["paragraph_index"] for item in instructions.get("apply_pStyle", [])}
    styled_count = len(styled_indices & classifiable_indices)

    coverage = styled_count / classifiable if classifiable > 0 else 1.0
    return coverage, styled_count, classifiable
