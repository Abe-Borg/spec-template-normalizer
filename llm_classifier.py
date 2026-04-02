"""
LLM classifier module for automated CSI paragraph classification.

Calls the Anthropic API with the master prompt + slim bundle to produce
classification instructions (same schema as instructions.json).

Design constraint: pure module with no CLI — imported by gui.py.
"""
from __future__ import annotations

import bisect
import json
import re
import time
from typing import Any, Dict, List, Optional, Set, Tuple

from paragraph_rules import (
    is_classifiable_paragraph,
    detect_role_signal,
    infer_expected_roles,
    RE_ALPHA_PARA,
    RE_NUMERIC_SUB,
)


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


def _repair_strong_signal_mismatches(
    instructions: dict,
    slim_bundle: dict,
) -> int:
    """Auto-correct apply_pStyle entries where strong text signals contradict
    the assigned styleId.

    When a paragraph's text unambiguously matches a CSI role pattern
    (e.g. ``SECTION 23 05 00`` → SectionID) but the LLM assigned a
    different styleId than what it declared in the ``roles`` mapping,
    this function overwrites the apply_pStyle entry with the correct
    styleId.

    This avoids an unnecessary retry/patch API call for errors that are
    deterministically fixable from regex evidence alone.

    Returns:
        Number of corrections applied.
    """
    roles = instructions.get("roles")
    if not roles or not isinstance(roles, dict):
        return 0

    # Build role → styleId mapping
    role_style: Dict[str, str] = {}
    for role, spec in roles.items():
        if isinstance(spec, dict) and spec.get("styleId"):
            role_style[role] = spec["styleId"]

    if not role_style:
        return 0

    # Build paragraph_index → apply_pStyle entry mapping (by reference)
    apply_list = instructions.get("apply_pStyle", [])
    apply_map: Dict[int, dict] = {item["paragraph_index"]: item for item in apply_list}

    paragraphs = slim_bundle.get("paragraphs", [])
    classifiable = [p for p in paragraphs if is_classifiable_paragraph(p)]

    # Determine whether numeric/lower patterns are strong signals
    has_alpha = any(
        RE_ALPHA_PARA.match((p.get("text") or "").strip()) for p in classifiable
    )
    has_numeric = any(
        RE_NUMERIC_SUB.match((p.get("text") or "").strip()) for p in classifiable
    )

    corrections = 0
    for p in classifiable:
        idx = p["paragraph_index"]
        text = (p.get("text") or "").strip()

        signal = detect_role_signal(
            text, numeric_is_strong=has_alpha, lower_is_strong=has_numeric
        )
        if signal is None:
            continue

        expected_style = role_style.get(signal)
        if not expected_style:
            continue

        entry = apply_map.get(idx)
        if entry is not None and entry["styleId"] != expected_style:
            entry["styleId"] = expected_style
            corrections += 1

    return corrections


def _repair_missing_roles(
    instructions: dict,
    slim_bundle: dict,
) -> int:
    """Auto-add missing roles when strong text signals prove they exist in
    the document but the LLM omitted them from ``roles``.

    For each missing role detected by :func:`infer_expected_roles`:

    1. Select the first strong-hit paragraph as the exemplar.
    2. Determine the styleId — if all strong-hit paragraphs already share
       one assigned style, reuse it; otherwise create a ``CSI_*__ARCH``
       style derived from the exemplar.
    3. Insert the role into ``roles``, the style into ``create_styles``
       (if new), and ensure all strong-hit paragraphs carry the correct
       styleId in ``apply_pStyle``.

    Returns:
        Number of roles added.
    """
    paragraphs = slim_bundle.get("paragraphs", [])
    expected_roles, strong_hits = infer_expected_roles(paragraphs)

    roles = instructions.get("roles")
    if not isinstance(roles, dict):
        instructions["roles"] = roles = {}

    missing = sorted(expected_roles - set(roles.keys()))
    if not missing:
        return 0

    # Mapping: role name → canonical CSI_*__ARCH styleId
    ROLE_TO_ARCH_STYLE = {
        "SectionID": "CSI_SectionID__ARCH",
        "SectionTitle": "CSI_SectionTitle__ARCH",
        "PART": "CSI_Part__ARCH",
        "ARTICLE": "CSI_Article__ARCH",
        "PARAGRAPH": "CSI_Paragraph__ARCH",
        "SUBPARAGRAPH": "CSI_Subparagraph__ARCH",
        "SUBSUBPARAGRAPH": "CSI_Subsubparagraph__ARCH",
    }

    ROLE_TO_STYLE_NAME = {
        "SectionID": "CSI SectionID (Architect Template)",
        "SectionTitle": "CSI SectionTitle (Architect Template)",
        "PART": "CSI Part (Architect Template)",
        "ARTICLE": "CSI Article (Architect Template)",
        "PARAGRAPH": "CSI Paragraph (Architect Template)",
        "SUBPARAGRAPH": "CSI Subparagraph (Architect Template)",
        "SUBSUBPARAGRAPH": "CSI Subsubparagraph (Architect Template)",
    }

    # Build working indices
    apply_list = instructions.get("apply_pStyle", [])
    apply_map: Dict[int, dict] = {
        item["paragraph_index"]: item for item in apply_list
    }
    existing_style_ids = set(slim_bundle.get("style_catalog", {}).keys())
    created_style_ids = {
        sd["styleId"]
        for sd in instructions.get("create_styles", [])
        if isinstance(sd, dict)
    }

    added = 0
    for role in missing:
        hits = strong_hits.get(role, [])
        if not hits:
            # SectionTitle has no regex hits — skip for now (position-based,
            # requires more context than a simple auto-repair can provide).
            continue

        exemplar_idx = hits[0]

        # Check if these paragraphs already share a single assigned style
        assigned_styles: Set[str] = set()
        for idx in hits:
            entry = apply_map.get(idx)
            if entry:
                assigned_styles.add(entry["styleId"])

        if len(assigned_styles) == 1:
            # All strong-hit paragraphs already have the same style — reuse it
            style_id = assigned_styles.pop()
        else:
            # Need a CSI_*__ARCH style
            style_id = ROLE_TO_ARCH_STYLE.get(role)
            if not style_id:
                continue

            # Create the style definition if it doesn't already exist
            if style_id not in existing_style_ids and style_id not in created_style_ids:
                create_list = instructions.setdefault("create_styles", [])
                create_list.append({
                    "styleId": style_id,
                    "name": ROLE_TO_STYLE_NAME.get(role, style_id),
                    "type": "paragraph",
                    "derive_from_paragraph_index": exemplar_idx,
                    "basedOn": "Normal",
                })
                created_style_ids.add(style_id)

        # Register the role
        roles[role] = {
            "styleId": style_id,
            "exemplar_paragraph_index": exemplar_idx,
        }

        # Ensure every strong-hit paragraph is assigned this style
        for idx in hits:
            entry = apply_map.get(idx)
            if entry is None:
                new_entry = {"paragraph_index": idx, "styleId": style_id}
                instructions.setdefault("apply_pStyle", []).append(new_entry)
                apply_map[idx] = new_entry
            elif entry["styleId"] != style_id:
                entry["styleId"] = style_id

        added += 1

    # Keep apply_pStyle sorted after mutations
    if added:
        instructions["apply_pStyle"] = sorted(
            instructions.get("apply_pStyle", []),
            key=lambda x: x["paragraph_index"],
        )

    return added


def _repair_role_exemplar_mismatches(
    instructions: dict,
    slim_bundle: dict,
) -> int:
    """Repair role exemplars whose paragraph text strongly signals a different role."""
    roles = instructions.get("roles")
    if not isinstance(roles, dict):
        return 0

    paragraphs = slim_bundle.get("paragraphs", [])
    classifiable = [p for p in paragraphs if is_classifiable_paragraph(p)]
    by_index: Dict[int, dict] = {p["paragraph_index"]: p for p in classifiable}
    if not by_index:
        return 0

    _, strong_hits = infer_expected_roles(paragraphs)

    has_alpha = any(
        RE_ALPHA_PARA.match((p.get("text") or "").strip()) for p in classifiable
    )
    has_numeric = any(
        RE_NUMERIC_SUB.match((p.get("text") or "").strip()) for p in classifiable
    )

    def _find_correct_exemplar(role: str) -> Optional[int]:
        if role == "SectionTitle":
            for section_id_idx in sorted(strong_hits.get("SectionID", [])):
                candidate_idx = section_id_idx + 1
                para = by_index.get(candidate_idx)
                if not para:
                    continue
                text = (para.get("text") or "").strip()
                signal = detect_role_signal(
                    text, numeric_is_strong=has_alpha, lower_is_strong=has_numeric
                )
                if signal is None:
                    return candidate_idx
            return None

        hits = strong_hits.get(role, [])
        return hits[0] if hits else None

    corrections = 0
    for role, spec in list(roles.items()):
        if not isinstance(spec, dict):
            continue

        exemplar_idx = spec.get("exemplar_paragraph_index")
        if not isinstance(exemplar_idx, int):
            continue

        para = by_index.get(exemplar_idx)
        if not para:
            continue

        text = (para.get("text") or "").strip()
        signal = detect_role_signal(
            text, numeric_is_strong=has_alpha, lower_is_strong=has_numeric
        )

        if signal is None or signal == role:
            continue

        correct_idx = _find_correct_exemplar(role)
        if correct_idx is None:
            # No valid exemplar exists for this role in the document.
            del roles[role]
            corrections += 1
            print(f"Removed phantom role '{role}' (no valid exemplar exists)")
            continue
        if correct_idx == exemplar_idx:
            continue

        spec["exemplar_paragraph_index"] = correct_idx
        style_id = spec.get("styleId")
        for sd in instructions.get("create_styles", []):
            if not isinstance(sd, dict):
                continue
            if sd.get("styleId") == style_id and sd.get("derive_from_paragraph_index") == exemplar_idx:
                sd["derive_from_paragraph_index"] = correct_idx

        corrections += 1

    return corrections


def _repair_coverage_gaps(
    instructions: dict,
    slim_bundle: dict,
) -> int:
    """Fill remaining unclassified classifiable paragraphs by nearest-neighbor style."""
    paragraphs = slim_bundle.get("paragraphs", [])
    classifiable_indices = {
        p["paragraph_index"] for p in paragraphs if is_classifiable_paragraph(p)
    }
    apply_list = instructions.setdefault("apply_pStyle", [])
    apply_map: Dict[int, str] = {
        item["paragraph_index"]: item["styleId"] for item in apply_list if isinstance(item, dict)
    }

    missing = sorted(classifiable_indices - set(apply_map.keys()))
    if not missing:
        return 0

    classified_indices = sorted(apply_map.keys())
    repairs = 0

    for idx in missing:
        insert_at = bisect.bisect_left(classified_indices, idx)

        prev_style: Optional[str] = None
        if insert_at > 0:
            prev_style = apply_map.get(classified_indices[insert_at - 1])

        next_style: Optional[str] = None
        if insert_at < len(classified_indices):
            next_style = apply_map.get(classified_indices[insert_at])

        if prev_style and next_style:
            chosen = prev_style if prev_style == next_style else prev_style
        else:
            chosen = prev_style or next_style

        if not chosen:
            continue

        apply_list.append({"paragraph_index": idx, "styleId": chosen})
        apply_map[idx] = chosen
        bisect.insort(classified_indices, idx)
        repairs += 1

    if repairs:
        instructions["apply_pStyle"] = sorted(
            apply_list,
            key=lambda x: x["paragraph_index"],
        )

    return repairs


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
    max_patch_attempts: int = 3,
) -> dict:
    """
    Classify all paragraphs in a slim bundle using the Anthropic API.

    If the initial classification misses a small number of paragraphs
    (coverage mismatch), automatically makes targeted follow-up calls
    to classify just the missing indices. If coverage still fails after
    all patch attempts, applies a nearest-neighbor fallback repair.

    Deterministic repairs are applied locally before validation:
    missing roles, role-exemplar mismatches, and strong-signal style
    mismatches (where regex-detectable text patterns like
    "SECTION 23 05 00" or "PART 1 - GENERAL" contradict the LLM's style
    assignment).

    Args:
        slim_bundle: The slim bundle dict from build_slim_bundle().
        master_prompt: Content of master_prompt.txt (system prompt).
        run_instruction: Content of run_instruction_prompt.txt (task prompt).
        api_key: Anthropic API key.
        model: Model ID to use.
        max_patch_attempts: Max number of targeted patch calls (default 3).

    Returns:
        Parsed instructions dict (same schema as instructions.json).

    Raises:
        ValueError: If the LLM response is not valid JSON or fails validation
                    after all patch attempts are exhausted.
    """
    paragraphs = slim_bundle.get("paragraphs", [])
    if len(paragraphs) > MAX_SINGLE_PASS_PARAGRAPHS:
        raise ValueError(
            f"Unsupported document size: {len(paragraphs)} paragraphs exceeds "
            f"single-pass limit ({MAX_SINGLE_PASS_PARAGRAPHS}). Chunked classification is intentionally disabled "
            "pending two-pass redesign."
        )

    import anthropic
    from docx_decomposer import validate_instructions

    client = anthropic.Anthropic(api_key=api_key)

    bundle_json = json.dumps(slim_bundle, indent=2)
    user_message = f"{run_instruction}\n\nSlim bundle:\n{bundle_json}"

    raw = _call_api(client, master_prompt, user_message, model)
    instructions = _parse_response(raw)

    # Auto-repair missing roles before other repairs.
    # When the LLM omits a role that strong text signals prove exists
    # (e.g. SECTION XX XX XX paragraphs exist but roles has no SectionID),
    # add the role, create the style if needed, and assign strong-hit paragraphs.
    added_roles = _repair_missing_roles(instructions, slim_bundle)
    if added_roles:
        print(f"Auto-added {added_roles} missing role(s) from strong text signals")

    exemplar_repairs = _repair_role_exemplar_mismatches(instructions, slim_bundle)
    if exemplar_repairs:
        print(f"Auto-repaired {exemplar_repairs} role exemplar mismatch(es)")

    # Auto-repair strong-signal mismatches before validation.
    # These are paragraphs whose text unambiguously identifies their CSI role
    # (via regex) but were assigned the wrong styleId by the LLM.
    repairs = _repair_strong_signal_mismatches(instructions, slim_bundle)
    if repairs:
        print(f"Auto-repaired {repairs} strong-signal style mismatch(es)")

    # Attempt validation; if coverage mismatch, try targeted patching
    for patch_attempt in range(max_patch_attempts + 1):
        try:
            validate_instructions(instructions, slim_bundle=slim_bundle)
            return instructions  # Clean pass
        except ValueError as exc:
            missing = _extract_missing_indices(exc)
            if missing is None:
                raise  # Not a coverage error

            if patch_attempt >= max_patch_attempts:
                gap_fills = _repair_coverage_gaps(instructions, slim_bundle)
                if gap_fills:
                    print(f"Gap-filled {gap_fills} paragraph(s) via nearest-neighbor")
                    break
                raise  # Out of patch attempts and no deterministic gap fill was possible

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

            # Re-run strong-signal repair after patching (new entries may need it)
            repairs = _repair_strong_signal_mismatches(instructions, slim_bundle)
            if repairs:
                print(f"Auto-repaired {repairs} strong-signal style mismatch(es) after patch")

    # Final validation (should not reach here, but safety net)
    validate_instructions(instructions, slim_bundle=slim_bundle)
    return instructions


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
