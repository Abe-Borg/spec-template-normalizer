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

from paragraph_rules import (
    is_classifiable_paragraph,
    infer_expected_roles,
)


MAX_SINGLE_PASS_INPUT_TOKENS = 150_000


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
        except (anthropic.APIConnectionError, anthropic.RateLimitError) as e:
            last_error = e
            if attempt < 2:
                time.sleep(2 ** (attempt + 1))
        except anthropic.APIStatusError as e:
            # Invalid credentials/requests will not heal on retry. Retry only
            # transient server failures.
            if getattr(e, "status_code", 0) < 500:
                raise
            last_error = e
            if attempt < 2:
                time.sleep(2 ** (attempt + 1))
    raise last_error  # type: ignore[misc]


def _parse_response(raw: str) -> dict:
    """Parse JSON from LLM response, handling code fences."""
    cleaned = _strip_code_fences(raw)

    def reject_duplicate_keys(pairs: List[Tuple[str, Any]]) -> Dict[str, Any]:
        result: Dict[str, Any] = {}
        for key, value in pairs:
            if key in result:
                raise ValueError(f"LLM response contains duplicate JSON key: {key!r}")
            result[key] = value
        return result

    try:
        parsed = json.loads(
            cleaned,
            object_pairs_hook=reject_duplicate_keys,
            parse_constant=lambda value: (_ for _ in ()).throw(
                ValueError(f"LLM response contains non-finite JSON number: {value}")
            ),
        )
    except (json.JSONDecodeError, ValueError) as e:
        raise ValueError(
            f"LLM response is not valid JSON: {e}\n\nRaw response (first 2000 chars):\n{raw[:2000]}"
        ) from e
    if not isinstance(parsed, dict):
        raise ValueError("LLM response must be a JSON object")
    return parsed


def _normalize_known_exclusions(instructions: dict, slim_bundle: dict) -> int:
    """Move deterministic editorial/copyright exclusions into the audit list."""
    known_reasons = {
        "editor_note": "Detected editor/specifier note",
        "specifier_note": "Detected publisher editing instruction",
        "copyright_notice": "Detected copyright/distribution notice",
    }
    excluded = {
        int(paragraph["paragraph_index"]): known_reasons[paragraph["skip_reason"]]
        for paragraph in slim_bundle.get("paragraphs", [])
        if paragraph.get("skip_reason") in known_reasons
    }
    if not excluded:
        return 0

    original_apply = instructions.get("apply_pStyle", [])
    instructions["apply_pStyle"] = [
        item for item in original_apply if item.get("paragraph_index") not in excluded
    ]
    ignored = {
        item["paragraph_index"]: item
        for item in instructions.get("ignored_paragraphs", [])
        if isinstance(item, dict) and isinstance(item.get("paragraph_index"), int)
    }
    for index, reason in excluded.items():
        ignored.setdefault(index, {"paragraph_index": index, "reason": reason})
    instructions["ignored_paragraphs"] = sorted(
        ignored.values(), key=lambda item: item["paragraph_index"]
    )
    return len(excluded)


def _extract_missing_indices(error: ValueError) -> Optional[List[int]]:
    """Parse missing paragraph indices from a coverage-mismatch ValueError.

    Expected format: 'classification coverage mismatch; missing=[35, 59], unexpected=[]'
    Returns the list of missing indices, or None if this isn't a coverage error.
    """
    msg = str(error)
    if "classification coverage mismatch" not in msg:
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

    _expected, strong_hits = infer_expected_roles(
        paragraphs,
        numbering_catalog=slim_bundle.get("numbering_catalog"),
    )
    signal_by_index = {
        int(index): role
        for role, indices in strong_hits.items()
        for index in indices
    }

    corrections = 0
    for p in classifiable:
        idx = p["paragraph_index"]
        signal = signal_by_index.get(int(idx))
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
    expected_roles, strong_hits = infer_expected_roles(
        paragraphs,
        numbering_catalog=slim_bundle.get("numbering_catalog"),
    )

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
        "END_OF_SECTION": "CSI_EndOfSection__ARCH",
    }

    ROLE_TO_STYLE_NAME = {
        "SectionID": "CSI SectionID (Architect Template)",
        "SectionTitle": "CSI SectionTitle (Architect Template)",
        "PART": "CSI Part (Architect Template)",
        "ARTICLE": "CSI Article (Architect Template)",
        "PARAGRAPH": "CSI Paragraph (Architect Template)",
        "SUBPARAGRAPH": "CSI Subparagraph (Architect Template)",
        "SUBSUBPARAGRAPH": "CSI Subsubparagraph (Architect Template)",
        "END_OF_SECTION": "CSI End of Section (Architect Template)",
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

        # A combined "SECTION 01 23 45 - TITLE" paragraph carries both
        # SectionID and SectionTitle semantics and can only have one pStyle.
        if role == "SectionTitle":
            section_id_spec = roles.get("SectionID")
            if (
                isinstance(section_id_spec, dict)
                and int(section_id_spec.get("exemplar_paragraph_index", -1)) == exemplar_idx
                and section_id_spec.get("styleId")
            ):
                style_id = section_id_spec["styleId"]
                roles[role] = {
                    "styleId": style_id,
                    "exemplar_paragraph_index": exemplar_idx,
                }
                added += 1
                continue

        hit_paragraphs = [paragraphs[idx] for idx in hits]
        source_styles = {p.get("pStyle") for p in hit_paragraphs}
        source_style = hit_paragraphs[0].get("pStyle")
        has_direct_paragraph_formatting = any(
            p.get("pPr_hints")
            or p.get("rPr_hints")
            or p.get("numPr")
            or p.get("has_direct_pPr")
            or p.get("has_uniform_direct_rPr")
            for p in hit_paragraphs
        )

        if len(source_styles) == 1 and source_style and not has_direct_paragraph_formatting:
            style_id = source_style
        else:
            style_id = ROLE_TO_ARCH_STYLE.get(role)
            if not style_id:
                continue
            if style_id not in existing_style_ids and style_id not in created_style_ids:
                style_spec = {
                    "styleId": style_id,
                    "name": ROLE_TO_STYLE_NAME.get(role, style_id),
                    "type": "paragraph",
                    "derive_from_paragraph_index": exemplar_idx,
                }
                if source_style:
                    style_spec["basedOn"] = source_style
                else:
                    defaults = [
                        sid
                        for sid, info in slim_bundle.get("style_catalog", {}).items()
                        if isinstance(info, dict) and info.get("default") is True
                    ]
                    if len(defaults) == 1:
                        style_spec["basedOn"] = defaults[0]
                    elif "Normal" in existing_style_ids:
                        style_spec["basedOn"] = "Normal"
                instructions.setdefault("create_styles", []).append(style_spec)
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

    _, strong_hits = infer_expected_roles(
        paragraphs,
        numbering_catalog=slim_bundle.get("numbering_catalog"),
    )

    def _find_correct_exemplar(role: str) -> Optional[int]:
        hits = strong_hits.get(role, [])
        return hits[0] if hits else None

    roles_by_index: Dict[int, Set[str]] = {}
    for detected_role, indices in strong_hits.items():
        for index in indices:
            roles_by_index.setdefault(int(index), set()).add(detected_role)

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

        detected_roles = roles_by_index.get(exemplar_idx, set())
        if not detected_roles or role in detected_roles:
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
                corrected_source_style = by_index[correct_idx].get("pStyle")
                if corrected_source_style:
                    sd["basedOn"] = corrected_source_style
                else:
                    sd.pop("basedOn", None)

        corrections += 1

    return corrections


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
        f"Use ONLY the styleIds from the roles above. If a paragraph is genuinely non-CSI content, "
        f"put it in ignored_paragraphs with a specific reason; never guess from a neighboring style.\n\n"
        f"Context paragraphs (surrounding the missed ones):\n"
        f"{json.dumps(context_paragraphs, indent=2)}\n\n"
        f'Return ONLY a JSON object with these keys: '
        f'{{"apply_pStyle": [{{"paragraph_index": N, "styleId": "..."}}], '
        f'"ignored_paragraphs": [{{"paragraph_index": N, "reason": "..."}}]}}\n'
        f"No prose, no markdown, no extra keys."
    )
    return prompt


def _validate_patch_result(
    patch_result: Dict[str, Any],
    missing_indices: List[int],
    instructions: Dict[str, Any],
) -> None:
    """Validate a targeted classifier response before merging any of it."""
    if set(patch_result) != {"apply_pStyle", "ignored_paragraphs"}:
        raise ValueError(
            "Targeted classification response must contain exactly "
            "apply_pStyle and ignored_paragraphs"
        )
    if not isinstance(patch_result["apply_pStyle"], list) or not isinstance(
        patch_result["ignored_paragraphs"], list
    ):
        raise ValueError("Targeted classification response fields must be arrays")

    allowed_indices = set(missing_indices)
    allowed_styles = {
        spec.get("styleId")
        for spec in instructions.get("roles", {}).values()
        if isinstance(spec, dict) and isinstance(spec.get("styleId"), str)
    }
    seen: Set[int] = set()
    for position, item in enumerate(patch_result["apply_pStyle"]):
        if not isinstance(item, dict) or set(item) != {"paragraph_index", "styleId"}:
            raise ValueError(f"Targeted apply_pStyle[{position}] has an invalid shape")
        index = item["paragraph_index"]
        if type(index) is not int or index not in allowed_indices or index in seen:
            raise ValueError(
                f"Targeted apply_pStyle[{position}] must reference one unique missing paragraph"
            )
        if item["styleId"] not in allowed_styles:
            raise ValueError(
                f"Targeted apply_pStyle[{position}] uses a style not declared in roles"
            )
        seen.add(index)
    for position, item in enumerate(patch_result["ignored_paragraphs"]):
        if not isinstance(item, dict) or set(item) != {"paragraph_index", "reason"}:
            raise ValueError(f"Targeted ignored_paragraphs[{position}] has an invalid shape")
        index = item["paragraph_index"]
        reason = item["reason"]
        if type(index) is not int or index not in allowed_indices or index in seen:
            raise ValueError(
                f"Targeted ignored_paragraphs[{position}] must reference one unique missing paragraph"
            )
        if not isinstance(reason, str) or not reason.strip():
            raise ValueError(
                f"Targeted ignored_paragraphs[{position}].reason must be non-empty"
            )
        seen.add(index)


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

    If the initial classification misses paragraphs, targeted follow-up calls
    classify only those indices. Ambiguous gaps fail closed after the bounded
    attempts; they are never assigned a neighboring style by position.

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
    bundle_json = json.dumps(slim_bundle, indent=2)
    input_tokens = estimate_tokens(master_prompt + run_instruction + bundle_json)
    if input_tokens > MAX_SINGLE_PASS_INPUT_TOKENS:
        raise ValueError(
            f"Document classification input is approximately {input_tokens:,} tokens, exceeding the "
            f"safe single-pass limit ({MAX_SINGLE_PASS_INPUT_TOKENS:,}). Reduce the template or use a model "
            "with a larger context window; no partial classification was produced."
        )

    import anthropic
    from docx_decomposer import validate_instructions

    # _call_api owns the bounded retry policy. Disable the SDK's implicit
    # retries so transport attempts do not multiply behind that policy.
    client = anthropic.Anthropic(api_key=api_key, timeout=180.0, max_retries=0)
    user_message = f"{run_instruction}\n\nSlim bundle:\n{bundle_json}"

    raw = _call_api(client, master_prompt, user_message, model)
    instructions = _parse_response(raw)

    normalized_exclusions = _normalize_known_exclusions(instructions, slim_bundle)
    if normalized_exclusions:
        print(f"Recorded {normalized_exclusions} deterministic non-CSI exclusion(s)")

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
                raise ValueError(
                    f"Classification remained incomplete after {max_patch_attempts} targeted patch attempt(s); "
                    f"unresolved paragraph indices: {missing}"
                ) from exc

            print(
                f"Coverage patch attempt {patch_attempt + 1}/{max_patch_attempts}: "
                f"classifying {len(missing)} missing paragraph(s): {missing}"
            )

            patch_prompt = _build_patch_prompt(slim_bundle, instructions, missing)
            patch_raw = _call_api(client, master_prompt, patch_prompt, model)
            patch_result = _parse_response(patch_raw)
            _validate_patch_result(patch_result, missing, instructions)

            # Merge patch results into instructions
            existing_indices = {
                item["paragraph_index"]
                for item in instructions.get("apply_pStyle", [])
            }
            existing_indices.update(
                item["paragraph_index"]
                for item in instructions.get("ignored_paragraphs", [])
                if isinstance(item, dict) and isinstance(item.get("paragraph_index"), int)
            )
            for item in patch_result.get("apply_pStyle", []):
                idx = item["paragraph_index"]
                if idx not in existing_indices:
                    instructions.setdefault("apply_pStyle", []).append(item)
                    existing_indices.add(idx)
            for item in patch_result.get("ignored_paragraphs", []):
                idx = item["paragraph_index"]
                if idx not in existing_indices:
                    instructions.setdefault("ignored_paragraphs", []).append(item)
                    existing_indices.add(idx)

            # Re-sort for deterministic output
            instructions["apply_pStyle"] = sorted(
                instructions.get("apply_pStyle", []),
                key=lambda x: x["paragraph_index"],
            )
            instructions["ignored_paragraphs"] = sorted(
                instructions.get("ignored_paragraphs", []),
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
    Compute what percentage of classifiable paragraphs were explicitly handled.

    Returns:
        (coverage_fraction, handled_count, classifiable_count)
    """
    paragraphs = slim_bundle.get("paragraphs", [])
    classifiable_indices = {p["paragraph_index"] for p in paragraphs if is_classifiable_paragraph(p)}
    classifiable = len(classifiable_indices)

    styled_indices = {item["paragraph_index"] for item in instructions.get("apply_pStyle", [])}
    ignored_indices = {
        item["paragraph_index"]
        for item in instructions.get("ignored_paragraphs", [])
        if isinstance(item, dict) and isinstance(item.get("paragraph_index"), int)
    }
    handled_count = len((styled_indices | ignored_indices) & classifiable_indices)

    coverage = handled_count / classifiable if classifiable > 0 else 1.0
    return coverage, handled_count, classifiable
