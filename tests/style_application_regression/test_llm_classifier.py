import json
import threading
import time
import types

import pytest

from spec_formatter.style_application.core.llm_classifier import (
    _build_user_message,
    _merge_chunk_results,
    _parse_classification_response,
    classify_target_document,
)


class _FakeStream:
    def __init__(self, payload):
        self.payload = payload

    def __enter__(self):
        return self

    def __exit__(self, exc_type, exc, tb):
        return False

    def get_final_text(self):
        return self.payload


class _FakeMessages:
    def __init__(self, payload='{"classifications": []}'):
        self.last_kwargs = None
        self.payload = payload

    def stream(self, **kwargs):
        self.last_kwargs = kwargs
        return _FakeStream(self.payload)


class _FakeClient:
    def __init__(self):
        self.messages = _FakeMessages()


def test_output_config_is_dict(monkeypatch):
    fake = _FakeClient()

    fake_anthropic = types.SimpleNamespace(Anthropic=lambda **_kwargs: fake)
    monkeypatch.setitem(__import__("sys").modules, "anthropic", fake_anthropic)

    bundle = {
        "paragraphs": [],
        "available_roles": ["PART"],
        "deterministic_classifications": [],
    }
    result = classify_target_document(bundle, ["PART"], api_key="x", model="m")
    assert fake.messages.last_kwargs is None
    assert result["notes"] == ["LLM skipped: all paragraphs classified deterministically."]


def test_classify_calls_llm_for_unresolved(monkeypatch):
    fake = _FakeClient()
    fake.messages.payload = '{"classifications": [{"paragraph_index": 3, "csi_role": "PART"}]}'
    fake_anthropic = types.SimpleNamespace(Anthropic=lambda **_kwargs: fake)
    monkeypatch.setitem(__import__("sys").modules, "anthropic", fake_anthropic)

    bundle = {
        "paragraphs": [{"paragraph_index": 3, "text": "A"}],
        "available_roles": ["PART"],
        "deterministic_classifications": [],
    }
    classify_target_document(bundle, ["PART"], api_key="x", model="m")
    output_config = fake.messages.last_kwargs["output_config"]
    assert output_config["effort"] == "high"
    assert output_config["format"]["type"] == "json_schema"
    role_schema = output_config["format"]["schema"]["properties"][
        "classifications"
    ]["items"]["properties"]["csi_role"]
    assert role_schema["enum"] == ["PART"]
    ignored_schema = output_config["format"]["schema"]["properties"][
        "ignored_paragraphs"
    ]
    assert ignored_schema["items"]["properties"]["reason"]["enum"] == [
        "non_csi_content"
    ]
    schema = output_config["format"]["schema"]
    assert set(schema["properties"]) == {"classifications", "ignored_paragraphs"}
    assert set(schema["required"]) == set(schema["properties"])


def test_user_message_exposes_only_unresolved_paragraphs():
    bundle = {
        "paragraphs": [{"paragraph_index": 3, "text": "A"}],
        "deterministic_classifications": [
            {"paragraph_index": 1, "csi_role": "PART"}
        ],
        "filter_report": {
            "paragraphs_removed_entirely": [{"paragraph_index": 2}],
        },
    }

    content = _build_user_message(bundle)
    prompt_bundle = json.loads(content)

    assert prompt_bundle == {
        "paragraphs": [{"paragraph_index": 3, "text": "A"}],
    }


@pytest.mark.parametrize(
    "payload",
    [
        '```json\n{"classifications": []}\n```',
        'Here is the requested result:\n{"classifications": []}',
        '<json>{"classifications": []}</json>',
    ],
)
def test_parse_classification_response_recovers_one_wrapped_object(payload):
    assert _parse_classification_response(payload) == {"classifications": []}


def test_parse_classification_response_rejects_multiple_distinct_objects():
    with pytest.raises(json.JSONDecodeError, match="multiple distinct"):
        _parse_classification_response(
            '{"classifications": []}\n'
            '{"classifications": [{"paragraph_index": 1, "csi_role": "PART"}]}'
        )


def test_empty_response_retry_uses_stricter_json_instruction(monkeypatch):
    class SequenceMessages:
        def __init__(self):
            self.payloads = [
                "",
                'Result: {"classifications": '
                '[{"paragraph_index": 3, "csi_role": "PART"}]}',
            ]
            self.prompts = []

        def stream(self, **kwargs):
            self.prompts.append(kwargs["messages"][0]["content"])
            return _FakeStream(self.payloads.pop(0))

    messages = SequenceMessages()
    fake = types.SimpleNamespace(messages=messages)
    fake_anthropic = types.SimpleNamespace(Anthropic=lambda **_kwargs: fake)
    monkeypatch.setitem(__import__("sys").modules, "anthropic", fake_anthropic)
    monkeypatch.setattr(
        "spec_formatter.style_application.core.llm_classifier.time.sleep",
        lambda _seconds: None,
    )
    bundle = {
        "paragraphs": [{"paragraph_index": 3, "text": "A"}],
        "available_roles": ["PART"],
        "deterministic_classifications": [],
    }

    result = classify_target_document(bundle, ["PART"], api_key="x", model="m")

    assert result["classifications"] == [
        {"paragraph_index": 3, "csi_role": "PART"}
    ]
    assert len(messages.prompts) == 2
    assert "RETRY REQUIREMENT" not in messages.prompts[0]
    assert "RETRY REQUIREMENT" in messages.prompts[1]


def test_validation_retry_names_exact_allowed_indices(monkeypatch):
    class SequenceMessages:
        def __init__(self):
            self.payloads = [
                json.dumps({
                    "classifications": [
                        {"paragraph_index": 1, "csi_role": "PART"},
                    ],
                    "notes": [],
                }),
                json.dumps({
                    "classifications": [
                        {"paragraph_index": 3, "csi_role": "PART"},
                    ],
                    "notes": [],
                }),
            ]
            self.prompts = []

        def stream(self, **kwargs):
            self.prompts.append(kwargs["messages"][0]["content"])
            return _FakeStream(self.payloads.pop(0))

    messages = SequenceMessages()
    fake = types.SimpleNamespace(messages=messages)
    fake_anthropic = types.SimpleNamespace(Anthropic=lambda **_kwargs: fake)
    monkeypatch.setitem(__import__("sys").modules, "anthropic", fake_anthropic)
    monkeypatch.setattr(
        "spec_formatter.style_application.core.llm_classifier.time.sleep",
        lambda _seconds: None,
    )
    bundle = {
        "paragraphs": [{"paragraph_index": 3, "text": "A"}],
        "available_roles": ["PART"],
        "deterministic_classifications": [
            {"paragraph_index": 1, "csi_role": "PART"}
        ],
    }

    result = classify_target_document(bundle, ["PART"], api_key="x", model="m")

    assert result["classifications"] == [
        {"paragraph_index": 1, "csi_role": "PART"},
        {"paragraph_index": 3, "csi_role": "PART"},
    ]
    assert "classification index not allowed: 1" in messages.prompts[1]
    assert "and no other indices: [3]" in messages.prompts[1]


def test_classify_accepts_explicit_non_csi_disposition(monkeypatch):
    fake = _FakeClient()
    fake.messages.payload = json.dumps({
        "classifications": [],
        "ignored_paragraphs": [
            {"paragraph_index": 3, "reason": "non_csi_content"},
        ],
        "notes": [],
    })
    fake_anthropic = types.SimpleNamespace(Anthropic=lambda **_kwargs: fake)
    monkeypatch.setitem(__import__("sys").modules, "anthropic", fake_anthropic)
    bundle = {
        "paragraphs": [{"paragraph_index": 3, "text": "Document control note"}],
        "available_roles": ["PART"],
        "deterministic_classifications": [],
    }

    result = classify_target_document(bundle, ["PART"], api_key="x", model="m")

    assert result["classifications"] == []
    assert result["ignored_paragraphs"] == [
        {"paragraph_index": 3, "reason": "non_csi_content"},
    ]


def test_split_bundle_terminates_when_filter_report_dominates():
    # Regression: a bundle pushed over the char threshold by a huge
    # filter_report (not by paragraph volume) used to yield a
    # paras_per_chunk <= _CHUNK_OVERLAP, so the chunk window walked
    # backwards and the split loop never terminated.
    from spec_formatter.style_application.core.llm_classifier import _CHUNK_OVERLAP, _split_bundle_into_chunks

    paragraphs = [{"paragraph_index": i, "text": f"P{i}"} for i in range(21)]
    bundle = {
        "available_roles": ["PART"],
        "filter_report": {
            "paragraphs_removed_entirely": [
                {
                    "paragraph_index": i,
                    "tags": ["masterspec_instruction"],
                    "original_text_preview": "x" * 120,
                }
                for i in range(3000)
            ],
            "paragraphs_stripped": [],
        },
        "paragraphs": paragraphs,
    }

    chunks = _split_bundle_into_chunks(bundle, max_chars=240_000)

    covered = [p["paragraph_index"] for chunk in chunks for p in chunk["paragraphs"]]
    assert set(covered) == set(range(21))
    for chunk in chunks[:-1]:
        assert len(chunk["paragraphs"]) > _CHUNK_OVERLAP


def test_merge_chunk_results_conflict_raises():
    with pytest.raises(ValueError, match="conflicts"):
        _merge_chunk_results([
            {"classifications": [{"paragraph_index": 4, "csi_role": "PART"}], "notes": []},
            {"classifications": [{"paragraph_index": 4, "csi_role": "ARTICLE"}], "notes": []},
        ])


class _CountingMessages:
    def __init__(self):
        self.call_count = 0
        self.calls = []
        self.lock = threading.Lock()

    def stream(self, **kwargs):
        with self.lock:
            self.call_count += 1
            self.calls.append(kwargs)
        content = kwargs["messages"][0]["content"]
        slim_bundle = json.loads(content)
        classifications = [
            {"paragraph_index": p["paragraph_index"], "csi_role": "PART"}
            for p in slim_bundle.get("paragraphs", [])
        ]
        return _FakeStream(json.dumps({"classifications": classifications}))


class _CountingClient:
    def __init__(self):
        self.messages = _CountingMessages()


def test_chunk_classification_runs_all_chunks(monkeypatch):
    fake = _CountingClient()
    fake_anthropic = types.SimpleNamespace(Anthropic=lambda **_kwargs: fake)
    monkeypatch.setitem(__import__("sys").modules, "anthropic", fake_anthropic)

    bundle = {
        "paragraphs": [{"paragraph_index": i, "text": f"P{i}"} for i in range(8)],
        "available_roles": ["PART"],
        "deterministic_classifications": [],
    }

    from spec_formatter.style_application.core import llm_classifier as lc

    monkeypatch.setattr(lc, "_split_bundle_into_chunks", lambda slim_bundle: [
        {"paragraphs": bundle["paragraphs"][:4]},
        {"paragraphs": bundle["paragraphs"][4:]},
    ])

    result = classify_target_document(bundle, ["PART"], api_key="x", model="m")

    assert fake.messages.call_count == 2
    assert len(result["classifications"]) == 8


def test_cached_system_block_is_byte_stable_across_chunks_and_targets(monkeypatch):
    fake = _CountingClient()
    monkeypatch.setitem(__import__("sys").modules, "anthropic", types.SimpleNamespace(Anthropic=lambda **_kwargs: fake))
    for target, count in (("first", 320), ("second", 330)):
        bundle = {"paragraphs": [{"paragraph_index": i, "text": f"{target} requirement {i}"} for i in range(count)],
                  "available_roles": ["PART"]}
        result = classify_target_document(bundle, ["PART"], api_key="x", model="m")
        assert len(result["classifications"]) == count
    assert len(fake.messages.calls) == 4
    blocks = [json.dumps(call["system"], separators=(",", ":")).encode("utf-8") for call in fake.messages.calls]
    assert len(set(blocks)) == 1
    assert all("available_roles" not in call["messages"][0]["content"] for call in fake.messages.calls)


# ---------------------------------------------------------------------------
# Transport hardening: typed retries, request limiter, stop reasons
# ---------------------------------------------------------------------------


def _fake_sdk(monkeypatch, client, constructed):
    class APIStatusError(Exception):
        def __init__(self, message, status_code=500, headers=None):
            super().__init__(message)
            self.status_code = status_code
            self.response = types.SimpleNamespace(headers=headers or {})

    class RateLimitError(APIStatusError):
        def __init__(self, message, headers=None):
            super().__init__(message, status_code=429, headers=headers)

    class AuthenticationError(APIStatusError):
        def __init__(self, message):
            super().__init__(message, status_code=401)

    class APIConnectionError(Exception):
        pass

    def anthropic_ctor(**kwargs):
        constructed.append(kwargs)
        return client

    fake_anthropic = types.SimpleNamespace(
        Anthropic=anthropic_ctor,
        APIStatusError=APIStatusError,
        RateLimitError=RateLimitError,
        AuthenticationError=AuthenticationError,
        APIConnectionError=APIConnectionError,
    )
    monkeypatch.setitem(__import__("sys").modules, "anthropic", fake_anthropic)
    return fake_anthropic


class _ScriptedStream(_FakeStream):
    def __init__(self, payload, stop_reason="end_turn"):
        super().__init__(payload)
        self.stop_reason = stop_reason

    def get_final_message(self):
        return types.SimpleNamespace(stop_reason=self.stop_reason)


class _ScriptedMessages:
    def __init__(self, outcomes):
        self.outcomes = list(outcomes)
        self.calls = []

    def stream(self, **kwargs):
        self.calls.append(kwargs)
        outcome = self.outcomes.pop(0)
        if isinstance(outcome, Exception):
            raise outcome
        return outcome


def _unresolved_bundle():
    return {
        "paragraphs": [{"paragraph_index": 0, "text": "A. Scope"}],
        "deterministic_classifications": [],
        "deterministic_ignored_paragraphs": [],
    }


_GOOD = '{"classifications": [{"paragraph_index": 0, "csi_role": "PART"}]}'


def _run(monkeypatch, outcomes):
    from spec_formatter.style_application.core import llm_classifier as lc

    constructed = []
    messages = _ScriptedMessages(outcomes)
    client = types.SimpleNamespace(messages=messages)
    sdk = _fake_sdk(monkeypatch, client, constructed)
    sleeps = []
    monkeypatch.setattr(lc.time, "sleep", sleeps.append)
    return sdk, messages, sleeps, constructed


def test_client_disables_sdk_retries_and_sets_a_connect_timeout(monkeypatch):
    _sdk, _messages, _sleeps, constructed = _run(monkeypatch, [_ScriptedStream(_GOOD)])

    classify_target_document(_unresolved_bundle(), ["PART"], api_key="k", model="m")

    assert constructed[0]["max_retries"] == 0
    timeout = constructed[0]["timeout"]
    assert timeout.connect == 5.0
    assert timeout.read == 600.0


@pytest.mark.parametrize("effort", ["low", "medium", "high", "xhigh", "max"])
def test_target_effort_used_on_initial_request_and_regeneration(monkeypatch, effort):
    _sdk, messages, _sleeps, _constructed = _run(
        monkeypatch, [_ScriptedStream("invalid JSON"), _ScriptedStream(_GOOD)]
    )
    classify_target_document(
        _unresolved_bundle(), ["PART"], api_key="k", model="m", target_effort=effort
    )
    assert len(messages.calls) == 2
    assert [call["output_config"]["effort"] for call in messages.calls] == [effort, effort]
    assert all(call["thinking"] == {"type": "adaptive"} for call in messages.calls)


def test_authentication_error_makes_one_request_and_never_sleeps(monkeypatch):
    sdk, messages, sleeps, _constructed = _run(monkeypatch, [])
    messages.outcomes = [sdk.AuthenticationError("invalid x-api-key")]

    with pytest.raises(sdk.AuthenticationError, match="invalid x-api-key"):
        classify_target_document(_unresolved_bundle(), ["PART"], api_key="bad", model="m")

    assert len(messages.calls) == 1
    assert sleeps == []


def test_bad_request_is_not_retried(monkeypatch):
    sdk, messages, sleeps, _constructed = _run(monkeypatch, [])
    messages.outcomes = [sdk.APIStatusError("bad request", status_code=400)]

    with pytest.raises(sdk.APIStatusError, match="bad request"):
        classify_target_document(_unresolved_bundle(), ["PART"], api_key="k", model="m")

    assert len(messages.calls) == 1
    assert sleeps == []


def test_server_error_and_connection_error_back_off_then_succeed(monkeypatch):
    sdk, messages, sleeps, _constructed = _run(monkeypatch, [])
    messages.outcomes = [
        sdk.APIStatusError("upstream", status_code=503),
        sdk.APIConnectionError("offline"),
        _ScriptedStream(_GOOD),
    ]

    result = classify_target_document(_unresolved_bundle(), ["PART"], api_key="k", model="m")

    assert result["classifications"] == [{"paragraph_index": 0, "csi_role": "PART"}]
    assert len(messages.calls) == 3
    assert sleeps == [2.0, 4.0]
    # Transport retries never append the JSON regeneration instruction.
    assert all("RETRY REQUIREMENT" not in call["messages"][0]["content"] for call in messages.calls)


def test_rate_limit_honours_retry_after(monkeypatch):
    sdk, messages, sleeps, _constructed = _run(monkeypatch, [])
    messages.outcomes = [
        sdk.RateLimitError("slow down", headers={"retry-after": "3"}),
        _ScriptedStream(_GOOD),
    ]

    classify_target_document(_unresolved_bundle(), ["PART"], api_key="k", model="m")

    assert sleeps == [3.0]


def test_transient_failures_are_bounded(monkeypatch):
    sdk, messages, sleeps, _constructed = _run(monkeypatch, [])
    messages.outcomes = [sdk.APIConnectionError(f"offline {i}") for i in range(4)]

    with pytest.raises(sdk.APIConnectionError, match="offline 2"):
        classify_target_document(_unresolved_bundle(), ["PART"], api_key="k", model="m")

    assert len(messages.calls) == 3
    assert sleeps == [2.0, 4.0]


def test_refusal_is_terminal(monkeypatch):
    from spec_formatter.style_application.core.llm_classifier import ClassificationRefused

    _sdk, messages, sleeps, _constructed = _run(
        monkeypatch, [_ScriptedStream("", stop_reason="refusal"), _ScriptedStream(_GOOD)]
    )

    with pytest.raises(ClassificationRefused, match="refusal"):
        classify_target_document(_unresolved_bundle(), ["PART"], api_key="k", model="m")

    assert len(messages.calls) == 1
    assert sleeps == []


class _RefusalStream(_ScriptedStream):
    def __init__(self, stop_details):
        super().__init__("", stop_reason="refusal")
        self.stop_details = stop_details

    def get_final_message(self):
        return types.SimpleNamespace(
            stop_reason="refusal",
            stop_details=self.stop_details,
        )


@pytest.mark.parametrize(
    ("stop_details", "expected"),
    [
        (
            types.SimpleNamespace(
                type="refusal",
                category="general_harms",
                explanation="Free text the provider wrote about the request.",
            ),
            "general_harms",
        ),
        ({"type": "refusal", "category": "cyber", "explanation": "Free text."}, "cyber"),
        (types.SimpleNamespace(type="refusal", category=None, explanation=None), None),
        (None, None),
        (types.SimpleNamespace(category="Not an identifier!"), None),
    ],
)
def test_refusal_carries_its_code_and_only_an_identifier_category(
    monkeypatch, stop_details, expected
):
    from spec_formatter.style_application.core.errors import ERROR_REMEDIATIONS
    from spec_formatter.style_application.core.llm_classifier import ClassificationRefused

    _run(monkeypatch, [_RefusalStream(stop_details)])

    with pytest.raises(ClassificationRefused) as caught:
        classify_target_document(_unresolved_bundle(), ["PART"], api_key="k", model="m")

    error = caught.value
    assert error.safe_error_code == "classification_refused"
    assert error.safe_error_message == ERROR_REMEDIATIONS["classification_refused"]
    assert error.refusal_category == expected
    assert f"category={expected or 'none'}" in str(error)
    assert "Free text" not in str(error)


def test_max_tokens_regenerates_with_the_retry_requirement(monkeypatch):
    _sdk, messages, _sleeps, _constructed = _run(
        monkeypatch,
        [_ScriptedStream('{"classifications": [', stop_reason="max_tokens"), _ScriptedStream(_GOOD)],
    )

    result = classify_target_document(_unresolved_bundle(), ["PART"], api_key="k", model="m")

    assert result["classifications"] == [{"paragraph_index": 0, "csi_role": "PART"}]
    assert len(messages.calls) == 2
    assert "RETRY REQUIREMENT" in messages.calls[1]["messages"][0]["content"]
    assert "max_tokens" in messages.calls[1]["messages"][0]["content"]


def test_request_limiter_bounds_concurrent_streams(monkeypatch):
    from spec_formatter.style_application.core import llm_classifier as lc

    limit = 2
    lock = threading.Lock()
    state = {"in_flight": 0, "peak": 0}

    class LimitedMessages:
        def stream(self, **kwargs):
            with lock:
                state["in_flight"] += 1
                state["peak"] = max(state["peak"], state["in_flight"])
            time.sleep(0.02)
            content = kwargs["messages"][0]["content"]
            slim_bundle = json.loads(content)
            payload = json.dumps({
                "classifications": [
                    {"paragraph_index": p["paragraph_index"], "csi_role": "PART"}
                    for p in slim_bundle["paragraphs"]
                ]
            })
            with lock:
                state["in_flight"] -= 1
            return _FakeStream(payload)

    client = types.SimpleNamespace(messages=LimitedMessages())
    _fake_sdk(monkeypatch, client, [])
    monkeypatch.setattr(lc, "_REQUEST_LIMITER", threading.BoundedSemaphore(limit))
    paragraphs = [{"paragraph_index": i, "text": f"P{i}"} for i in range(12)]
    monkeypatch.setattr(
        lc,
        "_split_bundle_into_chunks",
        lambda slim_bundle: [{"paragraphs": [p]} for p in paragraphs],
    )
    bundle = {
        "paragraphs": paragraphs,
        "deterministic_classifications": [],
        "deterministic_ignored_paragraphs": [],
    }

    result = classify_target_document(bundle, ["PART"], api_key="k", model="m")

    assert len(result["classifications"]) == 12
    assert state["peak"] <= limit


def test_request_limit_env_var_is_bounded(monkeypatch):
    from spec_formatter.style_application.core import llm_classifier as lc

    monkeypatch.setenv(lc._MAX_CONCURRENT_REQUESTS_ENV, "0")
    assert lc._max_concurrent_requests() == 1
    monkeypatch.setenv(lc._MAX_CONCURRENT_REQUESTS_ENV, "1000")
    assert lc._max_concurrent_requests() == lc._MAX_CONCURRENT_REQUESTS_CEILING
    monkeypatch.setenv(lc._MAX_CONCURRENT_REQUESTS_ENV, "nonsense")
    assert lc._max_concurrent_requests() == lc._DEFAULT_MAX_CONCURRENT_REQUESTS
    monkeypatch.delenv(lc._MAX_CONCURRENT_REQUESTS_ENV)
    assert lc._max_concurrent_requests() == lc._DEFAULT_MAX_CONCURRENT_REQUESTS


# ---------------------------------------------------------------------------
# Compact wire JSON, cached system prefix, usage accounting, overlap re-ask
# ---------------------------------------------------------------------------


def test_system_prefix_is_one_cached_block_and_user_turn_is_compact(monkeypatch):
    from spec_formatter.style_application.core.classification import (
        PHASE2_MASTER_PROMPT,
        PHASE2_RUN_INSTRUCTION,
    )

    _sdk, messages, _sleeps, _constructed = _run(monkeypatch, [_ScriptedStream(_GOOD)])

    classify_target_document(_unresolved_bundle(), ["PART"], api_key="k", model="m")

    kwargs = messages.calls[0]
    assert isinstance(kwargs["system"], list) and len(kwargs["system"]) == 1
    block = kwargs["system"][0]
    assert block["type"] == "text"
    assert block["cache_control"] == {"type": "ephemeral"}
    assert block["text"].startswith(PHASE2_MASTER_PROMPT.strip())
    assert 'available_roles: ["PART"]' in block["text"]
    # The run instruction closes the system prompt, and it closes with the
    # think-first line: a structured-output response is JSON only, so the
    # model can work a classification out nowhere but in its thinking.
    assert block["text"].endswith(PHASE2_RUN_INSTRUCTION.strip())
    assert block["text"].index('available_roles: ["PART"]') < block["text"].index(
        PHASE2_RUN_INSTRUCTION.strip()
    )
    assert block["text"].endswith("Think the problem through before you answer.")
    content = kwargs["messages"][0]["content"]
    assert content.startswith('{')
    assert "available_roles" not in content
    assert "\n  " not in content  # compact JSON, no indentation
    assert PHASE2_RUN_INSTRUCTION.strip() not in content


def test_chunker_measures_the_bytes_the_request_sends():
    from spec_formatter.style_application.core.llm_classifier import (
        _build_user_message,
        _split_bundle_into_chunks,
        _wire_json,
    )

    paragraphs = [{"paragraph_index": i, "text": f"Paragraph {i}"} for i in range(40)]
    bundle = {"paragraphs": paragraphs, "available_roles": ["PART"]}
    compact = len(_wire_json({"paragraphs": paragraphs}))
    pretty = len(json.dumps({"paragraphs": paragraphs}, indent=2))
    assert compact < pretty

    # A limit between the two sizes keeps one chunk: the guard measures what
    # is sent, so an indent=2 layout can no longer under-count by a third.
    chunks = _split_bundle_into_chunks(bundle, max_chars=compact)
    assert len(chunks) == 1
    sent = _build_user_message(bundle)
    assert sent == _wire_json({"paragraphs": paragraphs})


def test_usage_numbers_are_summed_across_requests(monkeypatch):
    class UsageStream(_ScriptedStream):
        def __init__(self, payload, cache_read):
            super().__init__(payload)
            self.cache_read = cache_read

        def get_final_message(self):
            return types.SimpleNamespace(
                stop_reason="end_turn",
                usage=types.SimpleNamespace(
                    input_tokens=100,
                    output_tokens=10,
                    cache_read_input_tokens=self.cache_read,
                    cache_creation_input_tokens=0 if self.cache_read else 900,
                ),
            )

    _sdk, _messages, _sleeps, _constructed = _run(
        monkeypatch,
        [UsageStream('{"classifications": [', 0), UsageStream(_GOOD, 900)],
    )

    result = classify_target_document(_unresolved_bundle(), ["PART"], api_key="k", model="m")

    # "requests" used to mean both "attempted" and "answered". The shared
    # contract splits them, so an attempt whose usage never arrived is
    # visible instead of being averaged into a total that looks complete.
    assert result["usage"] == {
        "requests_attempted": 2,
        "responses_completed": 2,
        "responses_with_usage": 2,
        "requests_with_unknown_usage": 0,
        "usage_complete": True,
        "input_tokens": 200,
        "output_tokens": 20,
        "cache_read_input_tokens": 900,
        "cache_creation_input_tokens": 900,
    }


def _overlap_client(monkeypatch, answers):
    """Chunk calls answered from ``answers`` keyed by _chunk_info.chunk_index."""

    class OverlapMessages:
        def __init__(self):
            self.calls = []
            self.lock = threading.Lock()

        def stream(self, **kwargs):
            content = kwargs["messages"][0]["content"]
            # A regeneration attempt appends the retry requirement after the
            # JSON, so decode the object rather than the whole tail.
            slim_bundle, _end = json.JSONDecoder().raw_decode(
                content
            )
            chunk_index = slim_bundle["_chunk_info"]["chunk_index"]
            with self.lock:
                self.calls.append(slim_bundle)
            return _FakeStream(json.dumps(answers[chunk_index](slim_bundle)))

    client = types.SimpleNamespace(messages=OverlapMessages())
    _fake_sdk(monkeypatch, client, [])
    return client.messages


def _overlap_bundle(monkeypatch):
    from spec_formatter.style_application.core import llm_classifier as lc

    paragraphs = [{"paragraph_index": i, "text": f"P{i}"} for i in range(8)]
    monkeypatch.setattr(lc, "_split_bundle_into_chunks", lambda slim_bundle: [
        {"paragraphs": paragraphs[:5], "_chunk_info": {"chunk_index": 0}},
        {"paragraphs": paragraphs[3:], "_chunk_info": {"chunk_index": 1}},
    ])
    return {
        "paragraphs": paragraphs,
        "available_roles": ["PART", "ARTICLE"],
        "deterministic_classifications": [],
        "deterministic_ignored_paragraphs": [],
    }


def _all_as(role):
    return lambda bundle: {
        "classifications": [
            {"paragraph_index": p["paragraph_index"], "csi_role": role}
            for p in bundle["paragraphs"]
        ]
    }


def test_overlap_disagreement_is_re_asked_once_for_the_whole_window(monkeypatch):
    bundle = _overlap_bundle(monkeypatch)
    messages = _overlap_client(
        monkeypatch,
        {0: _all_as("PART"), 1: _all_as("ARTICLE"), 2: _all_as("ARTICLE")},
    )

    result = classify_target_document(bundle, ["PART", "ARTICLE"], api_key="k", model="m")

    assert len(messages.calls) == 3
    reask = next(call for call in messages.calls if call["_chunk_info"]["chunk_index"] == 2)
    assert reask["_chunk_info"]["overlap_reask"] is True
    assert [p["paragraph_index"] for p in reask["paragraphs"]] == [3, 4]
    roles = {item["paragraph_index"]: item["csi_role"] for item in result["classifications"]}
    assert roles == {0: "PART", 1: "PART", 2: "PART", 3: "ARTICLE", 4: "ARTICLE",
                     5: "ARTICLE", 6: "ARTICLE", 7: "ARTICLE"}


def test_overlap_re_ask_that_omits_a_disputed_index_fails_closed(monkeypatch):
    bundle = _overlap_bundle(monkeypatch)
    messages = _overlap_client(
        monkeypatch,
        {
            0: _all_as("PART"),
            1: _all_as("ARTICLE"),
            2: lambda b: {"classifications": [{"paragraph_index": 3, "csi_role": "PART"}]},
        },
    )

    with pytest.raises((ValueError, RuntimeError), match="coverage|indices|4"):
        classify_target_document(bundle, ["PART", "ARTICLE"], api_key="k", model="m")

    # The re-ask window is asked once (plus its own bounded regeneration
    # attempts); the original chunks are never re-run and there is no second
    # re-ask round.
    reask_calls = [c for c in messages.calls if c["_chunk_info"]["chunk_index"] == 2]
    assert 1 <= len(reask_calls) <= 3
    assert all(
        [p["paragraph_index"] for p in call["paragraphs"]] == [3, 4] for call in reask_calls
    )
    assert sum(1 for c in messages.calls if c["_chunk_info"]["chunk_index"] in (0, 1)) == 2



# --- Usage survives target failure (W2) ------------------------------------


def _counted_stream(payload, stop_reason="end_turn"):
    class Stream(_ScriptedStream):
        def get_final_message(self):
            return types.SimpleNamespace(
                stop_reason=stop_reason,
                usage=types.SimpleNamespace(
                    input_tokens=100,
                    output_tokens=10,
                    cache_read_input_tokens=0,
                    cache_creation_input_tokens=0,
                ),
            )

    return Stream(payload)


def test_usage_survives_a_refused_classification(monkeypatch):
    """A refusal is paid for; it used to leave no trace in the accounting."""
    from spec_formatter.llm_usage import usage_from_exception

    _run(monkeypatch, [_counted_stream("", stop_reason="refusal")])

    with pytest.raises(Exception) as raised:
        classify_target_document(_unresolved_bundle(), ["PART"], api_key="k", model="m")

    observed = usage_from_exception(raised.value)
    assert observed["input_tokens"] == 100
    assert observed["output_tokens"] == 10
    assert observed["responses_completed"] == 1


def test_usage_survives_exhausted_regeneration(monkeypatch):
    """Every wasted attempt is counted, not just the ones that parsed."""
    from spec_formatter.llm_usage import usage_from_exception

    _run(monkeypatch, [_counted_stream("not json") for _ in range(6)])

    with pytest.raises(Exception) as raised:
        classify_target_document(_unresolved_bundle(), ["PART"], api_key="k", model="m")

    observed = usage_from_exception(raised.value)
    assert observed["requests_attempted"] >= 2
    assert observed["input_tokens"] == 100 * observed["responses_completed"]


def test_deterministic_only_target_reports_no_requests(monkeypatch):
    """No unresolved paragraphs means no client and no usage to report."""
    bundle = {
        "paragraphs": [],
        "available_roles": ["PART"],
        "deterministic_classifications": [],
        "deterministic_ignored_paragraphs": [],
        "filter_report": {"paragraphs_removed_entirely": [], "paragraphs_stripped": []},
    }
    _sdk, _messages, _sleeps, constructed = _run(monkeypatch, [])

    result = classify_target_document(bundle, ["PART"], api_key="", model="m")

    assert constructed == []
    # An explicit known zero, not an absent key: "we sent nothing" and "we
    # could not tell you" are different answers and must look different.
    assert result["usage"]["requests_attempted"] == 0
    assert result["usage"]["usage_complete"] is True


def _sse_client(*bodies):
    """A real pinned-SDK client whose requests replay recorded SSE bodies.

    Fabricated final messages cannot show what the SDK's stream accumulator
    drops; replaying the wire events through the real client can.
    """
    import httpx
    import anthropic

    pending = list(bodies)

    def handler(_request):
        return httpx.Response(
            200,
            headers={"content-type": "text/event-stream"},
            content=pending.pop(0),
        )

    return anthropic.Anthropic(
        api_key="k",
        max_retries=0,
        http_client=httpx.Client(transport=httpx.MockTransport(handler)),
    )


def _sse_body(stop_reason, *, text=None, stop_details=None):
    """The wire events of one streamed response, as the API sends them."""
    events = [
        {
            "type": "message_start",
            "message": {
                "id": "msg_1",
                "type": "message",
                "role": "assistant",
                "model": "m",
                "content": [],
                "stop_reason": None,
                "stop_sequence": None,
                "usage": {"input_tokens": 10, "output_tokens": 0},
            },
        }
    ]
    if text is not None:
        events += [
            {"type": "content_block_start", "index": 0, "content_block": {"type": "text", "text": ""}},
            {"type": "content_block_delta", "index": 0, "delta": {"type": "text_delta", "text": text}},
            {"type": "content_block_stop", "index": 0},
        ]
    delta = {"stop_reason": stop_reason, "stop_sequence": None}
    if stop_details is not None:
        delta["stop_details"] = stop_details
    events += [
        {"type": "message_delta", "delta": delta, "usage": {"output_tokens": 3}},
        {"type": "message_stop"},
    ]
    return "".join(
        f"event: {event['type']}\ndata: {json.dumps(event)}\n\n" for event in events
    ).encode("utf-8")


def _classify_over_sse(monkeypatch, *bodies):
    """Run classify_target_document against the real SDK over replayed SSE."""
    import anthropic
    from spec_formatter.style_application.core import llm_classifier as lc

    client = _sse_client(*bodies)
    monkeypatch.setattr(anthropic, "Anthropic", lambda **_kwargs: client)
    monkeypatch.setattr(lc.time, "sleep", lambda _seconds: None)
    return classify_target_document(_unresolved_bundle(), ["PART"], api_key="k", model="m")


def test_a_real_streamed_refusal_without_text_is_a_refusal_with_its_category(monkeypatch):
    """A refusal with no text block used to escape as the SDK's RuntimeError.

    get_final_text() raises when a response holds no text block, and it ran
    before the stop-reason check, so a real refusal never became
    ClassificationRefused and its category was never read.
    """
    from spec_formatter.llm_usage import usage_from_exception
    from spec_formatter.style_application.core.llm_classifier import ClassificationRefused

    with pytest.raises(ClassificationRefused) as caught:
        _classify_over_sse(
            monkeypatch,
            _sse_body(
                "refusal",
                stop_details={"type": "refusal", "category": "general_harms", "explanation": "x"},
            ),
        )

    assert caught.value.safe_error_code == "classification_refused"
    assert caught.value.refusal_category == "general_harms"
    usage = usage_from_exception(caught.value)
    assert usage["input_tokens"] == 10 and usage["output_tokens"] == 3


def test_a_real_output_limit_stop_without_text_regenerates(monkeypatch):
    result = _classify_over_sse(
        monkeypatch,
        _sse_body("max_tokens"),
        _sse_body("end_turn", text=_GOOD),
    )

    assert result["classifications"] == [{"paragraph_index": 0, "csi_role": "PART"}]
    assert result["usage"]["requests_attempted"] == 2
