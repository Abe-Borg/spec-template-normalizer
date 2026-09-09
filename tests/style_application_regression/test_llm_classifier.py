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

    content = _build_user_message(bundle, ["PART"])
    prompt_bundle = json.loads(content[content.rfind("\n\n{") + 2:])

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
        self.lock = threading.Lock()

    def stream(self, **kwargs):
        with self.lock:
            self.call_count += 1
        content = kwargs["messages"][0]["content"]
        json_start = content.rfind("\n\n{")
        slim_bundle = json.loads(content[json_start + 2:])
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
            slim_bundle = json.loads(content[content.rfind("\n\n{") + 2:])
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
    assert PHASE2_RUN_INSTRUCTION.strip() in block["text"]
    assert block["text"].endswith('available_roles: ["PART"]')
    content = kwargs["messages"][0]["content"]
    assert content.startswith('available_roles: ["PART"]\n\n{')
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
    sent = _build_user_message(bundle, ["PART"])
    assert sent.endswith(_wire_json({"paragraphs": paragraphs}))


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
                content[content.find("\n\n{") + 2:]
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
    assert "usage" not in result or result["usage"].get("requests_attempted", 0) == 0
