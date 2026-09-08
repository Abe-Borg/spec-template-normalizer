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

