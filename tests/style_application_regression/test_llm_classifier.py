import json
import threading
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

    fake_anthropic = types.SimpleNamespace(Anthropic=lambda api_key: fake)
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
    fake_anthropic = types.SimpleNamespace(Anthropic=lambda api_key: fake)
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
    fake_anthropic = types.SimpleNamespace(Anthropic=lambda api_key: fake)
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
    fake_anthropic = types.SimpleNamespace(Anthropic=lambda api_key: fake)
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
    fake_anthropic = types.SimpleNamespace(Anthropic=lambda api_key: fake)
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
