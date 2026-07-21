from __future__ import annotations

import json
import sys
import types

import pytest

import llm_classifier
from llm_classifier import (
    _call_api,
    _normalize_known_exclusions,
    _parse_response,
    _repair_missing_roles,
    _repair_strong_signal_mismatches,
    _validate_patch_result,
    classify_document,
)


def _install_fake_anthropic(monkeypatch):
    module = types.ModuleType("anthropic")

    class APIConnectionError(Exception):
        pass

    class RateLimitError(Exception):
        pass

    class APIStatusError(Exception):
        def __init__(self, message: str, status_code: int):
            super().__init__(message)
            self.status_code = status_code

    module.APIConnectionError = APIConnectionError
    module.RateLimitError = RateLimitError
    module.APIStatusError = APIStatusError
    module.Anthropic = lambda **_kwargs: object()
    monkeypatch.setitem(sys.modules, "anthropic", module)
    return module


def _plain_bundle(count: int) -> dict:
    return {
        "paragraphs": [
            {
                "paragraph_index": index,
                "text": f"General requirement {index}",
                "pStyle": "Body",
                "numPr": None,
                "effective_numPr": None,
                "pPr_hints": None,
                "rPr_hints": None,
                "contains_sectPr": False,
                "in_table": False,
                "skip_reason": None,
                "text_was_truncated": False,
            }
            for index in range(count)
        ],
        "style_catalog": {
            "Body": {
                "styleId": "Body",
                "type": "paragraph",
                "default": True,
                "resolved_numPr": None,
            }
        },
        "numbering_catalog": {"nums": {}, "abstracts": {}},
    }


def _body_instructions(styled_indices) -> dict:
    return {
        "create_styles": [],
        "apply_pStyle": [
            {"paragraph_index": index, "styleId": "Body"}
            for index in styled_indices
        ],
        "ignored_paragraphs": [],
        "roles": {
            "PARAGRAPH": {
                "styleId": "Body",
                "exemplar_paragraph_index": 0,
            }
        },
        "notes": [],
    }


_MISSING_COMMA_INITIAL_RESPONSE = """{
  "create_styles": [],
  "apply_pStyle": [
    {"paragraph_index": 0, "styleId": "Body"}
  ]
  "ignored_paragraphs": [],
  "roles": {
    "PARAGRAPH": {
      "styleId": "Body",
      "exemplar_paragraph_index": 0
    }
  },
  "notes": []
}"""


def test_parse_response_rejects_duplicate_keys() -> None:
    with pytest.raises(ValueError, match="duplicate JSON key: 'roles'"):
        _parse_response('{"roles": {}, "roles": {}}')


@pytest.mark.parametrize("raw", ["[]", '"instructions"', "null", "42"])
def test_parse_response_requires_top_level_object(raw: str) -> None:
    with pytest.raises(ValueError, match="must be a JSON object"):
        _parse_response(raw)


@pytest.mark.parametrize("constant", ["NaN", "Infinity", "-Infinity"])
def test_parse_response_rejects_nonfinite_numbers(constant: str) -> None:
    with pytest.raises(ValueError, match="non-finite JSON number"):
        _parse_response(f'{{"value": {constant}}}')


def test_known_exclusions_are_deterministic_and_removed_from_style_assignments() -> None:
    bundle = {
        "paragraphs": [
            {"paragraph_index": 0, "skip_reason": "editor_note"},
            {"paragraph_index": 1, "skip_reason": "specifier_note"},
            {"paragraph_index": 2, "skip_reason": "copyright_notice"},
            {"paragraph_index": 3, "skip_reason": None},
        ]
    }
    instructions = {
        "apply_pStyle": [
            {"paragraph_index": index, "styleId": "Body"}
            for index in range(4)
        ],
        "ignored_paragraphs": [
            {"paragraph_index": 1, "reason": "Existing audited reason"},
        ],
    }

    count = _normalize_known_exclusions(instructions, bundle)

    assert count == 3
    assert instructions["apply_pStyle"] == [
        {"paragraph_index": 3, "styleId": "Body"}
    ]
    assert instructions["ignored_paragraphs"] == [
        {"paragraph_index": 0, "reason": "Detected editor/specifier note"},
        {"paragraph_index": 1, "reason": "Existing audited reason"},
        {"paragraph_index": 2, "reason": "Detected copyright/distribution notice"},
    ]


def test_missing_role_repair_captures_uniform_direct_run_formatting() -> None:
    bundle = {
        "paragraphs": [
            {
                "paragraph_index": 0,
                "text": "PART 1 - GENERAL",
                "pStyle": "ArchitectHeading",
                "pPr_hints": None,
                "rPr_hints": {"bold": True},
                "numPr": None,
                "skip_reason": None,
            }
        ],
        "style_catalog": {
            "ArchitectHeading": {"styleId": "ArchitectHeading", "default": False},
            "Normal": {"styleId": "Normal", "default": True},
        },
        "numbering_catalog": {"nums": {}, "abstracts": {}},
    }
    instructions = {
        "create_styles": [],
        "apply_pStyle": [
            {"paragraph_index": 0, "styleId": "ArchitectHeading"},
        ],
        "ignored_paragraphs": [],
        "roles": {},
    }

    added = _repair_missing_roles(instructions, bundle)

    assert added == 1
    assert instructions["roles"]["PART"] == {
        "styleId": "CSI_Part__ARCH",
        "exemplar_paragraph_index": 0,
    }
    assert instructions["create_styles"] == [
        {
            "styleId": "CSI_Part__ARCH",
            "name": "CSI Part (Architect Template)",
            "type": "paragraph",
            "derive_from_paragraph_index": 0,
            "basedOn": "ArchitectHeading",
        }
    ]
    assert instructions["apply_pStyle"][0]["styleId"] == "CSI_Part__ARCH"


def test_strong_signal_repair_is_contextual_not_global_for_numeric_markers() -> None:
    bundle = {
        "paragraphs": [
            {"paragraph_index": 0, "text": "Project Options", "skip_reason": None},
            {"paragraph_index": 1, "text": "1. Alternate finish", "skip_reason": None},
            {"paragraph_index": 2, "text": "A. Scope", "skip_reason": None},
            {"paragraph_index": 3, "text": "1. Requirement", "skip_reason": None},
        ],
        "numbering_catalog": {"nums": {}, "abstracts": {}},
    }
    instructions = {
        "apply_pStyle": [
            {"paragraph_index": 0, "styleId": "Body"},
            {"paragraph_index": 1, "styleId": "SubStyle"},
            {"paragraph_index": 2, "styleId": "SubStyle"},
            {"paragraph_index": 3, "styleId": "ParagraphStyle"},
        ],
        "roles": {
            "PARAGRAPH": {
                "styleId": "ParagraphStyle",
                "exemplar_paragraph_index": 2,
            },
            "SUBPARAGRAPH": {
                "styleId": "SubStyle",
                "exemplar_paragraph_index": 3,
            },
        },
    }

    corrections = _repair_strong_signal_mismatches(instructions, bundle)
    styles = {
        item["paragraph_index"]: item["styleId"]
        for item in instructions["apply_pStyle"]
    }

    assert corrections == 2
    assert styles[1] == "SubStyle"  # no established A. hierarchy yet
    assert styles[2] == "ParagraphStyle"
    assert styles[3] == "SubStyle"


def test_targeted_patch_merges_styled_and_ignored_paragraphs(monkeypatch) -> None:
    _install_fake_anthropic(monkeypatch)
    bundle = _plain_bundle(3)
    responses = [
        json.dumps(_body_instructions([0])),
        json.dumps(
            {
                "apply_pStyle": [
                    {"paragraph_index": 1, "styleId": "Body"},
                ],
                "ignored_paragraphs": [
                    {"paragraph_index": 2, "reason": "Non-CSI appendix marker"},
                ],
            }
        ),
    ]
    prompts = []

    def fake_call(
        _client, _system, user_message, _model, max_tokens=128000, response_schema=None
    ):
        prompts.append(user_message)
        return responses.pop(0)

    monkeypatch.setattr(llm_classifier, "_call_api", fake_call)

    result = classify_document(bundle, "master", "run", "fake-key", max_patch_attempts=1)

    assert result["apply_pStyle"] == [
        {"paragraph_index": 0, "styleId": "Body"},
        {"paragraph_index": 1, "styleId": "Body"},
    ]
    assert result["ignored_paragraphs"] == [
        {"paragraph_index": 2, "reason": "Non-CSI appendix marker"}
    ]
    assert "Classify ONLY the following paragraph indices: [1, 2]" in prompts[1]
    assert "never guess from a neighboring style" in prompts[1]


def test_classifier_regenerates_after_missing_comma_in_initial_response(
    monkeypatch,
) -> None:
    _install_fake_anthropic(monkeypatch)
    responses = [
        _MISSING_COMMA_INITIAL_RESPONSE,
        json.dumps(_body_instructions([0])),
    ]
    prompts = []

    def fake_call(
        _client, _system, user_message, _model, max_tokens=128000, response_schema=None
    ):
        prompts.append(user_message)
        return responses.pop(0)

    monkeypatch.setattr(llm_classifier, "_call_api", fake_call)
    monkeypatch.setattr(llm_classifier.time, "sleep", lambda _seconds: None)

    result = classify_document(
        _plain_bundle(1),
        "master",
        "run",
        "fake-key",
        max_patch_attempts=0,
        max_response_attempts=2,
    )

    assert result == _body_instructions([0])
    assert len(prompts) == 2
    assert prompts[1].startswith(prompts[0])
    assert "RETRY REQUIREMENT" in prompts[1]


def test_classifier_bounds_repeated_malformed_initial_responses(monkeypatch) -> None:
    _install_fake_anthropic(monkeypatch)
    prompts = []

    def fake_call(
        _client, _system, user_message, _model, max_tokens=128000, response_schema=None
    ):
        prompts.append(user_message)
        return _MISSING_COMMA_INITIAL_RESPONSE

    monkeypatch.setattr(llm_classifier, "_call_api", fake_call)
    monkeypatch.setattr(llm_classifier.time, "sleep", lambda _seconds: None)

    with pytest.raises(ValueError) as exc_info:
        classify_document(
            _plain_bundle(1),
            "master",
            "run",
            "fake-key",
            max_patch_attempts=0,
            max_response_attempts=2,
        )

    message = str(exc_info.value)
    assert len(prompts) == 2
    assert "after 2" in message
    assert "attempt" in message
    assert "Raw response" not in message
    assert len(message) < 500


def test_classifier_regenerates_malformed_targeted_patch_then_merges(
    monkeypatch,
) -> None:
    _install_fake_anthropic(monkeypatch)
    responses = [
        json.dumps(_body_instructions([0])),
        """{
  "apply_pStyle": [
    {"paragraph_index": 1, "styleId": "Body"}
  ]
  "ignored_paragraphs": []
}""",
        json.dumps(
            {
                "apply_pStyle": [
                    {"paragraph_index": 1, "styleId": "Body"},
                ],
                "ignored_paragraphs": [],
            }
        ),
    ]
    prompts = []
    schemas = []

    def fake_call(
        _client, _system, user_message, _model, max_tokens=128000, response_schema=None
    ):
        prompts.append(user_message)
        schemas.append(response_schema)
        return responses.pop(0)

    monkeypatch.setattr(llm_classifier, "_call_api", fake_call)
    monkeypatch.setattr(llm_classifier.time, "sleep", lambda _seconds: None)

    result = classify_document(
        _plain_bundle(2),
        "master",
        "run",
        "fake-key",
        max_patch_attempts=1,
        max_response_attempts=2,
    )

    assert result["apply_pStyle"] == [
        {"paragraph_index": 0, "styleId": "Body"},
        {"paragraph_index": 1, "styleId": "Body"},
    ]
    assert result["ignored_paragraphs"] == []
    assert len(prompts) == 3
    assert "Classify ONLY the following paragraph indices: [1]" in prompts[1]
    assert prompts[2].startswith(prompts[1])
    assert "RETRY REQUIREMENT" in prompts[2]
    assert set(schemas[0]["properties"]) == {
        "create_styles",
        "apply_pStyle",
        "ignored_paragraphs",
        "roles",
        "notes",
    }
    assert all(
        set(schema["properties"]) == {"apply_pStyle", "ignored_paragraphs"}
        for schema in schemas[1:]
    )


@pytest.mark.parametrize(
    "patch,match",
    [
        (
            {"apply_pStyle": [], "ignored_paragraphs": [], "notes": []},
            "must contain exactly",
        ),
        (
            {
                "apply_pStyle": [{"paragraph_index": 99, "styleId": "Body"}],
                "ignored_paragraphs": [],
            },
            "unique missing paragraph",
        ),
        (
            {
                "apply_pStyle": [{"paragraph_index": 1, "styleId": "Invented"}],
                "ignored_paragraphs": [],
            },
            "not declared in roles",
        ),
    ],
)
def test_targeted_patch_contract_rejects_untrusted_fields_and_indices(patch, match):
    with pytest.raises(ValueError, match=match):
        _validate_patch_result(patch, [1], _body_instructions([0]))


def test_incomplete_patches_fail_closed_without_nearest_neighbor_fill(monkeypatch) -> None:
    _install_fake_anthropic(monkeypatch)
    bundle = _plain_bundle(2)
    responses = [
        json.dumps(_body_instructions([0])),
        '{"apply_pStyle": [], "ignored_paragraphs": []}',
        '{"apply_pStyle": [], "ignored_paragraphs": []}',
    ]
    prompts = []

    def fake_call(
        _client, _system, user_message, _model, max_tokens=128000, response_schema=None
    ):
        prompts.append(user_message)
        return responses.pop(0)

    monkeypatch.setattr(llm_classifier, "_call_api", fake_call)

    with pytest.raises(
        ValueError,
        match=r"remained incomplete after 2 targeted patch attempt\(s\).*\[1\]",
    ):
        classify_document(bundle, "master", "run", "fake-key", max_patch_attempts=2)

    assert len(prompts) == 3
    assert all("neighboring style" in prompt for prompt in prompts[1:])


def test_call_api_retries_connection_and_rate_limit_errors(monkeypatch) -> None:
    anthropic = _install_fake_anthropic(monkeypatch)
    sleeps = []

    class FinalStream:
        def __enter__(self):
            return self

        def __exit__(self, *_args):
            return False

        def get_final_text(self):
            return "success"

        def get_final_message(self):
            return types.SimpleNamespace(stop_reason="end_turn")

    class Messages:
        def __init__(self):
            self.outcomes = [
                anthropic.APIConnectionError("offline"),
                anthropic.RateLimitError("slow down"),
                FinalStream(),
            ]

        def stream(self, **_kwargs):
            outcome = self.outcomes.pop(0)
            if isinstance(outcome, Exception):
                raise outcome
            return outcome

    client = types.SimpleNamespace(messages=Messages())
    monkeypatch.setattr(llm_classifier.time, "sleep", sleeps.append)

    assert _call_api(client, "system", "user", "model") == "success"
    assert sleeps == [2, 4]


def test_call_api_requests_phase1_json_schema(monkeypatch) -> None:
    _install_fake_anthropic(monkeypatch)
    captured = {}

    class FinalStream:
        def __enter__(self):
            return self

        def __exit__(self, *_args):
            return False

        def get_final_text(self):
            return "success"

        def get_final_message(self):
            return types.SimpleNamespace(stop_reason="end_turn")

    class Messages:
        def stream(self, **kwargs):
            captured.update(kwargs)
            return FinalStream()

    client = types.SimpleNamespace(messages=Messages())

    assert _call_api(client, "system", "user", "model") == "success"
    output_config = captured["output_config"]
    assert output_config["effort"] == "high"
    assert output_config["format"]["type"] == "json_schema"
    schema = output_config["format"]["schema"]
    assert schema["type"] == "object"
    assert schema["additionalProperties"] is False
    assert set(schema["properties"]) == {
        "create_styles",
        "apply_pStyle",
        "ignored_paragraphs",
        "roles",
        "notes",
    }
    assert set(schema["required"]) == set(schema["properties"])
    assert set(schema["properties"]["roles"]["properties"]) == set(
        llm_classifier.ROLE_ORDER
    )


@pytest.mark.parametrize(
    ("stop_reason", "match"),
    [
        ("max_tokens", "output-token limit.*stop_reason=max_tokens"),
        ("refusal", "refused.*stop_reason=refusal"),
    ],
)
def test_call_api_reports_non_json_stop_reasons_without_retry(
    monkeypatch,
    stop_reason: str,
    match: str,
) -> None:
    _install_fake_anthropic(monkeypatch)
    calls = 0

    class FinalStream:
        def __enter__(self):
            return self

        def __exit__(self, *_args):
            return False

        def get_final_text(self):
            return "not a complete structured response"

        def get_final_message(self):
            return types.SimpleNamespace(stop_reason=stop_reason)

    class Messages:
        def stream(self, **_kwargs):
            nonlocal calls
            calls += 1
            return FinalStream()

    client = types.SimpleNamespace(messages=Messages())

    with pytest.raises(ValueError, match=match):
        _call_api(client, "system", "user", "model")
    assert calls == 1


def test_call_api_does_not_retry_nontransient_status(monkeypatch) -> None:
    anthropic = _install_fake_anthropic(monkeypatch)
    calls = 0
    sleeps = []

    class Messages:
        def stream(self, **_kwargs):
            nonlocal calls
            calls += 1
            raise anthropic.APIStatusError("bad request", status_code=400)

    client = types.SimpleNamespace(messages=Messages())
    monkeypatch.setattr(llm_classifier.time, "sleep", sleeps.append)

    with pytest.raises(anthropic.APIStatusError, match="bad request"):
        _call_api(client, "system", "user", "model")
    assert calls == 1
    assert sleeps == []


def test_more_than_500_small_paragraphs_reach_classifier(monkeypatch) -> None:
    _install_fake_anthropic(monkeypatch)
    bundle = _plain_bundle(501)
    instructions = _body_instructions(range(501))
    calls = []

    def fake_call(
        _client, _system, user_message, _model, max_tokens=128000, response_schema=None
    ):
        calls.append(user_message)
        return json.dumps(instructions)

    monkeypatch.setattr(llm_classifier, "_call_api", fake_call)

    result = classify_document(bundle, "master", "run", "fake-key", max_patch_attempts=0)

    assert len(calls) == 1
    assert len(result["apply_pStyle"]) == 501


def test_classifier_disables_hidden_sdk_retries(monkeypatch) -> None:
    anthropic = _install_fake_anthropic(monkeypatch)
    captured = {}

    def client_factory(**kwargs):
        captured.update(kwargs)
        return object()

    anthropic.Anthropic = client_factory
    bundle = _plain_bundle(1)
    monkeypatch.setattr(
        llm_classifier,
        "_call_api",
        lambda *_args, **_kwargs: json.dumps(_body_instructions([0])),
    )

    classify_document(bundle, "master", "run", "fake-key", max_patch_attempts=0)

    assert captured["max_retries"] == 0
    timeout = captured["timeout"]
    assert timeout.connect == 5.0
    assert timeout.read == 600.0
    assert timeout.write == 600.0
    assert timeout.pool == 600.0
