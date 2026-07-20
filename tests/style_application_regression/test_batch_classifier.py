import sys

from spec_formatter.style_application.core.batch_classifier import (
    build_batch_requests,
    reassemble_file_classifications,
    submit_and_poll,
)


def test_build_batch_requests_creates_chunked_custom_ids(monkeypatch):
    from spec_formatter.style_application.core import batch_classifier as bc

    monkeypatch.setattr(
        bc,
        "_split_bundle_into_chunks",
        lambda bundle: [{"paragraphs": [{"paragraph_index": 1}]}, {"paragraphs": [{"paragraph_index": 2}]}],
    )

    reqs = build_batch_requests(
        file_bundles={"my-file.docx": {"paragraphs": [{"paragraph_index": 1}]}},
        available_roles=["PART"],
        model="m",
    )

    assert [r["custom_id"] for r in reqs] == ["my-file.docx__chunk0", "my-file.docx__chunk1"]
    assert reqs[0]["params"]["model"] == "m"
    assert reqs[0]["params"]["output_config"] == {"effort": "high"}


def test_build_batch_requests_skips_deterministic_only_files(monkeypatch):
    from spec_formatter.style_application.core import batch_classifier as bc

    def should_not_split(_bundle):
        raise AssertionError("deterministic-only bundle must not be chunked")

    monkeypatch.setattr(bc, "_split_bundle_into_chunks", should_not_split)
    reqs = build_batch_requests(
        file_bundles={
            "deterministic.docx": {
                "paragraphs": [],
                "deterministic_classifications": [
                    {"paragraph_index": 0, "csi_role": "PART"}
                ],
            }
        },
        available_roles=["PART"],
        model="m",
    )
    assert reqs == []


def test_submit_and_poll_empty_requests_never_constructs_api_client(monkeypatch):
    class ExplodingAnthropicModule:
        def __getattr__(self, name):
            raise AssertionError(f"Anthropic SDK must not be accessed: {name}")

    monkeypatch.setitem(sys.modules, "anthropic", ExplodingAnthropicModule())
    assert submit_and_poll([], api_key="") == {}


def test_reassemble_file_classifications_merges_chunks(monkeypatch):
    from spec_formatter.style_application.core import batch_classifier as bc

    split_chunks = [
        {"paragraphs": [{"paragraph_index": 1}]},
        {"paragraphs": [{"paragraph_index": 2}]},
    ]
    monkeypatch.setattr(bc, "_split_bundle_into_chunks", lambda bundle: split_chunks)

    out = reassemble_file_classifications(
        results={
            "a.docx__chunk0": {"classifications": [{"paragraph_index": 1, "csi_role": "PART"}]},
            "a.docx__chunk1": {"classifications": [{"paragraph_index": 2, "csi_role": "PART"}]},
        },
        file_bundles={
            "a.docx": {
                "paragraphs": [{"paragraph_index": 1}, {"paragraph_index": 2}],
                "deterministic_classifications": [],
                "available_roles": ["PART"],
            },
            "deterministic.docx": {
                "paragraphs": [],
                "deterministic_classifications": [
                    {"paragraph_index": 9, "csi_role": "PART"}
                ],
                "available_roles": ["PART"],
            },
        },
        available_roles=["PART"],
    )

    assert "a.docx" in out
    assert len(out["a.docx"]["classifications"]) == 2
    assert out["deterministic.docx"]["classifications"] == [
        {"paragraph_index": 9, "csi_role": "PART"}
    ]


def test_batch_requests_are_unique_for_duplicate_basenames(monkeypatch):
    from spec_formatter.style_application.core import batch_classifier as bc

    monkeypatch.setattr(bc, "_split_bundle_into_chunks", lambda bundle: [{"paragraphs": [{"paragraph_index": 1}]}])
    reqs = build_batch_requests(
        file_bundles={
            "spec__aaa111": {"paragraphs": [{"paragraph_index": 1}]},
            "spec__bbb222": {"paragraphs": [{"paragraph_index": 1}]},
        },
        available_roles=["PART"],
        model="m",
    )
    custom_ids = [r["custom_id"] for r in reqs]
    assert len(custom_ids) == len(set(custom_ids))
