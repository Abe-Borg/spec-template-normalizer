"""Request compaction preserves evidence without changing the local slim bundle."""

import copy
import json
import zipfile

import pytest

from spec_formatter.style_application.core import classification as classification_module
from spec_formatter.style_application.core.classification import (
    _build_numbering_catalog,
    _resolve_numbering_pattern,
    build_phase2_slim_bundle,
    coerce_to_final_classifications,
    validate_phase2_final_payload,
    validate_phase2_llm_payload,
)
from spec_formatter.style_application.core.llm_classifier import (
    _CHUNK_OVERLAP,
    _build_user_message,
    _split_bundle_into_chunks,
    _validate_classifications,
    _wire_json,
)
from tests.fixtures.sanitized_format_only_corpus import write_sanitized_format_only_pair


@pytest.fixture
def forced_unresolved_corpus(tmp_path, monkeypatch):
    target = tmp_path / "target.docx"
    write_sanitized_format_only_pair(tmp_path / "architect.docx", target)
    extracted = tmp_path / "extracted"
    with zipfile.ZipFile(target) as package:
        package.extractall(extracted)
    monkeypatch.setattr(classification_module, "preclassify_paragraphs", lambda *_args: {})
    bundle = build_phase2_slim_bundle(extracted)
    catalog = _build_numbering_catalog((extracted / "word/numbering.xml").read_text())
    assert len(bundle["paragraphs"]) == 122  # 32 editorial/filtered source rows stay excluded.
    for row in bundle["paragraphs"]:
        assert row["numbering_pattern"] == _resolve_numbering_pattern(row["effective_numPr"], catalog)
    return bundle


def _assert_all_facts_preserved(bundle):
    original = copy.deepcopy(bundle)
    request = json.loads(_build_user_message(bundle))
    assert [row["paragraph_index"] for row in request["paragraphs"]] == [
        row["paragraph_index"] for row in bundle["paragraphs"]
    ]
    for position, (old, new) in enumerate(zip(bundle["paragraphs"], request["paragraphs"])):
        for field, value in old.items():
            if field in new:
                assert new[field] == value
            elif field == "numbering_pattern" and value:
                ref = new["effective_numPr"]
                num_id, ilvl = ref["numId"], ref.get("ilvl", "0")
                level = request["numbering_levels"][num_id][ilvl]
                assert {"numId": num_id, "ilvl": ilvl, **level} == value
            elif field in ("prev_text", "next_text") and value:
                neighbour = position + (-1 if field == "prev_text" else 1)
                assert 0 <= neighbour < len(request["paragraphs"])
                assert request["paragraphs"][neighbour]["text"][:80] == value
            else:
                assert value is None or value is False or value in ("", {}, [])
    assert bundle == original
    return request


def test_projection_preserves_every_corpus_fact_and_shares_six_levels(forced_unresolved_corpus):
    bundle = forced_unresolved_corpus
    request = _assert_all_facts_preserved(bundle)
    assert set(request["numbering_levels"]) == {"17"}
    assert set(request["numbering_levels"]["17"]) == {"0", "3", "4", "5", "6", "7"}
    for row in request["paragraphs"]:
        assert not {"prev_text", "next_text", "numbering_pattern", "in_table", "numPr"} & row.keys()
    old_message = "available_roles: " + json.dumps(bundle["available_roles"]) + "\n\n" + _wire_json({"paragraphs": bundle["paragraphs"]})
    assert len(_build_user_message(bundle)) < len(old_message) / 2


def test_projection_retains_nonempty_hints_conflicts_direct_numbering_and_zero_index():
    row = {
        "paragraph_index": 0, "text": "Requirement", "pStyle": "Base",
        "pPr_hints": {"keepNext": False, "ind": {"left": "0"}},
        "rPr_hints": {"bold": False, "color": "445566"},
        "numPr": {"ilvl": "2"}, "effective_numPr": {"numId": "7", "ilvl": "2"},
        "numbering_pattern": {
            "numId": "7", "ilvl": "2", "abstractNumId": "9", "start": "1",
            "numFmt": "decimal", "lvlText": "", "lvlRestart": "0",
            "startOverride": "4", "suff": "nothing", "isLgl": "1", "pStyle": "List",
        },
        "numbering_match_candidates": ["PARAGRAPH", "SUBPARAGRAPH"],
        "numbering_semantic_conflict": True, "numbering_role": "SUBPARAGRAPH",
        "contains_sectPr": True, "in_table": False, "prev_text": "", "next_text": "",
        "future_hint": 0,
    }
    request = _assert_all_facts_preserved({"paragraphs": [row]})
    assert request["paragraphs"][0]["future_hint"] == 0
    assert request["paragraphs"][0]["numPr"] == {"ilvl": "2"}


def test_numbering_table_uses_effective_reference_with_default_level():
    pattern = {"numId": "7", "ilvl": "0", "lvlText": "%1."}
    paragraphs = [
        {"paragraph_index": 0, "text": "Inherited", "numPr": None,
         "effective_numPr": {"numId": "7"}, "numbering_pattern": pattern},
        {"paragraph_index": 1, "text": "Direct", "numPr": {"numId": "7", "ilvl": "0"},
         "effective_numPr": {"numId": "7", "ilvl": "0"}, "numbering_pattern": pattern},
        {"paragraph_index": 2, "text": "Another list",
         "effective_numPr": {"numId": "8", "ilvl": "0"},
         "numbering_pattern": {**pattern, "numId": "8", "startOverride": "3"}},
    ]
    request = _assert_all_facts_preserved({"paragraphs": paragraphs})
    assert request["numbering_levels"] == {
        "7": {"0": {"lvlText": "%1."}},
        "8": {"0": {"lvlText": "%1.", "startOverride": "3"}},
    }


def test_bundle_patterns_derive_from_effective_numbering_with_inheritance_and_overrides(tmp_path, monkeypatch):
    namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
    word = tmp_path / "word"
    word.mkdir()
    (word / "document.xml").write_text(
        f'<w:document xmlns:w="{namespace}"><w:body>'
        '<w:p><w:pPr><w:pStyle w:val="Base"/></w:pPr><w:r><w:t>Inherited</w:t></w:r></w:p>'
        '<w:p><w:pPr><w:pStyle w:val="Base"/><w:numPr><w:ilvl w:val="1"/></w:numPr></w:pPr><w:r><w:t>Partial override</w:t></w:r></w:p>'
        '<w:p><w:pPr><w:numPr><w:numId w:val="8"/></w:numPr></w:pPr><w:r><w:t>Missing definition</w:t></w:r></w:p>'
        '<w:p><w:pPr><w:pStyle w:val="Base"/><w:numPr><w:numId w:val="0"/></w:numPr></w:pPr><w:r><w:t>Disabled</w:t></w:r></w:p>'
        '</w:body></w:document>'
    )
    (word / "styles.xml").write_text(
        f'<w:styles xmlns:w="{namespace}"><w:style w:type="paragraph" w:styleId="Base">'
        '<w:pPr><w:numPr><w:numId w:val="7"/><w:ilvl w:val="0"/></w:numPr></w:pPr></w:style></w:styles>'
    )
    numbering = (
        f'<w:numbering xmlns:w="{namespace}"><w:abstractNum w:abstractNumId="9">'
        '<w:lvl w:ilvl="0"><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/></w:lvl>'
        '<w:lvl w:ilvl="1"><w:numFmt w:val="lowerLetter"/><w:lvlText w:val="%2)"/></w:lvl>'
        '</w:abstractNum><w:num w:numId="7"><w:abstractNumId w:val="9"/>'
        '<w:lvlOverride w:ilvl="1"><w:startOverride w:val="4"/><w:lvl w:ilvl="1">'
        '<w:lvlRestart w:val="0"/><w:suff w:val="nothing"/><w:isLgl/></w:lvl></w:lvlOverride>'
        '</w:num></w:numbering>'
    )
    (word / "numbering.xml").write_text(numbering)
    monkeypatch.setattr(classification_module, "preclassify_paragraphs", lambda *_args: {})
    bundle = build_phase2_slim_bundle(tmp_path)
    assert len(bundle["paragraphs"]) == 4
    catalog = _build_numbering_catalog(numbering)
    for row in bundle["paragraphs"]:
        assert row["numbering_pattern"] == _resolve_numbering_pattern(row["effective_numPr"], catalog)
    rows = bundle["paragraphs"]
    assert rows[0]["numPr"] is None
    assert rows[1]["numPr"] == {"numId": None, "ilvl": "1"}
    assert rows[1]["numbering_pattern"]["startOverride"] == "4"
    assert rows[2]["numbering_pattern"] == {"numId": "8", "ilvl": "0"}
    assert rows[3]["numbering_pattern"] is None
    _assert_all_facts_preserved(bundle)


@pytest.mark.parametrize("mismatch", [False, True])
def test_custom_inconsistent_patterns_stay_inline(mismatch):
    rows = [
        {"paragraph_index": i, "text": f"P{i}",
         "effective_numPr": {"numId": "7", "ilvl": "0"},
         "numbering_pattern": {"numId": "8" if mismatch else "7", "ilvl": "0", "lvlText": str(i)}}
        for i in range(2)
    ]
    request = _assert_all_facts_preserved({"paragraphs": rows})
    assert "numbering_levels" not in request
    assert [row["numbering_pattern"] for row in request["paragraphs"]] == [row["numbering_pattern"] for row in rows]


def test_neighbour_context_survives_deterministic_gaps_and_chunk_boundaries():
    texts = ["First " + "a" * 100, "Resolved heading", "Requirement", "Next requirement", "Resolved end"]
    paragraphs = [
        {"paragraph_index": i, "text": text,
         "prev_text": texts[i - 1][:80] if i else "",
         "next_text": texts[i + 1][:80] if i < len(texts) - 1 else ""}
        for i, text in enumerate(texts)
    ]
    subset = {"paragraphs": [paragraphs[i] for i in (0, 2, 3)]}
    request = _assert_all_facts_preserved(subset)
    first, middle, last = request["paragraphs"]
    assert first["next_text"] == "Resolved heading"
    assert middle["prev_text"] == "Resolved heading"
    assert "next_text" not in middle
    assert "prev_text" not in last
    assert last["next_text"] == "Resolved end"
    boundary = _assert_all_facts_preserved({"paragraphs": paragraphs[2:3]})["paragraphs"][0]
    assert boundary["prev_text"] == "Resolved heading"
    assert boundary["next_text"] == "Next requirement"
    # Filtered source indices are not evidence of a missing neighbour.
    all_rows = _assert_all_facts_preserved({"paragraphs": paragraphs})["paragraphs"]
    assert "prev_text" not in all_rows[1]  # Long neighbour text is truncated to 80.


def test_chunker_fits_projected_corpus_at_exact_wire_limit(forced_unresolved_corpus):
    bundle = forced_unresolved_corpus
    wire_bytes = len(_build_user_message(bundle).encode("utf-8"))
    assert len(_wire_json({"paragraphs": bundle["paragraphs"]})) > wire_bytes
    assert _split_bundle_into_chunks(bundle, max_chars=wire_bytes) == [bundle]
    chunks = _split_bundle_into_chunks(bundle, max_chars=wire_bytes - 1)
    assert len(chunks) > 1
    for chunk in chunks:
        assert len(_build_user_message(chunk).encode("utf-8")) <= wire_bytes - 1
        _assert_all_facts_preserved(chunk)


def test_chunker_bounds_uneven_unicode_rows_and_keeps_overlap():
    paragraphs = [
        {"paragraph_index": i, "text": ("é" * 200 if i % 7 == 0 else f"P{i}"),
         "pPr_hints": {"large_hint": "x" * 1800} if i % 11 == 0 else {}}
        for i in range(95)
    ]
    for i, row in enumerate(paragraphs):
        row["prev_text"] = paragraphs[i - 1]["text"][:80] if i else ""
        row["next_text"] = paragraphs[i + 1]["text"][:80] if i < len(paragraphs) - 1 else ""
    chunks = _split_bundle_into_chunks({"paragraphs": paragraphs}, max_chars=15000)
    assert len(chunks) > 1
    covered = set()
    for chunk in chunks:
        assert len(_build_user_message(chunk).encode("utf-8")) <= 15000
        assert chunk["_chunk_info"]["paragraph_range"] == [chunk["paragraphs"][0]["paragraph_index"], chunk["paragraphs"][-1]["paragraph_index"]]
        covered.update(row["paragraph_index"] for row in chunk["paragraphs"])
        _assert_all_facts_preserved(chunk)
    assert covered == set(range(95))
    for left, right in zip(chunks, chunks[1:]):
        assert left["paragraphs"][-_CHUNK_OVERLAP:] == right["paragraphs"][:_CHUNK_OVERLAP]


def test_chunker_advances_under_tiny_budget_and_isolates_oversize_rows():
    bundle = {"paragraphs": [{"paragraph_index": i, "text": "x" * (2000 if i == 1 else 10)} for i in range(4)]}
    original = copy.deepcopy(bundle)
    chunks = _split_bundle_into_chunks(bundle, max_chars=130)
    assert [chunk["paragraphs"][0]["paragraph_index"] for chunk in chunks] == list(range(4))
    assert all(len(chunk["paragraphs"]) == 1 for chunk in chunks)
    assert bundle == original


@pytest.mark.parametrize("notes", [None, ["legacy response"]])
def test_validators_and_local_reassembly_accept_response_without_notes(notes):
    bundle = {"paragraphs": [{"paragraph_index": 1, "text": "Requirement"}],
              "deterministic_classifications": [{"paragraph_index": 0, "csi_role": "PART"}]}
    result = {"classifications": [{"paragraph_index": 1, "csi_role": "PARAGRAPH"}], "ignored_paragraphs": []}
    if notes is not None:
        result["notes"] = notes
    roles = ["PART", "PARAGRAPH"]
    assert _validate_classifications(result, roles, {1})["notes"] == (notes or [])
    validate_phase2_llm_payload(bundle, result, roles)
    final = coerce_to_final_classifications(bundle, result, roles)
    validate_phase2_final_payload(bundle, final, roles)
    assert final["notes"] == (notes or [])
    assert final["classifications"] == [{"paragraph_index": 0, "csi_role": "PART"}, {"paragraph_index": 1, "csi_role": "PARAGRAPH"}]
