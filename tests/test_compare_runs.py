"""Exercise the offline comparison CLI with hand-written run artifacts."""

import json
import subprocess
import sys
from pathlib import Path

import pytest


SCRIPT = Path(__file__).resolve().parents[1] / "scripts" / "compare_runs.py"
HASH_A = "a" * 64
HASH_B = "b" * 64
SECRET = "CONFIDENTIAL paragraph text must never be reported"
USAGE = {
    "input_tokens": 100,
    "output_tokens": 200,
    "cache_read_input_tokens": 300,
    "cache_creation_input_tokens": 400,
    "usage_complete": True,
    "requests_attempted": 1,
}


def _write_run(path, *, role="PART", usage=None, digest=HASH_A, ignored=True, extra=False):
    path.mkdir()
    audit_name = "target-0001.audit.json"
    audit = {
        "source": {"sha256": digest, "path": str(path / "source-no-longer-exists.docx")},
        "disposition_counts": {"unresolved": 0},
        "application_audit": {
            "classifications": [{"paragraph_index": 7, "csi_role": role, "text": SECRET}],
            "ignored_paragraphs": [{"paragraph_index": 8, "reason": SECRET}] if ignored else [],
            "out_of_scope": [{"paragraph_index": 9, "text": SECRET}],
        },
        "diagnostics": [{"component": "target", "event": "classify",
                         "fields": dict(USAGE if usage is None else usage)}],
    }
    (path / audit_name).write_text(json.dumps(audit), encoding="utf-8")
    manifest = {
        "models": {"target": "claude-sonnet-5-5", "target_effort": "medium"},
        # Deliberately no matching phase events here: warning-level runs still
        # compare from full audits, and architect totals must not be charged.
        "diagnostics": {"usage": {"template": {"output_tokens": 99999}}},
        "targets": [{"source_sha256": digest, "source_path": SECRET,
                     "audit_path": "C:\\old\\run\\" + audit_name}],
    }
    if extra:
        manifest["targets"].append({"source_sha256": HASH_B, "audit_path": None})
    (path / "run.json").write_text(json.dumps(manifest), encoding="utf-8")
    return path


def _run(*paths, as_json=True):
    # -S disables site packages, proving this works without formatter or SDK
    # dependencies even when invoked outside the repository.
    return subprocess.run(
        [sys.executable, "-S", str(SCRIPT), *(str(path) for path in paths),
         *(["--json"] if as_json else [])],
        capture_output=True, text=True, cwd=paths[0].parent if paths else SCRIPT.parent,
    )


def _report(*paths):
    result = _run(*paths)
    assert result.returncode == 0, result.stderr
    assert SECRET not in result.stdout
    return json.loads(result.stdout)


def test_three_runs_compare_by_source_hash_and_audit_dispositions(tmp_path):
    a = _write_run(tmp_path / "renamed-high")
    b = _write_run(tmp_path / "renamed-medium", usage={**USAGE, "output_tokens": 25})
    c = _write_run(tmp_path / "renamed-low", role="ARTICLE", ignored=False)
    report = _report(a, b, c)
    assert len(report["targets"]) == 1
    target = report["targets"][0]
    assert target["source_sha256"] == HASH_A
    assert [(x["runs"], x["agreed"], x["total"], x["exact"]) for x in target["agreement"]] == [
        ([1, 2], 2, 2, True), ([1, 3], 0, 2, False), ([2, 3], 0, 2, False),
    ]
    assert target["agreement"][1]["differing_paragraph_indices"] == [7, 8]
    assert target["usage"][0] == {key: value for key, value in USAGE.items()
                                  if key != "requests_attempted"}
    assert target["usage"][1]["output_tokens"] == 25
    human = _run(a, b, c, as_json=False)
    assert human.returncode == 0
    assert "Agreement 1/2: 2/2" in human.stdout
    assert "cache_creation_input_tokens=400" in human.stdout
    assert SECRET not in human.stdout


def test_role_to_ignored_disagreement_ignores_reasons_and_text(tmp_path):
    a = _write_run(tmp_path / "a")
    b = _write_run(tmp_path / "b")
    audit_path = b / "target-0001.audit.json"
    audit = json.loads(audit_path.read_text())
    audit["application_audit"]["classifications"] = []
    audit["application_audit"]["ignored_paragraphs"].append({"paragraph_index": 7, "reason": "other"})
    audit_path.write_text(json.dumps(audit))
    agreement = _report(a, b)["targets"][0]["agreement"][0]
    assert agreement["agreed"] == 1
    assert agreement["total"] == 2
    assert agreement["differing_paragraph_indices"] == [7]


def test_missing_targets_audits_and_unknown_usage_are_visible(tmp_path):
    a = _write_run(tmp_path / "a", extra=True)
    b = _write_run(tmp_path / "b", usage={"input_tokens": 10, "usage_complete": False})
    report = _report(a, b)
    usage = report["targets"][0]["usage"][1]
    assert usage["input_tokens"] == 10
    assert usage["output_tokens"] is None
    assert usage["cache_creation_input_tokens"] is None
    assert usage["usage_complete"] is False
    unmatched = report["targets"][1]
    assert unmatched["usage"][1] is None
    assert unmatched["agreement"][0]["available"] is False
    (b / "target-0001.audit.json").unlink()
    report = _report(a, b)
    assert report["targets"][0]["agreement"][0]["available"] is False
    assert report["targets"][0]["usage"][1]["input_tokens"] is None


@pytest.mark.parametrize("usage, expected", [
    ({"requests_attempted": 0, "usage_complete": True}, [0, 0, 0, 0]),
    ({"input_tokens": 1, "output_tokens": 2, "usage_complete": True}, [1, 2, 0, 0]),
    ({"input_tokens": True, "output_tokens": "2", "usage_complete": False}, [None] * 4),
])
def test_zero_optional_cache_and_invalid_counters(tmp_path, usage, expected):
    a = _write_run(tmp_path / "a")
    b = _write_run(tmp_path / "b", usage=usage)
    result = _report(a, b)["targets"][0]["usage"][1]
    assert [result[key] for key in USAGE if key not in ("usage_complete", "requests_attempted")] == expected


@pytest.mark.parametrize("damage", ["unresolved", "duplicate_index", "missing_dispositions"])
def test_incomplete_dispositions_do_not_claim_agreement(tmp_path, damage):
    a = _write_run(tmp_path / "a")
    b = _write_run(tmp_path / "b")
    path = b / "target-0001.audit.json"
    audit = json.loads(path.read_text())
    if damage == "unresolved":
        audit["disposition_counts"]["unresolved"] = 1
    elif damage == "duplicate_index":
        audit["application_audit"]["ignored_paragraphs"].append({"paragraph_index": 7})
    else:
        audit["application_audit"] = {}
    path.write_text(json.dumps(audit))
    assert _report(a, b)["targets"][0]["agreement"][0]["available"] is False


def test_requires_two_runs_and_rejects_mismatched_audit_identity(tmp_path):
    a = _write_run(tmp_path / "a")
    assert _run(a).returncode != 0
    b = _write_run(tmp_path / "b")
    path = b / "target-0001.audit.json"
    audit = json.loads(path.read_text())
    audit["source"]["sha256"] = HASH_B
    path.write_text(json.dumps(audit))
    result = _run(a, b)
    assert result.returncode != 0
    assert "SHA-256 mismatch" in result.stderr


def test_duplicate_target_hashes_are_not_silently_overwritten(tmp_path):
    a = _write_run(tmp_path / "a")
    b = _write_run(tmp_path / "b")
    path = b / "run.json"
    manifest = json.loads(path.read_text())
    manifest["targets"].append(dict(manifest["targets"][0]))
    path.write_text(json.dumps(manifest))
    result = _run(a, b)
    assert result.returncode != 0
    assert "Duplicate target source SHA-256" in result.stderr
