"""Canadian architect failures must precede every target's classifier work."""

from __future__ import annotations

import json
import re
import zipfile
from pathlib import Path

import pytest

from spec_formatter import builtin_scheme, pipeline
from spec_formatter.style_application import batch_runner
from spec_formatter.style_application.core.classification import coerce_to_final_classifications
from spec_formatter.style_application.core.csi_to_canadian import plan_csi_to_canadian
from spec_formatter.style_application.core.errors import EngineError, ERROR_REMEDIATIONS
from spec_formatter.style_application.core.registry import preflight_validate_canadian_architect
from tests import test_unified_roundtrip as roundtrip


@pytest.mark.parametrize("role", ["ARTICLE", "PARAGRAPH"])
@pytest.mark.parametrize(
    ("field", "value"),
    [
        ("numbering_provenance", "literal"),
        ("numFmt", "lowerLetter"),
        ("lvlText", "%1."),
        ("start", "2"),
        ("startOverride", "2"),
        ("lvlRestart", "0"),
        ("ilvl", "8"),
        ("numId", "999"),
    ],
)
def test_common_role_contract_uses_existing_strict_rules(role, field, value):
    specs = builtin_scheme.build_role_specs()
    if field == "numbering_provenance":
        specs[role][field] = value
    else:
        specs[role]["numbering_pattern"][field] = value

    with pytest.raises(EngineError) as caught:
        preflight_validate_canadian_architect(specs, builtin_scheme.build_env_registry())
    assert caught.value.code == "canadian_architect_contract"


@pytest.mark.parametrize("role", ["PART", "ARTICLE", "PARAGRAPH"])
def test_common_hierarchy_requires_complete_role_contracts(role):
    specs = builtin_scheme.build_role_specs()
    del specs[role]
    with pytest.raises(EngineError) as caught:
        preflight_validate_canadian_architect(specs, builtin_scheme.build_env_registry())
    assert caught.value.code == "canadian_architect_contract"


@pytest.mark.parametrize("defect", ["missing", "start", "restart", "override", "list"])
def test_architect_numbering_failures_have_the_initialization_contract_code(defect):
    registry = builtin_scheme.build_env_registry()
    xml = registry["numbering"]["numbering_xml"]
    if defect == "missing":
        xml = ""
    elif defect == "start":
        xml = re.sub(r'(<w:lvl w:ilvl="1">\s*<w:start w:val=)"1"',
                     r'\g<1>"5"', xml)
    elif defect == "restart":
        xml = xml.replace('<w:lvl w:ilvl="2">',
                          '<w:lvl w:ilvl="2"><w:lvlRestart w:val="1"/>')
    elif defect == "override":
        xml = xml.replace('</w:num>',
                          '<w:lvlOverride w:ilvl="2"><w:startOverride w:val="5"/>'
                          '</w:lvlOverride></w:num>')
    else:
        xml = xml.replace('<w:num w:numId="900">', '<w:num w:numId="999">')
    registry["numbering"]["numbering_xml"] = xml

    with pytest.raises(EngineError) as caught:
        preflight_validate_canadian_architect(builtin_scheme.build_role_specs(), registry)
    assert caught.value.code == "canadian_architect_contract"


def test_deeper_roles_still_fail_only_when_the_target_uses_them():
    specs = builtin_scheme.build_role_specs()
    specs["SUBPARAGRAPH"]["numbering_pattern"]["numFmt"] = "lowerLetter"
    preflight_validate_canadian_architect(specs, builtin_scheme.build_env_registry())

    document = roundtrip._document_xml(architect=False)
    document = document.replace("Architect paragraph one", "A. Work Included")
    document = document.replace("Architect paragraph two", "1. Pumps")
    with pytest.raises(EngineError) as caught:
        plan_csi_to_canadian(
            document,
            f'<w:styles xmlns:w="{roundtrip.W_NS}"/>',
            {"classifications": [{"paragraph_index": 2, "csi_role": "SUBPARAGRAPH"}]},
            specs,
        )
    assert caught.value.code == "canadian_architect_contract"


class CountingClassifier:
    def __init__(self):
        self.calls = 0

    def __call__(self, *, slim_bundle, available_roles, **_kwargs):
        self.calls += 1
        return coerce_to_final_classifications(
            slim_bundle,
            {
                "classifications": [
                    {"paragraph_index": item["paragraph_index"], "csi_role": "PARAGRAPH"}
                    for item in slim_bundle.get("paragraphs", [])
                ],
                "ignored_paragraphs": [],
                "notes": [],
            },
            available_roles,
        )


def _run(tmp_path: Path, architect: Path, targets: list[Path], **kwargs):
    return pipeline.format_specifications(
        architect,
        targets,
        tmp_path / "formatted",
        cache_dir=tmp_path / "template-cache",
        api_key="offline-test-key",
        conversion_mode=pipeline.CSI_TO_CANADIAN,
        template_classifier=roundtrip._deterministic_classifier,
        **kwargs,
    )


@pytest.mark.parametrize("defect", ["missing_article", "paragraph_start", "article_restart"])
def test_broken_template_fails_initialization_with_zero_classifier_calls(
    tmp_path, monkeypatch, defect
):
    architect, target = roundtrip._write_canadian_pair(tmp_path)
    with zipfile.ZipFile(architect) as package:
        document = package.read("word/document.xml").decode("utf-8")
        numbering = package.read("word/numbering.xml").decode("utf-8")
    if defect == "missing_article":
        # Preserve the former two-role fixture: targets using only its deeper
        # roles could convert, but the common ARTICLE contract is absent.
        document = re.sub(r"(<w:body>)(?:<w:p>.*?</w:p>){2}", r"\1", document, count=1)
    elif defect == "paragraph_start":
        numbering = numbering.replace('<w:lvl w:ilvl="2"><w:start w:val="1"/>',
                                      '<w:lvl w:ilvl="2"><w:start w:val="5"/>')
    else:
        numbering = numbering.replace('<w:lvl w:ilvl="1">',
                                      '<w:lvl w:ilvl="1"><w:lvlRestart w:val="1"/>')
    roundtrip._rewrite_docx_parts(
        architect, {"word/document.xml": document, "word/numbering.xml": numbering}
    )
    second = tmp_path / "second.docx"
    second.write_bytes(target.read_bytes())
    classifier = CountingClassifier()
    monkeypatch.setattr(batch_runner, "classify_target_document", classifier)
    events = []

    with pytest.raises(EngineError) as caught:
        _run(tmp_path, architect, [target, second], progress=events.append)

    assert caught.value.code == "canadian_architect_contract"
    assert classifier.calls == 0
    assert not any(event.startswith("Queued ") for event in events)
    manifest = json.loads(caught.value.manifest_path.read_text(encoding="utf-8"))
    assert manifest["status"] == "failed"
    assert manifest["failure_phase"] == "initialization"
    assert manifest["error_code"] == "canadian_architect_contract"
    assert manifest["summary"]["failed"] == 2
    remediation = ERROR_REMEDIATIONS["canadian_architect_contract"]
    assert manifest["error"] == remediation
    log = (caught.value.run_dir / "run.log").read_text(encoding="utf-8")
    assert remediation in log
    assert not list(caught.value.run_dir.glob("*.docx"))
    for record in manifest["targets"]:
        assert record["stage"] == "not_started"
        assert record["error_code"] == "canadian_architect_contract"
        audit = json.loads(Path(record["audit_path"]).read_text(encoding="utf-8"))
        assert audit["phase"] == audit["stage"] == "not_started"
        assert audit["error_code"] == "canadian_architect_contract"
        assert audit["error"] == remediation


def test_valid_template_checks_once_before_multiple_target_classifiers(tmp_path, monkeypatch):
    architect, target = roundtrip._write_canadian_pair(tmp_path)
    second = tmp_path / "second.docx"
    second.write_bytes(target.read_bytes())
    classifier = CountingClassifier()
    monkeypatch.setattr(batch_runner, "classify_target_document", classifier)
    preflight_calls = []

    def counting_preflight(*args):
        assert classifier.calls == 0
        preflight_calls.append(args)
        preflight_validate_canadian_architect(*args)

    monkeypatch.setattr(pipeline, "preflight_validate_canadian_architect", counting_preflight)
    run = _run(tmp_path, architect, [target, second])
    assert run.success
    assert len(preflight_calls) == 1
    assert classifier.calls == 2
    assert len(run.output_paths) == 2
