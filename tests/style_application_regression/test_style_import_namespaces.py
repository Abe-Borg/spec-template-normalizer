"""Style import and docDefaults application across extension namespaces.

Every document current Word creates writes ``w14:ligatures`` into
``docDefaults/rPrDefault/rPr``, and styles can carry ``w14:`` children of their
own. The portable stylesheet carries the architect's styles verbatim, so those
children reach the target engine. These tests pin down the three things that
must then hold:

* materialization reads property children lexically, keyed by qualified
  name, and carries extension children after the ``w:`` children;
* every part a fragment is written into declares the prefixes the fragment
  uses, with the meaning they had in the architect (and the ``mc:Ignorable``
  tokens the architect gave them);
* a prefix the target already binds to a different namespace fails closed
  with ``style_import_namespace_conflict`` before anything is written.
"""

from __future__ import annotations

import importlib.util
import io
import re
import xml.etree.ElementTree as ET
from contextlib import redirect_stdout
from pathlib import Path

import pytest

from spec_formatter.style_application.core.errors import EngineError
from spec_formatter.style_application.core.style_import import (
    extract_style_block_raw,
    import_arch_styles_into_target,
)

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
W14_NS = "http://schemas.microsoft.com/office/word/2010/wordml"
MC_NS = "http://schemas.openxmlformats.org/markup-compatibility/2006"
# A binding no Word version writes: the same prefix, a different namespace.
FOREIGN_NS = "urn:spec-template-normalizer:test:not-word-2010"
# A synthetic extension namespace for pPr children, so the test does not
# depend on the semantics of any real Word extension element.
EXT_NS = "urn:spec-template-normalizer:test:extension"

LIGATURES = '<w14:ligatures w14:val="standardContextual"/>'
REPO_ROOT = Path(__file__).resolve().parents[2]


def _target(tmp_path: Path, styles_xml: str) -> Path:
    word_dir = tmp_path / "word"
    word_dir.mkdir()
    styles_path = word_dir / "styles.xml"
    styles_path.write_text(styles_xml, encoding="utf-8")
    return styles_path


def _root_tag(xml_text: str) -> str:
    match = re.search(r"<w:styles\b[^>]*>", xml_text)
    assert match is not None, "no w:styles root"
    return match.group(0)


def _rpr_inner(style_block: str) -> str:
    match = re.search(r"<w:rPr>([\s\S]*?)</w:rPr>", style_block)
    assert match is not None, style_block
    return match.group(1)


BARE_TARGET_STYLES = (
    f'<w:styles xmlns:w="{W_NS}">'
    '<w:style w:type="paragraph" w:default="1" w:styleId="Normal">'
    '<w:name w:val="Normal"/></w:style>'
    "</w:styles>"
)

# docDefaults carries the ligatures between two w: children, so the test can
# tell "carried after the w: children" apart from "carried where it was".
ARCHITECT_DOCDEFAULT_LIGATURES = (
    f'<w:styles xmlns:w="{W_NS}" xmlns:w14="{W14_NS}">'
    "<w:docDefaults><w:rPrDefault><w:rPr>"
    f'<w:rFonts w:ascii="Aptos"/>{LIGATURES}<w:sz w:val="24"/>'
    "</w:rPr></w:rPrDefault>"
    '<w:pPrDefault><w:pPr><w:spacing w:after="160"/></w:pPr></w:pPrDefault>'
    "</w:docDefaults>"
    '<w:style w:type="paragraph" w:default="1" w:styleId="Normal">'
    '<w:name w:val="Normal"/></w:style>'
    '<w:style w:type="paragraph" w:styleId="Role"><w:name w:val="Role"/>'
    '<w:basedOn w:val="Normal"/><w:pPr><w:keepNext/></w:pPr>'
    "<w:rPr><w:b/></w:rPr></w:style>"
    "</w:styles>"
)


# ── (a) Format-only body roots ─────────────────────────────────────────────


def test_format_only_body_root_carries_docdefault_ligatures_and_declares_w14(tmp_path):
    styles_path = _target(tmp_path, BARE_TARGET_STYLES)
    log: list = []

    result = import_arch_styles_into_target(
        tmp_path,
        ARCHITECT_DOCDEFAULT_LIGATURES,
        ["Role"],
        log,
        format_only_body_style_ids={"Role"},
    )

    assert result.declared_namespace_prefixes == ("w14",)
    assert (
        "Added XML namespace declarations to the target styles.xml root "
        "for imported architect styles: w14"
    ) in log
    out = styles_path.read_text(encoding="utf-8")
    block = extract_style_block_raw(out, result.body_style_id_map["Role"])
    assert block is not None
    # First-seen order for w: children (the role's own w:b, then the
    # defaults), the extension child after every w: child, and every child
    # carried as the source wrote it.
    assert _rpr_inner(block) == (
        f'<w:b/><w:rFonts w:ascii="Aptos"/><w:sz w:val="24"/>{LIGATURES}'
    )
    assert "<w:basedOn" not in block
    assert f'xmlns:w14="{W14_NS}"' in _root_tag(out)
    root = ET.fromstring(out.encode("utf-8"))
    clone = root.find(
        f"{{{W_NS}}}style[@{{{W_NS}}}styleId='{result.body_style_id_map['Role']}']"
    )
    assert clone is not None
    ligatures = clone.find(f"{{{W_NS}}}rPr/{{{W14_NS}}}ligatures")
    assert ligatures is not None
    assert ligatures.get(f"{{{W14_NS}}}val") == "standardContextual"


def test_format_only_body_root_carries_extension_ppr_children_after_w_children(tmp_path):
    styles_path = _target(tmp_path, BARE_TARGET_STYLES)
    architect = (
        f'<w:styles xmlns:w="{W_NS}" xmlns:x1="{EXT_NS}">'
        '<w:style w:type="paragraph" w:styleId="Base">'
        '<w:pPr><x1:marker x1:val="1"/><w:spacing w:before="120"/></w:pPr></w:style>'
        '<w:style w:type="paragraph" w:styleId="Role"><w:basedOn w:val="Base"/>'
        '<w:pPr><w:ind w:left="720"/></w:pPr></w:style>'
        "</w:styles>"
    )

    result = import_arch_styles_into_target(
        tmp_path, architect, ["Role"], [], format_only_body_style_ids={"Role"}
    )

    out = styles_path.read_text(encoding="utf-8")
    block = extract_style_block_raw(out, result.body_style_id_map["Role"])
    assert block is not None
    assert (
        '<w:pPr><w:ind w:left="720"/><w:spacing w:before="120"/>'
        '<x1:marker x1:val="1"/></w:pPr>'
    ) in block
    assert f'xmlns:x1="{EXT_NS}"' in _root_tag(out)
    ET.fromstring(out.encode("utf-8"))


# ── (b) Canadian shell clones ──────────────────────────────────────────────


def test_canadian_shell_clone_of_normal_with_w14_declares_it_in_a_bare_target(tmp_path):
    styles_path = _target(tmp_path, BARE_TARGET_STYLES)
    architect = ARCHITECT_DOCDEFAULT_LIGATURES.replace(
        '<w:name w:val="Normal"/></w:style>',
        f'<w:name w:val="Normal"/><w:rPr>{LIGATURES}</w:rPr></w:style>',
        1,
    )

    result = import_arch_styles_into_target(tmp_path, architect, ["Role"], [])

    out = styles_path.read_text(encoding="utf-8")
    normal_clone = extract_style_block_raw(out, result.style_id_map["Normal"])
    assert normal_clone is not None
    # The minimal typography injected into the clone goes ahead of the
    # extension child it already carried, not after it.
    assert _rpr_inner(normal_clone) == (
        f'<w:rFonts w:ascii="Aptos"/><w:sz w:val="24"/>{LIGATURES}'
    )
    # The target's own Normal is untouched.
    assert extract_style_block_raw(out, "Normal") == (
        '<w:style w:type="paragraph" w:default="1" w:styleId="Normal">'
        '<w:name w:val="Normal"/></w:style>\n'
    )
    assert f'xmlns:w14="{W14_NS}"' in _root_tag(out)
    ET.fromstring(out.encode("utf-8"))


def test_imports_that_use_no_new_prefix_leave_the_target_root_alone(tmp_path):
    styles_path = _target(tmp_path, BARE_TARGET_STYLES)
    architect = (
        f'<w:styles xmlns:w="{W_NS}" xmlns:w14="{W14_NS}" xmlns:mc="{MC_NS}" '
        'mc:Ignorable="w14">'
        '<w:style w:type="paragraph" w:styleId="Role"><w:rPr><w:b/></w:rPr></w:style>'
        "</w:styles>"
    )

    log: list = []
    result = import_arch_styles_into_target(tmp_path, architect, ["Role"], log)

    out = styles_path.read_text(encoding="utf-8")
    assert _root_tag(out) == f'<w:styles xmlns:w="{W_NS}">'
    assert result.declared_namespace_prefixes == ()
    assert not any("XML namespace" in line for line in log)


# ── (c) a conflicting binding fails closed ─────────────────────────────────


@pytest.mark.parametrize("body_roots", [{"Role"}, None], ids=["format_only", "canadian"])
def test_target_binding_w14_to_another_namespace_fails_closed(tmp_path, body_roots):
    target = BARE_TARGET_STYLES.replace(
        f'<w:styles xmlns:w="{W_NS}">',
        f'<w:styles xmlns:w="{W_NS}" xmlns:w14="{FOREIGN_NS}">',
    )
    styles_path = _target(tmp_path, target)
    architect = ARCHITECT_DOCDEFAULT_LIGATURES.replace(
        '<w:rPr><w:b/></w:rPr>', f"<w:rPr><w:b/>{LIGATURES}</w:rPr>", 1
    )

    with pytest.raises(EngineError) as raised:
        import_arch_styles_into_target(
            tmp_path, architect, ["Role"], [], format_only_body_style_ids=body_roots
        )

    assert raised.value.code == "style_import_namespace_conflict"
    assert "w14" in str(raised.value)
    assert styles_path.read_text(encoding="utf-8") == target


def test_a_prefix_the_architect_root_does_not_declare_fails_closed(tmp_path):
    # Declared only on the style element: resolvable in the architect, but a
    # detached clone takes the child out of that scope.
    styles_path = _target(tmp_path, BARE_TARGET_STYLES)
    architect = (
        f'<w:styles xmlns:w="{W_NS}">'
        f'<w:style xmlns:w14="{W14_NS}" w:type="paragraph" w:styleId="Base">'
        f"<w:rPr>{LIGATURES}</w:rPr></w:style>"
        '<w:style w:type="paragraph" w:styleId="Role"><w:basedOn w:val="Base"/></w:style>'
        "</w:styles>"
    )

    with pytest.raises(EngineError) as raised:
        import_arch_styles_into_target(
            tmp_path, architect, ["Role"], [], format_only_body_style_ids={"Role"}
        )

    assert raised.value.code == "style_import_namespace_conflict"
    assert styles_path.read_text(encoding="utf-8") == BARE_TARGET_STYLES


# ── (d) mc:Ignorable follows the architect ─────────────────────────────────

IGNORABLE_ARCHITECT = ARCHITECT_DOCDEFAULT_LIGATURES.replace(
    f'<w:styles xmlns:w="{W_NS}" xmlns:w14="{W14_NS}">',
    f'<w:styles xmlns:w="{W_NS}" xmlns:w14="{W14_NS}" xmlns:mc="{MC_NS}" '
    'mc:Ignorable="w14">',
)


def test_ignorable_w14_gains_mc_declaration_and_token_in_the_target(tmp_path):
    styles_path = _target(tmp_path, BARE_TARGET_STYLES)

    result = import_arch_styles_into_target(
        tmp_path, IGNORABLE_ARCHITECT, ["Role"], [], format_only_body_style_ids={"Role"}
    )

    assert result.declared_namespace_prefixes == ("mc", "w14", "mc:Ignorable=w14")
    root_tag = _root_tag(styles_path.read_text(encoding="utf-8"))
    assert f'xmlns:w14="{W14_NS}"' in root_tag
    assert f'xmlns:mc="{MC_NS}"' in root_tag
    assert 'mc:Ignorable="w14"' in root_tag
    ET.fromstring(styles_path.read_bytes())


@pytest.mark.parametrize("mce_prefix", ["mc", "ve"])
def test_ignorable_token_joins_the_targets_existing_attribute(tmp_path, mce_prefix):
    # Word 2007 bound markup compatibility to "ve"; the attribute is found by
    # namespace, never by the literal prefix.
    w15 = "http://schemas.microsoft.com/office/word/2012/wordml"
    target = BARE_TARGET_STYLES.replace(
        f'<w:styles xmlns:w="{W_NS}">',
        f'<w:styles xmlns:w="{W_NS}" xmlns:{mce_prefix}="{MC_NS}" '
        f'xmlns:w15="{w15}" {mce_prefix}:Ignorable="w15">',
    )
    styles_path = _target(tmp_path, target)

    import_arch_styles_into_target(
        tmp_path, IGNORABLE_ARCHITECT, ["Role"], [], format_only_body_style_ids={"Role"}
    )

    root_tag = _root_tag(styles_path.read_text(encoding="utf-8"))
    assert f'{mce_prefix}:Ignorable="w15 w14"' in root_tag
    assert root_tag.count("Ignorable=") == 1
    assert root_tag.count(f'="{MC_NS}"') == 1
    ET.fromstring(styles_path.read_bytes())


# ── (e) qualified names do not shadow each other ───────────────────────────

W14_SHADOW = (
    '<w14:shadow w14:blurRad="38100" w14:dist="19050" w14:dir="2700000" '
    'w14:sx="100000" w14:sy="100000" w14:kx="0" w14:ky="0" w14:algn="tl">'
    '<w14:srgbClr w14:val="000000"/></w14:shadow>'
)


@pytest.mark.parametrize(
    ("base_rpr", "role_rpr"),
    [("<w:shadow/>", W14_SHADOW), (W14_SHADOW, "<w:shadow/>")],
    ids=["w14_in_role", "w14_in_parent"],
)
def test_w_and_w14_children_with_one_local_name_are_both_materialized(
    tmp_path, base_rpr, role_rpr
):
    # w:shadow (a character effect) and w14:shadow (a drawing shadow) share
    # the local name "shadow". Keyed by local name, the nearer one used to
    # hide the other.
    styles_path = _target(tmp_path, BARE_TARGET_STYLES)
    architect = (
        f'<w:styles xmlns:w="{W_NS}" xmlns:w14="{W14_NS}">'
        f'<w:style w:type="paragraph" w:styleId="Base"><w:rPr>{base_rpr}</w:rPr></w:style>'
        '<w:style w:type="paragraph" w:styleId="Role"><w:basedOn w:val="Base"/>'
        f"<w:rPr>{role_rpr}</w:rPr></w:style>"
        "</w:styles>"
    )

    result = import_arch_styles_into_target(
        tmp_path, architect, ["Role"], [], format_only_body_style_ids={"Role"}
    )

    out = styles_path.read_text(encoding="utf-8")
    block = extract_style_block_raw(out, result.body_style_id_map["Role"])
    assert block is not None
    assert _rpr_inner(block) == f"<w:shadow/>{W14_SHADOW}"
    ET.fromstring(out.encode("utf-8"))


# ── malformed fragments name the style, never the XML ──────────────────────


@pytest.mark.parametrize("body_roots", [{"Broken"}, None], ids=["format_only", "canadian"])
@pytest.mark.parametrize(
    "broken_rpr",
    ["<w:rPr><w:b></w:rPr>", "<w:rPr><w:b></w:i></w:rPr>"],
    ids=["unclosed", "mismatched"],
)
def test_a_malformed_style_fragment_fails_naming_the_style_id(
    tmp_path, body_roots, broken_rpr
):
    styles_path = _target(tmp_path, BARE_TARGET_STYLES)
    architect = (
        f'<w:styles xmlns:w="{W_NS}">'
        f'<w:style w:type="paragraph" w:styleId="Broken">{broken_rpr}</w:style>'
        "</w:styles>"
    )

    with pytest.raises(ValueError) as raised:
        import_arch_styles_into_target(
            tmp_path, architect, ["Broken"], [], format_only_body_style_ids=body_roots
        )

    message = str(raised.value)
    assert "'Broken'" in message
    assert "<w:b>" not in message and "</w:i>" not in message
    assert styles_path.read_text(encoding="utf-8") == BARE_TARGET_STYLES


# ── apply_doc_defaults carries the same guarantee ──────────────────────────

TARGET_WITH_DEFAULTS = (
    f'<w:styles xmlns:w="{W_NS}">'
    '<w:docDefaults><w:rPrDefault><w:rPr><w:sz w:val="20"/></w:rPr></w:rPrDefault>'
    "</w:docDefaults>"
    '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"/>'
    "</w:styles>"
)


def _registry_defaults(rpr: str) -> dict:
    return {
        "doc_defaults": {
            "default_run_props": {"rPr": rpr},
            "default_paragraph_props": {"pPr": '<w:pPr><w:spacing w:after="160"/></w:pPr>'},
        }
    }


def test_apply_doc_defaults_declares_what_the_architect_defaults_use():
    from spec_formatter.style_application.arch_env_applier import apply_doc_defaults
    from spec_formatter.style_application.core.xml_helpers import RootNamespaces

    log: list = []
    out = apply_doc_defaults(
        TARGET_WITH_DEFAULTS,
        _registry_defaults(f'<w:rPr><w:rFonts w:ascii="Aptos"/>{LIGATURES}</w:rPr>'),
        log,
        architect_namespaces=RootNamespaces.of(IGNORABLE_ARCHITECT),
    )

    assert (
        "Added XML namespace declarations to the target styles.xml root "
        "for the architect docDefaults: mc, w14, mc:Ignorable=w14"
    ) in log
    root_tag = _root_tag(out)
    assert f'xmlns:w14="{W14_NS}"' in root_tag
    assert 'mc:Ignorable="w14"' in root_tag
    root = ET.fromstring(out.encode("utf-8"))
    ligatures = root.find(
        f"{{{W_NS}}}docDefaults/{{{W_NS}}}rPrDefault/{{{W_NS}}}rPr/{{{W14_NS}}}ligatures"
    )
    assert ligatures is not None
    assert '<w:sz w:val="20"/>' not in out  # the target's defaults were replaced


def test_apply_doc_defaults_leaves_the_root_alone_when_only_w_is_used():
    from spec_formatter.style_application.arch_env_applier import apply_doc_defaults
    from spec_formatter.style_application.core.xml_helpers import RootNamespaces

    out = apply_doc_defaults(
        TARGET_WITH_DEFAULTS,
        _registry_defaults('<w:rPr><w:rFonts w:ascii="Aptos"/></w:rPr>'),
        [],
        architect_namespaces=RootNamespaces.of(IGNORABLE_ARCHITECT),
    )

    assert _root_tag(out) == f'<w:styles xmlns:w="{W_NS}">'
    ET.fromstring(out.encode("utf-8"))


def test_apply_doc_defaults_fails_closed_on_a_conflicting_binding():
    from spec_formatter.style_application.arch_env_applier import apply_doc_defaults
    from spec_formatter.style_application.core.xml_helpers import RootNamespaces

    target = TARGET_WITH_DEFAULTS.replace(
        f'<w:styles xmlns:w="{W_NS}">',
        f'<w:styles xmlns:w="{W_NS}" xmlns:w14="{FOREIGN_NS}">',
    )

    with pytest.raises(EngineError) as raised:
        apply_doc_defaults(
            target,
            _registry_defaults(f"<w:rPr>{LIGATURES}</w:rPr>"),
            [],
            architect_namespaces=RootNamespaces.of(IGNORABLE_ARCHITECT),
        )

    assert raised.value.code == "style_import_namespace_conflict"


# ── the probe is a regression test ─────────────────────────────────────────


def test_style_import_w14_probe_passes():
    probe_path = (
        REPO_ROOT / "docs" / "docx_method_hardening" / "probes" / "probe_style_import_w14.py"
    )
    spec = importlib.util.spec_from_file_location("probe_style_import_w14", probe_path)
    assert spec is not None and spec.loader is not None
    probe = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(probe)

    printed = io.StringIO()
    with redirect_stdout(printed):
        exit_code = probe.main()

    rows = printed.getvalue().splitlines()
    assert exit_code == 0, printed.getvalue()
    assert [row.split()[:2] for row in rows] == [["format_only", "OK"], ["csi_to_canadian", "OK"]]
