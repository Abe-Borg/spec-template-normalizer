"""Every package member outside the mode's remit comes through byte-identical.

``patch_docx`` copies every source member it is not handed as a replacement,
in source order and compression. That is the transform, and until WI-05 it
was also the only thing standing behind "the architect-free modes apply no
document shell of any kind": the packaging step hands ``word/styles.xml``,
``word/settings.xml``, the theme, the font table, ``word/numbering.xml``, the
content types and the document relationships over as replacements in *every*
mode, read back from the extraction directory, so a stray write to any of
them published. ``validate_docx_package`` checks structure, never identity
with the source.

The final gate now takes a census of every member of the source and the
output (name and SHA-256 of the bytes, never decoded text) and proves each
change, addition and removal falls inside what the mode's
``ApplicationPolicy`` allows -- with the header/footer parts, their
relationships and their media additionally cross-checked against the
importer's manifest and the packages' own relationships.

The out-of-remit damage here is injected into the *packaged* output (or into
a replacement the packaging step reads), because a member outside the
replacement set is copied from the source and would otherwise never reach
the gate changed.
"""

from __future__ import annotations

import codecs
import hashlib
import json
import re
import zipfile
from dataclasses import replace
from pathlib import Path
from typing import Callable, Mapping, Optional

import pytest

from spec_formatter.pipeline import (
    CANADIAN_TO_CSI,
    CSI_TO_CANADIAN,
    CSI_TO_CANADIAN_STANDALONE,
    format_specifications,
)
from spec_formatter.style_application import batch_runner
from spec_formatter.style_application.core.application_policy import (
    application_policy_for_mode,
)
from spec_formatter.style_application.phase2_invariants import (
    _package_member_census,
    _verify_package_member_remit,
    validate_docx_package,
)
from tests import test_architect_free_modes as free
from tests import test_unified_roundtrip as roundtrip

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
CT_NS = "http://schemas.openxmlformats.org/package/2006/content-types"
OFFICE_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
FORMAT_ONLY = "format_only"
ALL_MODES = (FORMAT_ONLY, CSI_TO_CANADIAN, CSI_TO_CANADIAN_STANDALONE, CANADIAN_TO_CSI)
OUT_OF_REMIT = "package member outside the mode's remit"

FOOTNOTES = (
    f'<w:footnotes xmlns:w="{W_NS}"><w:footnote w:id="1"><w:p><w:r>'
    "<w:t>A footnote the formatter never touches</w:t></w:r></w:p></w:footnote>"
    "</w:footnotes>"
)
COMMENTS = (
    f'<w:comments xmlns:w="{W_NS}"><w:comment w:id="0" w:author="Reviewer">'
    "<w:p><w:r><w:t>A reviewer comment</w:t></w:r></w:p></w:comment></w:comments>"
)
TRASH_ITEM = b"\x00a discarded physical ZIP item\x00"

# (payload, content type Override or None, document relationship type or None)
_Part = tuple[bytes, Optional[str], Optional[str]]

UNTOUCHED_PARTS: dict[str, _Part] = {
    "word/footnotes.xml": (
        FOOTNOTES.encode("utf-8"),
        "application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml",
        f"{OFFICE_REL}/footnotes",
    ),
    "word/comments.xml": (
        COMMENTS.encode("utf-8"),
        "application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml",
        f"{OFFICE_REL}/comments",
    ),
    # An OOXML trash item: a discarded physical item with no content type
    # and no relationship, which Word keeps and so must this application.
    "[trash]/0000.dat": (TRASH_ITEM, None, None),
}


# --- Package helpers --------------------------------------------------------


def _members(docx: Path) -> dict[str, bytes]:
    with zipfile.ZipFile(docx) as package:
        return {name: package.read(name) for name in package.namelist()}


def _write_members(path: Path, members: Mapping[str, bytes | str]) -> Path:
    path.parent.mkdir(parents=True, exist_ok=True)
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as package:
        for name, payload in members.items():
            package.writestr(name, payload)
    return path


def _with_parts(docx: Path, parts: Mapping[str, _Part]) -> Path:
    """``docx`` with ``parts`` added and wired in, rewritten in place."""

    members = _members(docx)
    overrides = ""
    relationships = ""
    for name, (payload, content_type, relationship_type) in parts.items():
        members[name] = payload
        if content_type is not None:
            overrides += f'<Override PartName="/{name}" ContentType="{content_type}"/>'
        if relationship_type is not None:
            target = name[len("word/"):]
            relationships += (
                f'<Relationship Id="rIdWi05{re.sub(r"[^A-Za-z0-9]", "", name)}" '
                f'Type="{relationship_type}" '
                f'Target="{target}"/>'
            )
    members["[Content_Types].xml"] = members["[Content_Types].xml"].replace(
        b"</Types>", overrides.encode("utf-8") + b"</Types>"
    )
    members["word/_rels/document.xml.rels"] = members[
        "word/_rels/document.xml.rels"
    ].replace(b"</Relationships>", relationships.encode("utf-8") + b"</Relationships>")
    return _write_members(docx, members)


def _independent_census(source: Path, output: Path) -> dict[str, set[str]]:
    """Classify members by hash with nothing from the engine."""

    def digests(docx: Path) -> dict[str, str]:
        return {
            name: hashlib.sha256(payload).hexdigest()
            for name, payload in _members(docx).items()
        }

    before, after = digests(source), digests(output)
    return {
        "compared": set(before) | set(after),
        "changed": {name for name in set(before) & set(after) if before[name] != after[name]},
        "added": set(after) - set(before),
        "removed": set(before) - set(after),
    }


# --- Driving each mode without an API key ------------------------------------


def _format_only_run(tmp_path: Path):
    architect = tmp_path / "architect.docx"
    target = tmp_path / "target.docx"
    roundtrip._write_docx(architect, architect=True)
    roundtrip._write_docx(target, architect=False)
    _with_parts(target, UNTOUCHED_PARTS)
    return _architect_run(tmp_path, architect, target, FORMAT_ONLY), target


def _architect_canadian_run(tmp_path: Path):
    architect, target = roundtrip._write_canadian_pair(tmp_path)
    _with_parts(target, UNTOUCHED_PARTS)
    return _architect_run(tmp_path, architect, target, CSI_TO_CANADIAN), target


def _architect_run(tmp_path: Path, architect: Path, target: Path, mode: str):
    return format_specifications(
        architect_template=architect,
        target_specs=[target],
        output_dir=tmp_path / "formatted",
        cache_dir=tmp_path / "template-cache",
        api_key="",
        max_workers=1,
        conversion_mode=mode,
        template_model=f"wi05-{mode}-fixture",
        template_classifier=roundtrip._deterministic_classifier,
    )


def _standalone_source(tmp_path: Path) -> Path:
    return _with_parts(free._write_docx(tmp_path / "source" / "spec.docx"), UNTOUCHED_PARTS)


def _standalone_run(tmp_path: Path):
    source = _standalone_source(tmp_path)
    return free._run(tmp_path, source, CSI_TO_CANADIAN_STANDALONE, "canadian"), source


def _canadian_source(tmp_path: Path) -> Path:
    """The standalone Canadian output, used as the reverse mode's input.

    Built before any damage is installed, so the forward run it comes from is
    clean. The untouched parts travel through it unchanged.
    """

    forward, _source = _standalone_run(tmp_path)
    assert forward.success, "\n".join(forward.targets[0].log)
    output = forward.targets[0].output_path
    assert output is not None
    return output


def _reverse_run(tmp_path: Path, source: Optional[Path] = None):
    source = source or _canadian_source(tmp_path)
    return free._run(tmp_path, source, CANADIAN_TO_CSI, "csi"), source


_RUNNERS: dict[str, Callable] = {
    FORMAT_ONLY: _format_only_run,
    CSI_TO_CANADIAN: _architect_canadian_run,
    CSI_TO_CANADIAN_STANDALONE: _standalone_run,
    CANADIAN_TO_CSI: _reverse_run,
}


def _build_output_fields(run) -> dict:
    return roundtrip._diagnostics_event(run, "build_output")["fields"]


# --- Happy paths: what each mode legitimately changes ------------------------

# The architect's header set replaces the target's: the target's header9 goes,
# the architect's headers, footer, header relationships and logo arrive.
_ARCHITECT_HEADER_SET_ADDED = {
    "word/header1.xml",
    "word/header2.xml",
    "word/header3.xml",
    "word/footer1.xml",
    "word/_rels/header1.xml.rels",
}
_EXPECTED: dict[str, dict[str, set[str]]] = {
    # Format-only: shell (settings, content types, document relationships),
    # role styles and docDefaults in styles.xml, the body. The target's own
    # numbering is left alone: nothing from the architect's list is imported.
    FORMAT_ONLY: {
        "changed": {
            "[Content_Types].xml",
            "word/_rels/document.xml.rels",
            "word/document.xml",
            "word/settings.xml",
            "word/styles.xml",
        },
        "removed": {"word/header9.xml"},
    },
    # The architect Canadian mode also imports the architect's list.
    CSI_TO_CANADIAN: {
        "changed": {
            "[Content_Types].xml",
            "word/_rels/document.xml.rels",
            "word/document.xml",
            "word/numbering.xml",
            "word/settings.xml",
            "word/styles.xml",
        },
        "removed": {"word/header9.xml"},
    },
    # No shell and no role styles: the built-in list is added and wired in,
    # and styles.xml comes through byte-identical.
    CSI_TO_CANADIAN_STANDALONE: {
        "changed": {
            "[Content_Types].xml",
            "word/_rels/document.xml.rels",
            "word/document.xml",
        },
        "added": {"word/numbering.xml"},
        "removed": set(),
    },
    # Literal markers only: the body and nothing else.
    CANADIAN_TO_CSI: {
        "changed": {"word/document.xml"},
        "added": set(),
        "removed": set(),
    },
}


@pytest.mark.parametrize("mode", ALL_MODES)
def test_every_mode_changes_only_what_its_remit_allows(tmp_path: Path, mode: str) -> None:
    run, source = _RUNNERS[mode](tmp_path)
    assert run.success, "\n".join(run.targets[0].log)
    output = run.targets[0].output_path
    assert output is not None
    validate_docx_package(output)

    census = _independent_census(source, output)
    expected = _EXPECTED[mode]
    assert census["changed"] == expected["changed"]
    assert census["removed"] == expected["removed"]
    if mode in (FORMAT_ONLY, CSI_TO_CANADIAN):
        media = {name for name in census["added"] if name.startswith("word/media/")}
        assert len(media) == 1, census["added"]
        assert census["added"] - media == _ARCHITECT_HEADER_SET_ADDED
    else:
        assert census["added"] == expected["added"]
    # The parts no mode has any business with come through byte-identical.
    for name in UNTOUCHED_PARTS:
        assert name not in census["changed"] | census["added"] | census["removed"]

    fields = _build_output_fields(run)
    assert fields["package_members_compared"] == len(census["compared"])
    assert fields["package_members_changed"] == len(census["changed"])
    assert fields["package_members_added"] == len(census["added"])
    assert fields["package_members_removed"] == len(census["removed"])


def test_member_names_reach_the_run_log_and_no_structured_artifact(tmp_path: Path) -> None:
    run, _source = _format_only_run(tmp_path)
    assert run.success, "\n".join(run.targets[0].log)

    run_log = (run.run_dir / "run.log").read_text(encoding="utf-8")
    assert "Package member removed: word/header9.xml" in run_log
    assert "Package member changed: word/styles.xml" in run_log
    assert "Package member added: word/header1.xml" in run_log
    assert "Package members: " in run_log
    # Counts, never names, in the structured artifacts.
    fields = _build_output_fields(run)
    assert not any(
        isinstance(value, str) and "header" in value for value in fields.values()
    )
    for artifact in (run.manifest_path, run.diagnostics_path):
        assert "Package member" not in Path(artifact).read_text(encoding="utf-8")


# --- Out-of-remit damage in the packaged output ------------------------------


def _alter_packaged_output(
    monkeypatch: pytest.MonkeyPatch,
    alterations: Mapping[str, Callable[[Optional[bytes]], Optional[bytes]]],
) -> None:
    """Rewrite members of the packaged output before the gate reads it.

    A member ``patch_docx`` is not handed as a replacement is copied from the
    source, so damage to it can only be staged here: after the package is
    built, before it is validated and verified. Returning ``None`` removes the
    member.
    """

    real_patch = batch_runner.patch_docx

    def patch_then_alter(src_docx, out_docx, *args, **kwargs):
        real_patch(src_docx, out_docx, *args, **kwargs)
        with zipfile.ZipFile(out_docx) as package:
            infos = package.infolist()
            payloads = {info.filename: package.read(info) for info in infos}
        with zipfile.ZipFile(out_docx, "w") as package:
            for info in infos:
                payload = payloads[info.filename]
                alter = alterations.get(info.filename)
                if alter is not None:
                    altered = alter(payload)
                    assert altered != payload, f"fixture drifted: {info.filename} unchanged"
                    if altered is None:
                        continue
                    payload = altered
                package.writestr(info, payload, compress_type=info.compress_type)
            for name, alter in alterations.items():
                if name not in payloads:
                    added = alter(None)
                    assert added is not None
                    package.writestr(name, added)

    monkeypatch.setattr(batch_runner, "patch_docx", patch_then_alter)


def _alter_before_packaging(
    monkeypatch: pytest.MonkeyPatch,
    alterations: Mapping[str, Callable[[Optional[bytes]], bytes]],
) -> None:
    """Stand in for a stray write into a part the packaging step reads back."""

    real_build = batch_runner._build_and_patch_output

    def build_after_a_stray_write(docx_path, extract_dir, *args, **kwargs):
        for name, alter in alterations.items():
            path = Path(extract_dir) / name
            before = path.read_bytes() if path.is_file() else None
            after = alter(before)
            assert after != before, f"fixture drifted: {name} unchanged"
            path.parent.mkdir(parents=True, exist_ok=True)
            # Bytes, not text: Path.write_text translates newlines on Windows.
            path.write_bytes(after)
        return real_build(docx_path, extract_dir, *args, **kwargs)

    monkeypatch.setattr(batch_runner, "_build_and_patch_output", build_after_a_stray_write)


def _assert_withheld_for(run, verb: str, name: str) -> dict:
    assert not run.success
    result = run.targets[0]
    assert result.output_path is None
    assert result.stage == "output_publication"
    # Withheld by the whitelist, naming the member, and by nothing earlier.
    assert f"{OUT_OF_REMIT} {verb}: {name}" in str(result.error)
    assert not list(run.run_dir.glob("*.docx"))
    fields = _build_output_fields(run)
    assert fields["failed"] is True
    # The census is recorded before it is enforced, so the failure still
    # shows the check ran and what it counted.
    assert fields["package_members_compared"] >= 1
    assert fields[f"package_members_{verb}"] >= 1
    return fields


def _changed(old: bytes, new: bytes) -> Callable[[Optional[bytes]], bytes]:
    def alter(payload: Optional[bytes]) -> bytes:
        assert payload is not None and payload.count(old) == 1, f"fixture drifted: {old!r}"
        return payload.replace(old, new)

    return alter


_UNTOUCHED_DAMAGE: dict[str, Callable[[Optional[bytes]], bytes]] = {
    "word/footnotes.xml": _changed(b"never touches", b"never  touches"),
    "word/comments.xml": _changed(b"w:author=\"Reviewer\"", b"w:author=\"Someone\""),
    "[trash]/0000.dat": _changed(b"discarded", b"DISCARDED"),
}


@pytest.mark.parametrize("member", sorted(_UNTOUCHED_DAMAGE))
@pytest.mark.parametrize("mode", ALL_MODES)
def test_a_change_to_a_part_no_mode_touches_is_withheld(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch, mode: str, member: str
) -> None:
    source = _canadian_source(tmp_path) if mode == CANADIAN_TO_CSI else None
    _alter_packaged_output(monkeypatch, {member: _UNTOUCHED_DAMAGE[member]})

    if mode == CANADIAN_TO_CSI:
        run, _source = _reverse_run(tmp_path, source)
    else:
        run, _source = _RUNNERS[mode](tmp_path)

    _assert_withheld_for(run, "changed", member)


@pytest.mark.parametrize("mode", ALL_MODES)
def test_a_removed_or_added_member_no_mode_touches_is_withheld(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch, mode: str
) -> None:
    source = _canadian_source(tmp_path) if mode == CANADIAN_TO_CSI else None
    _alter_packaged_output(
        monkeypatch,
        {
            "[trash]/0000.dat": lambda _payload: None,
            "customXml/item1.xml": lambda _payload: b"<stray/>",
        },
    )

    if mode == CANADIAN_TO_CSI:
        run, _source = _reverse_run(tmp_path, source)
    else:
        run, _source = _RUNNERS[mode](tmp_path)

    # The first out-of-remit member in name order is reported; both are
    # counted.
    fields = _assert_withheld_for(run, "removed", "[trash]/0000.dat")
    assert fields["package_members_added"] >= 1


# --- canadian_to_csi: literal markers, and nothing else ----------------------

_SETTINGS = (
    '<?xml version="1.0" encoding="UTF-16" standalone="yes"?>'
    f'<w:settings xmlns:w="{W_NS}"><w:zoom w:percent="120"/>'
    "<w:evenAndOddHeaders/></w:settings>"
)


def _utf16(text: str) -> bytes:
    return codecs.BOM_UTF16_LE + text.encode("utf-16-le")


def _u16(fragment: str) -> bytes:
    return fragment.encode("utf-16-le")


def _canadian_source_with_settings(tmp_path: Path) -> Path:
    canadian = _canadian_source(tmp_path)
    copy = tmp_path / "with-settings" / canadian.name
    _write_members(copy, _members(canadian))
    return _with_parts(
        copy,
        {
            "word/settings.xml": (
                _utf16(_SETTINGS),
                "application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml",
                f"{OFFICE_REL}/settings",
            )
        },
    )


_STRAY_STYLE = (
    b'<w:style w:type="character" w:styleId="Stray"><w:name w:val="Stray"/>'
    b"</w:style></w:styles>"
)
_THEME = (
    '<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" '
    'name="Stray"/>'
).encode("utf-8")
_FONT_TABLE = f'<w:fonts xmlns:w="{W_NS}"/>'.encode("utf-8")

_REVERSE_DAMAGE: dict[str, tuple[str, Callable[[Optional[bytes]], bytes]]] = {
    "word/styles.xml": ("changed", _changed(b"</w:styles>", _STRAY_STYLE)),
    "word/numbering.xml": (
        "changed",
        _changed(b"</w:numbering>", b"<!-- stray --></w:numbering>"),
    ),
    # UTF-16 bytes: compared as bytes, never decoded.
    "word/settings.xml": (
        "changed",
        _changed(_u16('w:percent="120"'), _u16('w:percent="90"')),
    ),
    "word/theme/theme1.xml": ("added", lambda _payload: _THEME),
    "word/fontTable.xml": ("added", lambda _payload: _FONT_TABLE),
}


@pytest.mark.parametrize("member", sorted(_REVERSE_DAMAGE))
def test_canadian_to_csi_withholds_any_change_outside_the_body(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch, member: str
) -> None:
    source = _canadian_source_with_settings(tmp_path)
    verb, damage = _REVERSE_DAMAGE[member]
    _alter_before_packaging(monkeypatch, {member: damage})

    run, _source = _reverse_run(tmp_path, source)

    _assert_withheld_for(run, verb, member)


def test_canadian_to_csi_happy_path_with_settings_changes_only_the_body(
    tmp_path: Path,
) -> None:
    source = _canadian_source_with_settings(tmp_path)
    run, _source = _reverse_run(tmp_path, source)
    assert run.success, "\n".join(run.targets[0].log)
    census = _independent_census(source, run.targets[0].output_path)
    assert census["changed"] == {"word/document.xml"}
    assert not census["added"] and not census["removed"]


# --- csi_to_canadian_standalone: numbering and its wiring, never styles ------


def test_standalone_withholds_a_change_to_the_targets_styles(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    _alter_before_packaging(
        monkeypatch, {"word/styles.xml": _changed(b"</w:styles>", _STRAY_STYLE)}
    )

    run, _source = _standalone_run(tmp_path)

    _assert_withheld_for(run, "changed", "word/styles.xml")


def test_standalone_withholds_a_theme_it_never_had(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    _alter_before_packaging(monkeypatch, {"word/theme/theme1.xml": lambda _p: _THEME})

    run, _source = _standalone_run(tmp_path)

    _assert_withheld_for(run, "added", "word/theme/theme1.xml")


# --- The census and the remit, directly --------------------------------------

_STYLES_REL = f'<Relationship Id="rIdS" Type="{OFFICE_REL}/styles" Target="styles.xml"/>'


def _rels(*relationships: str) -> str:
    return f'<Relationships xmlns="{REL_NS}">{"".join(relationships)}</Relationships>'


def _relationship(rid: str, kind: str, target: str) -> str:
    return f'<Relationship Id="{rid}" Type="{OFFICE_REL}/{kind}" Target="{target}"/>'


def _base_members(**extra: bytes | str) -> dict[str, bytes | str]:
    members: dict[str, bytes | str] = {
        "[Content_Types].xml": f'<Types xmlns="{CT_NS}"/>',
        "word/document.xml": f'<w:document xmlns:w="{W_NS}"><w:body/></w:document>',
        "word/_rels/document.xml.rels": _rels(_STYLES_REL),
        "word/styles.xml": f'<w:styles xmlns:w="{W_NS}"/>',
    }
    members.update(extra)
    return members


def _pair(
    tmp_path: Path,
    source: Mapping[str, bytes | str],
    output: Mapping[str, bytes | str],
) -> tuple[Path, Path]:
    return (
        _write_members(tmp_path / "source.docx", source),
        _write_members(tmp_path / "output.docx", output),
    )


def _remit(
    source: Path,
    output: Path,
    mode: str = FORMAT_ONLY,
    *,
    policy=None,
    manifest: Optional[Mapping] = None,
    verification_out: Optional[dict] = None,
    log: Optional[list] = None,
) -> None:
    _verify_package_member_remit(
        source,
        output,
        policy=policy or application_policy_for_mode(mode),
        header_footer_manifest=manifest,
        verification_out=verification_out,
        log=log,
    )


def test_census_compares_member_bytes_never_decoded_text(tmp_path: Path) -> None:
    text = f'<w:settings xmlns:w="{W_NS}"/>'
    source, output = _pair(
        tmp_path,
        {"same": b"x", "changed": _utf16(text), "removed": b"gone"},
        # The same text re-encoded is a different part: Word reads the bytes.
        {"same": b"x", "changed": text.encode("utf-8"), "added": b"new"},
    )

    census = _package_member_census(source, output)

    assert census.compared == 4
    assert census.changed == ("changed",)
    assert census.added == ("added",)
    assert census.removed == ("removed",)


def test_census_refuses_a_package_with_a_duplicated_member(tmp_path: Path) -> None:
    source = _write_members(tmp_path / "source.docx", {"a": b"1"})
    output = tmp_path / "output.docx"
    with zipfile.ZipFile(output, "w") as package:
        package.writestr("a", b"1")
        with pytest.warns(UserWarning):
            package.writestr("a", b"2")

    with pytest.raises(RuntimeError, match="more than once"):
        _package_member_census(source, output)


def test_the_body_is_always_within_the_remit(tmp_path: Path) -> None:
    changed_body = f'<w:document xmlns:w="{W_NS}"><w:body><w:p/></w:body></w:document>'
    for mode in ALL_MODES:
        source, output = _pair(
            tmp_path / mode, _base_members(), _base_members(**{"word/document.xml": changed_body})
        )
        _remit(source, output, mode)


def test_the_remit_is_derived_from_policy_fields_not_the_mode_name(tmp_path: Path) -> None:
    reverse = application_policy_for_mode(CANADIAN_TO_CSI)
    styled = _base_members(**{"word/styles.xml": f'<w:styles xmlns:w="{W_NS}"><!-- x --></w:styles>'})
    source, output = _pair(tmp_path / "styles", _base_members(), styled)

    with pytest.raises(RuntimeError, match=f"{OUT_OF_REMIT} changed: word/styles.xml"):
        _remit(source, output, policy=reverse)
    # The same mode name, with the one field that governs styles switched on.
    _remit(source, output, policy=replace(reverse, applies_role_styles=True))

    numbered = _base_members(
        **{
            "word/numbering.xml": f'<w:numbering xmlns:w="{W_NS}"/>',
            "[Content_Types].xml": f'<Types xmlns="{CT_NS}"><Default Extension="xml" ContentType="application/xml"/></Types>',
            "word/_rels/document.xml.rels": _rels(
                _STYLES_REL, _relationship("rIdN", "numbering", "numbering.xml")
            ),
        }
    )
    source, output = _pair(tmp_path / "numbering", _base_members(), numbered)
    with pytest.raises(RuntimeError, match=OUT_OF_REMIT):
        _remit(source, output, policy=reverse)
    _remit(source, output, policy=replace(reverse, import_body_numbering=True))

    # And a shell-applying mode stops accepting the shell once the field says so.
    related = _relationship("rIdSet", "settings", "settings.xml")
    shell_source = _base_members(
        **{
            "word/_rels/document.xml.rels": _rels(_STYLES_REL, related),
            "word/settings.xml": f'<w:settings xmlns:w="{W_NS}"/>',
        }
    )
    shell_output = dict(shell_source)
    shell_output["word/settings.xml"] = f'<w:settings xmlns:w="{W_NS}"><w:evenAndOddHeaders/></w:settings>'
    source, output = _pair(tmp_path / "shell", shell_source, shell_output)
    _remit(source, output, FORMAT_ONLY)
    format_only = application_policy_for_mode(FORMAT_ONLY)
    with pytest.raises(RuntimeError, match=f"{OUT_OF_REMIT} changed: word/settings.xml"):
        _remit(source, output, policy=replace(format_only, apply_full_architect_shell=False))


def test_the_settings_part_is_the_related_one_or_the_writers_name(tmp_path: Path) -> None:
    custom = _relationship("rIdSet", "settings", "custom/prefs.xml")
    source_members = _base_members(
        **{
            "word/_rels/document.xml.rels": _rels(_STYLES_REL, custom),
            "word/custom/prefs.xml": f'<w:settings xmlns:w="{W_NS}"/>',
            # A stray part under the conventional name, which nothing relates.
            "word/settings.xml": f'<w:settings xmlns:w="{W_NS}"/>',
            # And a leftover under another name, which nothing relates either.
            "word/custom/old-prefs.xml": f'<w:settings xmlns:w="{W_NS}"/>',
        }
    )
    switched = f'<w:settings xmlns:w="{W_NS}"><w:evenAndOddHeaders/></w:settings>'

    # The part Word reads, whatever it is called: header parity writes here.
    related_output = dict(source_members)
    related_output["word/custom/prefs.xml"] = switched
    source, output = _pair(tmp_path / "related", source_members, related_output)
    _remit(source, output)

    # The name the compat step edits whether or not it is related. That write
    # goes nowhere (a known defect of the compat step, reported in the
    # program's handoffs); the census records it, and it stays publishable.
    stray_output = dict(source_members)
    stray_output["word/settings.xml"] = switched
    source, output = _pair(tmp_path / "stray", source_members, stray_output)
    _remit(source, output)

    # Anything else is outside every remit, even a settings-shaped part.
    leftover_output = dict(source_members)
    leftover_output["word/custom/old-prefs.xml"] = switched
    source, output = _pair(tmp_path / "leftover", source_members, leftover_output)
    with pytest.raises(
        RuntimeError, match=f"{OUT_OF_REMIT} changed: word/custom/old-prefs.xml"
    ):
        _remit(source, output)
    # And none of it for a mode that applies no shell.
    source, output = _pair(tmp_path / "no-shell", source_members, related_output)
    with pytest.raises(RuntimeError, match=f"{OUT_OF_REMIT} changed: word/custom/prefs.xml"):
        _remit(source, output, CANADIAN_TO_CSI)


_SETTINGS_PART = f'<w:settings xmlns:w="{W_NS}"/>'


@pytest.mark.parametrize(
    ("kind", "writers_name", "other_name", "payload"),
    [
        ("settings", "word/settings.xml", "word/custom/prefs.xml", _SETTINGS_PART),
        ("theme", "word/theme/theme1.xml", "word/theme/theme2.xml", _THEME),
        ("fontTable", "word/fontTable.xml", "word/fonts/table.xml", _FONT_TABLE),
    ],
)
def test_a_shell_part_may_be_added_where_the_document_relates_it_or_under_the_writers_name(
    tmp_path: Path, kind: str, writers_name: str, other_name: str, payload
) -> None:
    def added(name: str, *, related: bool) -> dict:
        target = name[len("word/"):]
        extra: dict = {name: payload}
        if related:
            extra["word/_rels/document.xml.rels"] = _rels(
                _STYLES_REL, _relationship("rIdNew", kind, target)
            )
        return _base_members(**extra)

    for name, related in ((writers_name, True), (writers_name, False), (other_name, True)):
        case = tmp_path / f"{Path(name).stem}-{related}"
        source, output = _pair(case, _base_members(), added(name, related=related))
        _remit(source, output)
        # Not for a mode that applies no shell.
        with pytest.raises(RuntimeError, match=f"{OUT_OF_REMIT} added: {name}"):
            _remit(source, output, CSI_TO_CANADIAN_STANDALONE)

    # Under any other name, only where the output's document relates it.
    source, output = _pair(
        tmp_path / "unrelated", _base_members(), added(other_name, related=False)
    )
    with pytest.raises(RuntimeError, match=f"{OUT_OF_REMIT} added: {other_name}"):
        _remit(source, output)


# Header/footer replacement: the target's header9 goes, the architect's
# header1 arrives with its relationships and one media part.
_OLD_HEADER = _relationship("rIdOld", "header", "header9.xml")
_NEW_HEADER = _relationship("rIdNew", "header", "header1.xml")
_HEADER = f'<w:hdr xmlns:w="{W_NS}"/>'
_MEDIA = "word/media/hf_header1_01_0123abcd.png"
_HEADER_RELS = _rels(
    f'<Relationship Id="rIdImg" Type="{OFFICE_REL}/image" '
    'Target="media/hf_header1_01_0123abcd.png"/>'
)


def _header_replacement(tmp_path: Path, **output_overrides) -> tuple[Path, Path]:
    source = _base_members(
        **{
            "word/_rels/document.xml.rels": _rels(_STYLES_REL, _OLD_HEADER),
            "word/header9.xml": _HEADER,
            "word/_rels/header9.xml.rels": _rels(),
        }
    )
    output = _base_members(
        **{
            "word/_rels/document.xml.rels": _rels(_STYLES_REL, _NEW_HEADER),
            "word/header1.xml": _HEADER,
            "word/_rels/header1.xml.rels": _HEADER_RELS,
            _MEDIA: b"\x89PNG",
        }
    )
    output.update(output_overrides)
    return _pair(tmp_path, source, {k: v for k, v in output.items() if v is not None})


def _manifest(**overrides) -> dict:
    manifest = {
        "part_names": {"word/header1.xml"},
        "rels_names": {"word/_rels/header1.xml.rels"},
        "media_names": {_MEDIA},
        "removed_part_names": {"word/header9.xml"},
        "removed_rels_names": {"word/_rels/header9.xml.rels"},
        "replaced_target_parts": True,
    }
    manifest.update(overrides)
    return manifest


def test_an_imported_header_set_is_within_the_remit(tmp_path: Path) -> None:
    source, output = _header_replacement(tmp_path)
    verification: dict = {}
    log: list = []

    _remit(source, output, manifest=_manifest(), verification_out=verification, log=log)

    assert verification == {
        "package_members_compared": 9,
        "package_members_changed": 1,
        "package_members_added": 3,
        "package_members_removed": 2,
    }
    assert "Package members: 9 compared, 1 changed, 3 added, 2 removed" in log
    assert "Package member removed: word/header9.xml" in log
    assert f"Package member added: {_MEDIA}" in log


@pytest.mark.parametrize(
    ("manifest", "verb", "name"),
    [
        # The importer did not remove it, so nothing may have.
        (_manifest(removed_part_names=set()), "removed", "word/header9.xml"),
        # The importer did not write it.
        (_manifest(part_names=set()), "added", "word/header1.xml"),
        (_manifest(rels_names=set()), "added", "word/_rels/header1.xml.rels"),
        (_manifest(media_names=set()), "added", _MEDIA),
        # No manifest at all authorizes no header change, as an omitted
        # prediction authorizes no text change. The first member in name
        # order is the one reported.
        (None, "added", "word/_rels/header1.xml.rels"),
    ],
)
def test_a_header_change_the_manifest_does_not_name_is_withheld(
    tmp_path: Path, manifest, verb: str, name: str
) -> None:
    source, output = _header_replacement(tmp_path)
    with pytest.raises(RuntimeError, match=f"{OUT_OF_REMIT} {verb}: {name}"):
        _remit(source, output, manifest=manifest)


def test_a_manifest_entry_the_packages_do_not_bear_out_authorizes_nothing(
    tmp_path: Path,
) -> None:
    # The manifest names a part the output document does not relate...
    source, output = _header_replacement(
        tmp_path / "unrelated",
        **{
            "word/_rels/document.xml.rels": _rels(_STYLES_REL),
            "word/_rels/header1.xml.rels": None,
            _MEDIA: None,
        },
    )
    with pytest.raises(RuntimeError, match=f"{OUT_OF_REMIT} added: word/header1.xml"):
        _remit(source, output, manifest=_manifest())

    # ...media no imported header relates...
    source, output = _header_replacement(
        tmp_path / "unreferenced", **{"word/_rels/header1.xml.rels": _rels()}
    )
    with pytest.raises(RuntimeError, match=f"{OUT_OF_REMIT} added: {_MEDIA}"):
        _remit(source, output, manifest=_manifest())

    # ...and a removal that was no header of the source's.
    source_members = _base_members(
        **{
            "word/_rels/document.xml.rels": _rels(_STYLES_REL, _OLD_HEADER),
            "word/header9.xml": _HEADER,
            "word/footnotes.xml": FOOTNOTES,
        }
    )
    output_members = {k: v for k, v in source_members.items() if k != "word/footnotes.xml"}
    source, output = _pair(tmp_path / "footnotes", source_members, output_members)
    with pytest.raises(RuntimeError, match=f"{OUT_OF_REMIT} removed: word/footnotes.xml"):
        _remit(
            source,
            output,
            manifest=_manifest(
                part_names=set(),
                rels_names=set(),
                media_names=set(),
                removed_part_names={"word/header9.xml", "word/footnotes.xml"},
            ),
        )


def test_existing_media_is_never_rewritten_even_when_the_manifest_names_it(
    tmp_path: Path,
) -> None:
    media = _base_members(**{_MEDIA: b"\x89PNG"})
    changed = dict(media)
    changed[_MEDIA] = b"\x89PNG changed"
    source, output = _pair(tmp_path, media, changed)
    with pytest.raises(RuntimeError, match=f"{OUT_OF_REMIT} changed: {_MEDIA}"):
        _remit(source, output, manifest=_manifest(media_names={_MEDIA}))


def test_a_mode_with_no_shell_cannot_be_handed_a_header_manifest(tmp_path: Path) -> None:
    source, output = _pair(tmp_path, _base_members(), _base_members())
    for mode in (CSI_TO_CANADIAN_STANDALONE, CANADIAN_TO_CSI):
        with pytest.raises(ValueError, match="header/footer"):
            _remit(source, output, mode, manifest=_manifest())
        # An empty manifest is what those modes carry, and is accepted.
        _remit(source, output, mode, manifest={})


def test_the_census_is_recorded_before_it_is_enforced(tmp_path: Path) -> None:
    source, output = _pair(
        tmp_path,
        _base_members(**{"word/comments.xml": COMMENTS}),
        _base_members(**{"word/comments.xml": COMMENTS.replace("Reviewer", "Someone")}),
    )
    verification: dict = {}
    log: list = []

    with pytest.raises(RuntimeError, match=f"{OUT_OF_REMIT} changed: word/comments.xml"):
        _remit(source, output, verification_out=verification, log=log)

    assert verification["package_members_compared"] == 5
    assert verification["package_members_changed"] == 1
    assert "Package member changed: word/comments.xml" in log
    # Counts only: nothing in the verification record names a member.
    assert all(isinstance(value, int) for value in verification.values())
    json.dumps(verification)
