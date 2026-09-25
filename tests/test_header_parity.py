"""Even-page header parity follows the architect (WI-04).

``w:evenAndOddHeaders`` in ``word/settings.xml`` is the one global switch that
decides whether a section's ``even`` header and footer render. The header and
footer importer wires the architect's ``default``, ``first`` and ``even``
references into every target section, but before WI-04 nothing carried the
switch: an architect with distinct odd and even headers rendered its default
header on every page of the output, and an architect without them left a
target whose switch was on with no header at all on even pages.

The switch now follows the header set it governs. When the architect's set
replaces the target's, the output's switch reads as the architect's does;
when the target keeps its own set, its own switch is left exactly as written.
The final gate proves whichever applies and says so in ``build_output``.
"""

from __future__ import annotations

import codecs
import re
import zipfile
from pathlib import Path
import xml.etree.ElementTree as ET

import pytest

from spec_formatter.pipeline import (
    CANADIAN_TO_CSI,
    CSI_TO_CANADIAN,
    CSI_TO_CANADIAN_STANDALONE,
    format_specifications,
)
from spec_formatter.style_application import batch_runner
from spec_formatter.style_application.arch_env_applier import (
    apply_environment_to_target,
    apply_header_parity,
)
from spec_formatter.style_application.core.header_parity import (
    CT_SETTINGS_CHILD_ORDER,
    HeaderParityError,
    architect_even_and_odd_headers,
    even_and_odd_headers,
    even_and_odd_headers_as_written,
    set_even_and_odd_headers,
)
from spec_formatter.style_application.core.registry import preflight_validate_registries
from spec_formatter.style_application.core.xml_helpers import RootNamespaces
from spec_formatter.style_application.phase2_invariants import validate_docx_package
from tests import test_architect_free_modes as free
from tests import test_unified_roundtrip as roundtrip

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
PLAN = (
    Path(__file__).resolve().parents[1]
    / "docs"
    / "docx_method_hardening"
    / "DOCX_METHOD_HARDENING_PLAN.md"
)

FORMAT_ONLY = "format_only"
ARCHITECT_MODES = (FORMAT_ONLY, CSI_TO_CANADIAN)
ARCHITECT_COMPAT = (
    '<w:compat><w:compatSetting w:name="compatibilityMode" '
    'w:uri="http://schemas.microsoft.com/office/word" w:val="15"/></w:compat>'
)
SWITCH = "<w:evenAndOddHeaders/>"
EVEN_REFERENCE = '<w:headerReference w:type="even" r:id="rIdHdrEven"/>'


def _settings(body: str, *, root: str = f'<w:settings xmlns:w="{W_NS}">') -> str:
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        f"{root}{body}</w:settings>"
    )


def _child_names(settings_xml: str | bytes) -> list[str]:
    root = ET.fromstring(
        settings_xml.encode("utf-8") if isinstance(settings_xml, str) else settings_xml
    )
    return [child.tag.rsplit("}", 1)[-1] for child in root]


# --- The CT_Settings order table ------------------------------------------


def test_order_table_is_the_complete_ct_settings_sequence() -> None:
    assert len(CT_SETTINGS_CHILD_ORDER) == 98
    assert len(set(CT_SETTINGS_CHILD_ORDER)) == 98
    assert CT_SETTINGS_CHILD_ORDER[0] == "w:writeProtection"
    assert CT_SETTINGS_CHILD_ORDER[-1] == "w:listSeparator"
    switch = CT_SETTINGS_CHILD_ORDER.index("w:evenAndOddHeaders")
    assert CT_SETTINGS_CHILD_ORDER[switch - 1] == "w:defaultTableStyle"
    assert CT_SETTINGS_CHILD_ORDER[switch + 1] == "w:bookFoldRevPrinting"
    # The two children from other schemas sit where the schema puts them.
    assert CT_SETTINGS_CHILD_ORDER.index("m:mathPr") == (
        CT_SETTINGS_CHILD_ORDER.index("w:rsids") + 1
    )
    assert CT_SETTINGS_CHILD_ORDER.index("sl:schemaLibrary") == (
        CT_SETTINGS_CHILD_ORDER.index("w:smartTagType") + 1
    )
    assert CT_SETTINGS_CHILD_ORDER.index("w:compat") > switch


def test_order_table_matches_the_plans_appendix_b() -> None:
    """The constant and the plan's appendix name the same sequence.

    Both were checked against the schema text (ISO/IEC 29500-4:2012, the
    transitional ``wml.xsd``); this keeps either from drifting alone.
    """

    appendix = PLAN.read_text(encoding="utf-8").split("## Appendix B", 1)[1]
    listing = appendix.split("```", 2)[1]
    names = [
        token.strip().rstrip("*")
        for token in listing.replace("\n", " ").split(",")
        if token.strip()
    ]
    expected = [name if ":" in name else f"w:{name}" for name in names]
    assert list(CT_SETTINGS_CHILD_ORDER) == expected


# --- Reading the switch ---------------------------------------------------


@pytest.mark.parametrize(
    ("attribute", "expected"),
    [
        ("", True),
        (' w:val="1"', True),
        (' w:val="true"', True),
        (' w:val="on"', True),
        (' w:val="0"', False),
        (' w:val="false"', False),
        (' w:val="off"', False),
        (" w:val=' off '", False),
        (" w:val='true'", True),
    ],
)
def test_switch_reads_every_on_off_spelling(attribute: str, expected: bool) -> None:
    settings = _settings(f"<w:evenAndOddHeaders{attribute}/>")
    assert even_and_odd_headers(settings, "settings") is expected


def test_switch_absent_or_part_absent_reads_off() -> None:
    assert even_and_odd_headers(_settings('<w:zoom w:percent="100"/>'), "s") is False
    assert even_and_odd_headers(None, "s") is False


def test_switch_is_read_by_namespace_not_by_prefix() -> None:
    settings = (
        '<x:settings xmlns:x="' + W_NS + '"><x:evenAndOddHeaders x:val="on"/></x:settings>'
    )
    assert even_and_odd_headers(settings, "s") is True


@pytest.mark.parametrize(
    "body",
    [
        '<w:evenAndOddHeaders w:val="yes"/>',
        '<w:evenAndOddHeaders w:val="False"/>',
        "<w:evenAndOddHeaders/><w:evenAndOddHeaders/>",
        '<w:evenAndOddHeaders/><w:evenAndOddHeaders w:val="0"/>',
    ],
)
def test_switch_that_cannot_be_read_with_certainty_fails_closed(body: str) -> None:
    with pytest.raises(HeaderParityError):
        even_and_odd_headers(_settings(body), "settings")


def test_a_part_that_is_not_settings_is_refused() -> None:
    with pytest.raises(HeaderParityError, match="not a WordprocessingML w:settings"):
        even_and_odd_headers(f'<w:document xmlns:w="{W_NS}"/>', "settings")


def test_switch_as_written_is_uninterpreted() -> None:
    assert even_and_odd_headers_as_written(None, "s") == ()
    assert even_and_odd_headers_as_written(_settings(""), "s") == ()
    assert even_and_odd_headers_as_written(
        _settings('<w:evenAndOddHeaders/><w:evenAndOddHeaders w:val="yes"/>'), "s"
    ) == (None, "yes")


def test_utf16_settings_bytes_are_read() -> None:
    text = (
        '<?xml version="1.0" encoding="UTF-16"?>'
        f'<w:settings xmlns:w="{W_NS}">{SWITCH}</w:settings>'
    )
    payload = codecs.BOM_UTF16_BE + text.encode("utf-16-be")
    assert even_and_odd_headers(payload, "settings") is True


# --- The architect's setting, from the registry ----------------------------


def test_architect_setting_comes_from_the_captured_settings_part() -> None:
    on = {"settings": {"settings_xml": _settings(SWITCH)}}
    off = {"settings": {"settings_xml": _settings('<w:zoom w:percent="95"/>')}}
    none = {"settings": {"settings_xml": None}}
    assert architect_even_and_odd_headers(on) is True
    assert architect_even_and_odd_headers(off) is False
    # A template with no settings part renders with the switch off.
    assert architect_even_and_odd_headers(none) is False


@pytest.mark.parametrize(
    "registry",
    [
        {},
        {"settings": "not-a-dict"},
        {"settings": {"compat": {"compat_xml": None}}},
        {"settings": {"settings_xml": 7}},
    ],
)
def test_a_registry_that_does_not_record_the_setting_is_not_read_as_off(registry) -> None:
    with pytest.raises(HeaderParityError):
        architect_even_and_odd_headers(registry)


# --- Setting and clearing it ----------------------------------------------


@pytest.mark.parametrize(
    ("body", "expected_children"),
    [
        # Before its successor in the sequence.
        (
            "<w:bookFoldRevPrinting/>",
            ["evenAndOddHeaders", "bookFoldRevPrinting"],
        ),
        # After its predecessor in the sequence.
        (
            '<w:defaultTableStyle w:val="TableGrid"/>',
            ["defaultTableStyle", "evenAndOddHeaders"],
        ),
        # Between both neighbours.
        (
            '<w:defaultTableStyle w:val="TableGrid"/><w:bookFoldRevPrinting/>',
            ["defaultTableStyle", "evenAndOddHeaders", "bookFoldRevPrinting"],
        ),
        # Into an empty settings body.
        ("", ["evenAndOddHeaders"]),
        # A settings part shaped as current Word writes one, extension
        # children last: the switch goes after w:defaultTabStop, well before
        # w:compat and everything that follows it.
        (
            '<w:zoom w:percent="100"/><w:proofState w:spelling="clean"/>'
            '<w:defaultTabStop w:val="720"/>'
            '<w:characterSpacingControl w:val="doNotCompress"/>'
            "<w:compat/><w:rsids/><m:mathPr/>"
            '<w:themeFontLang w:val="en-US"/><w:clrSchemeMapping/>'
            '<w:shapeDefaults/><w:decimalSymbol w:val="."/>'
            '<w:listSeparator w:val=","/><w14:docId w14:val="1"/>',
            [
                "zoom",
                "proofState",
                "defaultTabStop",
                "evenAndOddHeaders",
                "characterSpacingControl",
                "compat",
                "rsids",
                "mathPr",
                "themeFontLang",
                "clrSchemeMapping",
                "shapeDefaults",
                "decimalSymbol",
                "listSeparator",
                "docId",
            ],
        ),
    ],
)
def test_switch_is_inserted_at_its_schema_position(body: str, expected_children) -> None:
    root = (
        f'<w:settings xmlns:w="{W_NS}" '
        'xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math" '
        'xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml">'
    )
    before = _settings(body, root=root)
    after = set_even_and_odd_headers(before, True)
    assert _child_names(after) == expected_children
    assert even_and_odd_headers(after, "s") is True
    # Nothing but the switch was added.
    assert after.replace(SWITCH, "", 1) == before
    # Clearing it again restores the part byte for byte.
    assert set_even_and_odd_headers(after, False) == before


def test_switch_is_inserted_into_a_self_closing_settings_root() -> None:
    before = f'<?xml version="1.0"?><w:settings xmlns:w="{W_NS}"/>'
    after = set_even_and_odd_headers(before, True)
    assert after == (
        f'<?xml version="1.0"?><w:settings xmlns:w="{W_NS}">{SWITCH}</w:settings>'
    )


def test_an_unknown_extension_child_does_not_move_the_switch() -> None:
    before = _settings(
        '<w:zoom w:percent="100"/><w15:chartTrackingRefBased xmlns:w15="urn:w15"/>'
        '<w:defaultTabStop w:val="720"/><w:compat/>'
    )
    after = set_even_and_odd_headers(before, True)
    assert _child_names(after) == [
        "zoom",
        "chartTrackingRefBased",
        "defaultTabStop",
        "evenAndOddHeaders",
        "compat",
    ]


def test_switch_is_written_with_the_prefix_the_part_binds() -> None:
    before = f'<ns:settings xmlns:ns="{W_NS}"><ns:zoom ns:percent="90"/></ns:settings>'
    after = set_even_and_odd_headers(before, True)
    assert after == (
        f'<ns:settings xmlns:ns="{W_NS}"><ns:zoom ns:percent="90"/>'
        "<ns:evenAndOddHeaders/></ns:settings>"
    )


def test_switch_already_on_is_left_exactly_as_written() -> None:
    before = _settings('<w:evenAndOddHeaders w:val="true"/>')
    assert set_even_and_odd_headers(before, True) == before
    already_bare = _settings(f'<w:zoom w:percent="90"/>{SWITCH}')
    assert set_even_and_odd_headers(already_bare, True) == already_bare


def test_switch_written_off_is_rewritten_on_at_its_position() -> None:
    # Out of schema order in the source, as the round-trip fixture writes it.
    before = _settings('<w:evenAndOddHeaders w:val="0"/><w:zoom w:percent="90"/>')
    after = set_even_and_odd_headers(before, True)
    assert _child_names(after) == ["zoom", "evenAndOddHeaders"]
    assert even_and_odd_headers(after, "s") is True


def test_clearing_removes_every_copy_and_nothing_else() -> None:
    before = _settings(
        f'{SWITCH}<w:zoom w:percent="90"/><w:evenAndOddHeaders w:val="on"/>'
    )
    after = set_even_and_odd_headers(before, False)
    assert after == _settings('<w:zoom w:percent="90"/>')
    assert set_even_and_odd_headers(after, False) is after


def test_setting_collapses_duplicate_switches_into_one() -> None:
    before = _settings(f'{SWITCH}<w:zoom w:percent="90"/>{SWITCH}')
    after = set_even_and_odd_headers(before, True)
    assert _child_names(after) == ["zoom", "evenAndOddHeaders"]


def test_a_part_with_no_valid_position_for_the_switch_fails_closed() -> None:
    # w:compat must follow the switch and w:zoom must precede it; written
    # the other way round there is nowhere schema-valid to put it.
    before = _settings('<w:compat/><w:zoom w:percent="90"/>')
    with pytest.raises(HeaderParityError, match="not in schema order"):
        set_even_and_odd_headers(before, True)


def test_switch_is_written_unprefixed_into_a_default_namespace_part() -> None:
    before = f'<settings xmlns="{W_NS}"><zoom percent="90"/></settings>'
    after = set_even_and_odd_headers(before, True)
    assert after == f'<settings xmlns="{W_NS}"><zoom percent="90"/><evenAndOddHeaders/></settings>'
    assert even_and_odd_headers(after, "s") is True


def test_a_settings_root_in_another_namespace_is_refused() -> None:
    # Spelled w:settings, but ``w`` is not WordprocessingML here.
    before = f'<w:settings xmlns="{W_NS}" xmlns:w="urn:other"/>'
    with pytest.raises(HeaderParityError, match="not a WordprocessingML w:settings"):
        set_even_and_odd_headers(before, True)


# --- Applying it to an extracted target -------------------------------------


def _extract_dir(tmp_path: Path, settings: bytes | None, *, related: bool = True) -> Path:
    """An extracted target; ``settings``, when given, is related unless told not."""

    extract = tmp_path / "extract"
    (extract / "word" / "_rels").mkdir(parents=True)
    (extract / "[Content_Types].xml").write_bytes(
        b'<?xml version="1.0" encoding="UTF-8"?>'
        b'<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
        b'<Override PartName="/word/document.xml" ContentType="application/'
        b'vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
        b"</Types>"
    )
    relationship = (
        b'<Relationship Id="rIdSettings" Type="http://schemas.openxmlformats.org/'
        b'officeDocument/2006/relationships/settings" Target="settings.xml"/>'
        if settings is not None and related
        else b""
    )
    (extract / "word" / "_rels" / "document.xml.rels").write_bytes(
        b'<?xml version="1.0" encoding="UTF-8"?>'
        b'<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        + relationship
        + b"</Relationships>"
    )
    if settings is not None:
        (extract / "word" / "settings.xml").write_bytes(settings)
    return extract


def test_apply_header_parity_sets_the_switch_in_a_utf16_settings_part(tmp_path: Path) -> None:
    text = (
        '<?xml version="1.0" encoding="UTF-16"?>'
        f'<w:settings xmlns:w="{W_NS}"><w:zoom w:percent="110"/>{ARCHITECT_COMPAT}</w:settings>'
    )
    extract = _extract_dir(tmp_path, codecs.BOM_UTF16_LE + text.encode("utf-16-le"))
    log: list[str] = []
    result = apply_header_parity(
        extract, {"settings": {"settings_xml": _settings(SWITCH)}}, log
    )
    assert result == {
        "even_and_odd_headers": True,
        "changed": True,
        "settings_part": "word/settings.xml",
    }
    written = (extract / "word" / "settings.xml").read_bytes()
    assert written.startswith(b'<?xml version="1.0" encoding="UTF-8"?>')
    assert _child_names(written) == ["zoom", "evenAndOddHeaders", "compat"]


def test_apply_header_parity_creates_a_settings_part_only_when_the_switch_is_on(
    tmp_path: Path,
) -> None:
    off = _extract_dir(tmp_path / "off", None)
    result = apply_header_parity(off, {"settings": {"settings_xml": None}}, [])
    assert result == {"even_and_odd_headers": False, "changed": False, "settings_part": None}
    assert not (off / "word" / "settings.xml").exists()

    on = _extract_dir(tmp_path / "on", None)
    result = apply_header_parity(on, {"settings": {"settings_xml": _settings(SWITCH)}}, [])
    assert result == {
        "even_and_odd_headers": True,
        "changed": True,
        "settings_part": "word/settings.xml",
    }
    assert even_and_odd_headers((on / "word" / "settings.xml").read_bytes(), "s")
    # The created part is wired in, as the compat path wires it.
    assert b"/word/settings.xml" in (on / "[Content_Types].xml").read_bytes()
    assert b'Target="settings.xml"' in (
        on / "word" / "_rels" / "document.xml.rels"
    ).read_bytes()


def test_apply_header_parity_does_not_rewrite_a_part_that_already_matches(
    tmp_path: Path,
) -> None:
    original = _settings(f'<w:zoom w:percent="110"/>{SWITCH}').encode("utf-8")
    extract = _extract_dir(tmp_path, original)
    result = apply_header_parity(
        extract, {"settings": {"settings_xml": _settings(SWITCH)}}, []
    )
    assert result == {
        "even_and_odd_headers": True,
        "changed": False,
        "settings_part": "word/settings.xml",
    }
    assert (extract / "word" / "settings.xml").read_bytes() == original


def _minimal_env_registry(*, switch: bool, header: bool) -> dict:
    sectpr = (
        '<w:sectPr><w:headerReference w:type="default" r:id="rIdArch"/>'
        '<w:pgSz w:w="12240" w:h="15840"/></w:sectPr>'
        if header
        else '<w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>'
    )
    section = {
        "sectPr": sectpr,
        "header_refs": {"default": "rIdArch"} if header else {},
        "footer_refs": {},
    }
    headers = (
        [
            {
                "part_name": "word/header1.xml",
                "rel_id": "rIdArch",
                "xml": f'<w:hdr xmlns:w="{W_NS}"><w:p/></w:hdr>',
            }
        ]
        if header
        else []
    )
    return {
        "settings": {"settings_xml": _settings(SWITCH if switch else "")},
        "page_layout": {"default_section": section, "section_chain": [section]},
        "headers_footers": {"headers": headers, "footers": [], "header_footer_media": []},
    }


def _target_extract_with_document(tmp_path: Path, settings_body: str) -> Path:
    extract = _extract_dir(tmp_path, _settings(settings_body).encode("utf-8"))
    (extract / "word" / "document.xml").write_bytes(
        (
            f'<w:document xmlns:w="{W_NS}" '
            'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
            "<w:body><w:p/><w:sectPr/></w:body></w:document>"
        ).encode("utf-8")
    )
    return extract


def _environment(extract: Path, registry: dict) -> tuple[dict, list[str]]:
    log: list[str] = []
    result = apply_environment_to_target(
        extract,
        registry,
        log,
        apply_theme_flag=False,
        apply_settings_flag=False,
        apply_doc_defaults_flag=False,
        apply_fonts_flag=False,
        architect_namespaces=RootNamespaces({"w": W_NS}),
    )
    return result, log


def test_environment_clears_the_target_switch_when_the_architect_headers_replace_its_own(
    tmp_path: Path,
) -> None:
    extract = _target_extract_with_document(tmp_path, SWITCH)
    result, _log = _environment(extract, _minimal_env_registry(switch=False, header=True))
    assert result["header_footer_import"]["replaced_target_parts"] is True
    assert result["header_parity"] == {
        "follows_architect": True,
        "even_and_odd_headers": False,
        "changed": True,
        "settings_part": "word/settings.xml",
    }
    assert not even_and_odd_headers((extract / "word" / "settings.xml").read_bytes(), "s")


def test_environment_keeps_the_target_switch_when_the_target_keeps_its_headers(
    tmp_path: Path,
) -> None:
    extract = _target_extract_with_document(tmp_path, SWITCH)
    before = (extract / "word" / "settings.xml").read_bytes()
    result, log = _environment(extract, _minimal_env_registry(switch=False, header=False))
    assert result["header_footer_import"]["replaced_target_parts"] is False
    assert result["header_parity"]["follows_architect"] is False
    assert (extract / "word" / "settings.xml").read_bytes() == before
    assert any("own w:evenAndOddHeaders setting is left unchanged" in line for line in log)


# --- Shared-profile preflight -----------------------------------------------


def _preflight_registry(settings_xml, *, even_reference: bool = False) -> dict:
    refs = '<w:headerReference w:type="default" r:id="rIdA"/>'
    header_refs = {"default": "rIdA"}
    headers = [{"part_name": "word/header1.xml", "rel_id": "rIdA", "xml": "<w:hdr/>"}]
    if even_reference:
        refs += '<w:headerReference w:type="even" r:id="rIdB"/>'
        header_refs["even"] = "rIdB"
        headers.append({"part_name": "word/header2.xml", "rel_id": "rIdB", "xml": "<w:hdr/>"})
    section = {
        "sectPr": (
            f'<w:sectPr>{refs}<w:pgSz w:w="12240" w:h="15840"/>'
            '<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" '
            'w:header="720" w:footer="720"/></w:sectPr>'
        ),
        "header_refs": header_refs,
        "footer_refs": {},
    }
    return {
        "styles": {"style_defs": [{"style_id": "CSIPart", "type": "paragraph", "name": "P"}]},
        "settings": {"settings_xml": settings_xml, "compat": {"compat_xml": None}},
        "page_layout": {"default_section": section, "section_chain": [section]},
        "headers_footers": {"headers": headers, "footers": [], "header_footer_media": []},
    }


def _parity_errors(registry: dict, **kwargs) -> list[str]:
    return [
        error
        for error in preflight_validate_registries({"PART": "CSIPart"}, registry, **kwargs)
        if "evenAndOddHeaders" in error or "settings.settings_xml" in error
    ]


def test_preflight_accepts_a_dormant_even_reference_with_the_switch_off() -> None:
    """An ``even`` part with the switch off is valid and common, not an error."""

    registry = _preflight_registry(_settings(""), even_reference=True)
    assert _parity_errors(registry) == []


@pytest.mark.parametrize(
    "settings_xml",
    [
        _settings('<w:evenAndOddHeaders w:val="maybe"/>'),
        _settings(f"{SWITCH}{SWITCH}"),
    ],
)
def test_preflight_rejects_a_template_whose_switch_cannot_be_read(settings_xml) -> None:
    errors = _parity_errors(_preflight_registry(settings_xml))
    assert len(errors) == 1
    # Once, before any target work -- and only when there is a shell to apply.
    assert _parity_errors(_preflight_registry(settings_xml), applies_shell=False) == []


def test_preflight_rejects_a_registry_that_does_not_record_the_setting() -> None:
    registry = _preflight_registry(None)
    del registry["settings"]["settings_xml"]
    assert len(_parity_errors(registry)) == 1


def test_preflight_ignores_the_switch_of_a_template_with_no_headers_or_footers() -> None:
    registry = _preflight_registry(_settings(f"{SWITCH}{SWITCH}"))
    registry["headers_footers"] = {"headers": [], "footers": [], "header_footer_media": []}
    assert _parity_errors(registry) == []


# --- End to end, through format_specifications -----------------------------


def _pair(
    tmp_path: Path,
    mode: str,
    *,
    architect_switch: bool,
    target_switch: bool,
    architect_even_reference: bool = True,
    architect_headers: bool = True,
) -> tuple[Path, Path]:
    if mode == FORMAT_ONLY:
        architect = tmp_path / "architect.docx"
        target = tmp_path / "target.docx"
        roundtrip._write_docx(architect, architect=True)
        roundtrip._write_docx(target, architect=False)
    else:
        architect, target = roundtrip._write_canadian_pair(tmp_path)

    with zipfile.ZipFile(architect) as package:
        architect_document = package.read("word/document.xml").decode("utf-8")
    if not architect_even_reference:
        assert architect_document.count(EVEN_REFERENCE) == 2
        architect_document = architect_document.replace(EVEN_REFERENCE, "")
    if not architect_headers:
        architect_document = re.sub(
            r"<w:(?:header|footer)Reference\b[^>]*/>", "", architect_document
        )
    roundtrip._rewrite_docx_parts(
        architect,
        {
            "word/settings.xml": _settings(
                '<w:zoom w:percent="95"/>'
                + (SWITCH if architect_switch else "")
                + ARCHITECT_COMPAT
            ),
            "word/document.xml": architect_document,
        },
    )
    roundtrip._rewrite_docx_parts(
        target,
        {
            "word/settings.xml": _settings(
                '<w:zoom w:percent="110"/>' + (SWITCH if target_switch else "")
            )
        },
    )
    return architect, target


def _format(tmp_path: Path, architect: Path, target: Path, mode: str):
    return format_specifications(
        architect_template=architect,
        target_specs=[target],
        output_dir=tmp_path / "formatted",
        cache_dir=tmp_path / "template-cache",
        api_key="",
        max_workers=1,
        conversion_mode=mode,
        template_model=f"wi04-{mode}-fixture",
        template_classifier=roundtrip._deterministic_classifier,
    )


def _output_parts(run) -> tuple[bytes, str]:
    result = run.targets[0]
    assert result.success, "\n".join(result.log)
    assert result.output_path is not None
    validate_docx_package(result.output_path)
    with zipfile.ZipFile(result.output_path) as package:
        return (
            package.read("word/settings.xml"),
            package.read("word/document.xml").decode("utf-8"),
        )


def _build_output_fields(run) -> dict:
    return roundtrip._diagnostics_event(run, "build_output")["fields"]


@pytest.mark.parametrize("mode", ARCHITECT_MODES)
def test_architect_without_even_headers_clears_the_target_switch(
    tmp_path: Path, mode: str
) -> None:
    """The inverse: the target's switch would leave even pages with no header."""

    architect, target = _pair(
        tmp_path,
        mode,
        architect_switch=False,
        target_switch=True,
        architect_even_reference=False,
    )
    run = _format(tmp_path, architect, target, mode)
    settings, document = _output_parts(run)

    assert even_and_odd_headers(settings, "output") is False
    assert _child_names(settings) == ["zoom", "compat"]
    assert 'w:type="even"' not in document
    fields = _build_output_fields(run)
    assert fields["header_parity_checked"] is True
    assert fields["header_parity_follows_architect"] is True
    assert fields["even_and_odd_headers"] is False


@pytest.mark.parametrize("mode", ARCHITECT_MODES)
def test_dormant_even_reference_is_imported_and_the_target_switch_cleared(
    tmp_path: Path, mode: str
) -> None:
    """An ``even`` part with the switch off reaches the output exactly so.

    Word renders the template with that part dormant and the default header
    on even pages. The reference is imported unchanged and the target's own
    switch is cleared, which reproduces that rendering; the run is not refused.
    """

    architect, target = _pair(
        tmp_path, mode, architect_switch=False, target_switch=True
    )
    run = _format(tmp_path, architect, target, mode)
    settings, document = _output_parts(run)

    assert even_and_odd_headers(settings, "output") is False
    sections = re.findall(r"<w:sectPr\b[\s\S]*?</w:sectPr>", document)
    assert len(sections) == 2
    assert all('w:type="even"' in section for section in sections)
    with zipfile.ZipFile(run.targets[0].output_path) as package:
        assert b"Even header" in package.read("word/header3.xml")


@pytest.mark.parametrize("mode", ARCHITECT_MODES)
def test_target_that_keeps_its_own_headers_keeps_its_own_switch(
    tmp_path: Path, mode: str
) -> None:
    """With no architect header set to import, the switch belongs to the target."""

    architect, target = _pair(
        tmp_path,
        mode,
        architect_switch=False,
        target_switch=True,
        architect_headers=False,
    )
    run = _format(tmp_path, architect, target, mode)
    settings, document = _output_parts(run)

    assert even_and_odd_headers(settings, "output") is True
    assert "rIdOldHeader" in document
    fields = _build_output_fields(run)
    assert fields["header_parity_checked"] is True
    assert fields["header_parity_follows_architect"] is False
    assert fields["even_and_odd_headers"] is True


def test_switch_reaches_a_target_with_no_settings_part(tmp_path: Path) -> None:
    """Parity does not depend on the compat block to create the settings part."""

    architect, target = _pair(
        tmp_path, FORMAT_ONLY, architect_switch=True, target_switch=False
    )
    roundtrip._rewrite_docx_parts(architect, {"word/settings.xml": _settings(SWITCH)})
    with zipfile.ZipFile(target) as package:
        parts = {name: package.read(name) for name in package.namelist()}
    del parts["word/settings.xml"]
    parts["word/_rels/document.xml.rels"] = re.sub(
        rb'<Relationship Id="rIdSettings"[^>]*/>', b"", parts["word/_rels/document.xml.rels"]
    )
    parts["[Content_Types].xml"] = re.sub(
        rb'<Override PartName="/word/settings.xml"[^>]*/>', b"", parts["[Content_Types].xml"]
    )
    with zipfile.ZipFile(target, "w", zipfile.ZIP_DEFLATED) as package:
        for name, payload in parts.items():
            package.writestr(name, payload)

    run = _format(tmp_path, architect, target, FORMAT_ONLY)
    settings, _document = _output_parts(run)
    assert _child_names(settings) == ["evenAndOddHeaders"]
    assert even_and_odd_headers(settings, "output") is True


def _relate_settings_under_another_name(target: Path, custom_settings: str) -> None:
    """Point the target's settings relationship at ``word/custom/prefs.xml``.

    The conventional ``word/settings.xml`` stays in the package, unrelated, as
    a leftover a tool might write. Word reads only the related part.
    """

    with zipfile.ZipFile(target) as package:
        parts = {name: package.read(name) for name in package.namelist()}
    parts["word/custom/prefs.xml"] = custom_settings.encode("utf-8")
    rels = parts["word/_rels/document.xml.rels"]
    assert rels.count(b'Target="settings.xml"') == 1
    parts["word/_rels/document.xml.rels"] = rels.replace(
        b'Target="settings.xml"', b'Target="custom/prefs.xml"'
    )
    parts["[Content_Types].xml"] = parts["[Content_Types].xml"].replace(
        b"</Types>",
        b'<Override PartName="/word/custom/prefs.xml" ContentType="application/'
        b'vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/></Types>',
    )
    with zipfile.ZipFile(target, "w", zipfile.ZIP_DEFLATED) as package:
        for name, payload in parts.items():
            package.writestr(name, payload)


def test_switch_is_applied_to_the_settings_part_the_document_relates(
    tmp_path: Path,
) -> None:
    """The part Word reads is the one the relationship names, not the usual name."""

    architect, target = _pair(
        tmp_path,
        FORMAT_ONLY,
        architect_switch=False,
        target_switch=False,
        architect_even_reference=False,
    )
    _relate_settings_under_another_name(
        target, _settings(f'<w:zoom w:percent="80"/>{SWITCH}')
    )
    run = _format(tmp_path, architect, target, FORMAT_ONLY)
    result = run.targets[0]
    assert result.success, "\n".join(result.log)
    validate_docx_package(result.output_path)
    with zipfile.ZipFile(result.output_path) as package:
        related = package.read("word/custom/prefs.xml")
        rels = package.read("word/_rels/document.xml.rels")
    assert b'Target="custom/prefs.xml"' in rels
    assert even_and_odd_headers(related, "output") is False
    assert _child_names(related) == ["zoom"]
    assert _build_output_fields(run)["even_and_odd_headers"] is False


def test_a_stray_unrelated_settings_part_is_not_activated_to_carry_the_switch(
    tmp_path: Path,
) -> None:
    """Relating a leftover part would make every other setting in it live."""

    extract = _extract_dir(tmp_path, _settings("").encode("utf-8"), related=False)
    before = (extract / "word" / "settings.xml").read_bytes()
    with pytest.raises(HeaderParityError, match="does not relate"):
        apply_header_parity(extract, {"settings": {"settings_xml": _settings(SWITCH)}}, [])
    assert (extract / "word" / "settings.xml").read_bytes() == before
    # With the switch off there is nothing to write, so nothing to refuse:
    # a document that relates no settings part already reads as off.
    result = apply_header_parity(extract, {"settings": {"settings_xml": None}}, [])
    assert result == {"even_and_odd_headers": False, "changed": False, "settings_part": None}


# --- The final gate ---------------------------------------------------------


def _corrupt_settings_before_publication(monkeypatch, damage) -> None:
    real_build = batch_runner._build_and_patch_output

    def build_after_a_late_corruption(docx_path, extract_dir, *args, **kwargs):
        # Stands in for any later step that loses or invents the switch. Bytes,
        # not text: Path.write_text translates newlines on Windows.
        settings = Path(extract_dir) / "word" / "settings.xml"
        payload = settings.read_bytes()
        damaged = damage(payload)
        assert damaged != payload, "fixture drifted: the damage did not apply"
        settings.write_bytes(damaged)
        return real_build(docx_path, extract_dir, *args, **kwargs)

    monkeypatch.setattr(batch_runner, "_build_and_patch_output", build_after_a_late_corruption)


def _assert_withheld(run) -> None:
    assert not run.success
    result = run.targets[0]
    assert result.output_path is None
    assert result.stage == "output_publication"
    # Withheld by the parity invariant, not by some other check.
    assert "even/odd header setting" in str(result.error)
    assert not list(run.run_dir.glob("*.docx"))
    fields = _build_output_fields(run)
    assert fields["failed"] is True
    assert fields["header_parity_checked"] is True


@pytest.mark.parametrize("mode", ARCHITECT_MODES)
def test_gate_withholds_an_output_that_lost_the_architects_switch(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch, mode: str
) -> None:
    architect, target = _pair(tmp_path, mode, architect_switch=True, target_switch=False)
    _corrupt_settings_before_publication(
        monkeypatch, lambda payload: payload.replace(SWITCH.encode(), b"", 1)
    )
    run = _format(tmp_path, architect, target, mode)
    _assert_withheld(run)
    assert _build_output_fields(run)["header_parity_follows_architect"] is True


def test_gate_withholds_an_output_that_changed_the_targets_own_switch(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    architect, target = _pair(
        tmp_path,
        FORMAT_ONLY,
        architect_switch=False,
        target_switch=False,
        architect_headers=False,
    )
    _corrupt_settings_before_publication(
        monkeypatch,
        lambda payload: payload.replace(b"</w:settings>", SWITCH.encode() + b"</w:settings>", 1),
    )
    run = _format(tmp_path, architect, target, FORMAT_ONLY)
    _assert_withheld(run)
    assert _build_output_fields(run)["header_parity_follows_architect"] is False


def _package(path: Path, parts: dict[str, str]) -> Path:
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as package:
        for name, payload in parts.items():
            package.writestr(name, payload)
    return path


def _document_rels(*targets: str) -> str:
    relationships = "".join(
        f'<Relationship Id="rId{index}" Type="http://schemas.openxmlformats.org/'
        f'officeDocument/2006/relationships/settings" Target="{target}"/>'
        for index, target in enumerate(targets)
    )
    return (
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/'
        f'relationships">{relationships}</Relationships>'
    )


def test_gate_reads_the_settings_part_the_document_relates(tmp_path: Path) -> None:
    """Word follows the relationship, not the conventional part name."""

    from spec_formatter.style_application.phase2_invariants import (
        _document_settings_part,
    )

    related = _settings(SWITCH)
    docx = _package(
        tmp_path / "custom.docx",
        {
            "word/_rels/document.xml.rels": _document_rels("custom/Prefs.xml"),
            "word/custom/prefs.xml": related,
            "word/settings.xml": _settings(""),
        },
    )
    assert _document_settings_part(docx, "output") == related.encode("utf-8")

    bare = _package(tmp_path / "bare.docx", {"word/document.xml": "<w:document/>"})
    assert _document_settings_part(bare, "output") is None
    unrelated = _package(
        tmp_path / "unrelated.docx",
        {"word/_rels/document.xml.rels": _document_rels(), "word/settings.xml": related},
    )
    assert _document_settings_part(unrelated, "output") is None

    two = _package(
        tmp_path / "two.docx",
        {
            "word/_rels/document.xml.rels": _document_rels("settings.xml", "other.xml"),
            "word/settings.xml": related,
            "word/other.xml": related,
        },
    )
    with pytest.raises(RuntimeError, match="more than one settings part"):
        _document_settings_part(two, "output")
    missing = _package(
        tmp_path / "missing.docx",
        {"word/_rels/document.xml.rels": _document_rels("settings.xml")},
    )
    with pytest.raises(RuntimeError, match="names no part"):
        _document_settings_part(missing, "output")


# --- Architect-free modes ---------------------------------------------------


def _with_settings(source: Path, settings: bytes) -> Path:
    """``source`` with a settings part added and wired in."""

    with zipfile.ZipFile(source) as package:
        parts = {name: package.read(name) for name in package.namelist()}
    parts["word/settings.xml"] = settings
    parts["word/_rels/document.xml.rels"] = parts["word/_rels/document.xml.rels"].replace(
        b"</Relationships>",
        b'<Relationship Id="rIdSettings" Type="http://schemas.openxmlformats.org/'
        b'officeDocument/2006/relationships/settings" Target="settings.xml"/>'
        b"</Relationships>",
    )
    parts["[Content_Types].xml"] = parts["[Content_Types].xml"].replace(
        b"</Types>",
        b'<Override PartName="/word/settings.xml" ContentType="application/'
        b'vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/></Types>',
    )
    with zipfile.ZipFile(source, "w", zipfile.ZIP_DEFLATED) as package:
        for name, payload in parts.items():
            package.writestr(name, payload)
    return source


def _utf16_settings_with_switch() -> bytes:
    text = (
        '<?xml version="1.0" encoding="UTF-16" standalone="yes"?>'
        f'<w:settings xmlns:w="{W_NS}"><w:zoom w:percent="120"/>{SWITCH}'
        '<w:trackRevisions w:val="false"/></w:settings>'
    )
    return codecs.BOM_UTF16_LE + text.encode("utf-16-le")


def _settings_bytes(docx: Path) -> bytes:
    with zipfile.ZipFile(docx) as package:
        return package.read("word/settings.xml")


def test_architect_free_modes_leave_the_target_settings_byte_identical(
    tmp_path: Path,
) -> None:
    settings = _utf16_settings_with_switch()
    source = _with_settings(free._write_docx(tmp_path / "source" / "spec.docx"), settings)

    canadian = free._run(tmp_path, source, CSI_TO_CANADIAN_STANDALONE, "canadian")
    assert canadian.success, "\n".join(canadian.targets[0].log)
    canadian_output = canadian.targets[0].output_path
    assert _settings_bytes(canadian_output) == settings

    back = free._run(tmp_path, canadian_output, CANADIAN_TO_CSI, "csi")
    assert back.success, "\n".join(back.targets[0].log)
    assert _settings_bytes(back.targets[0].output_path) == settings

    for run in (canadian, back):
        fields = _build_output_fields(run)
        assert fields["header_parity_checked"] is True
        assert fields["header_parity_follows_architect"] is False
        assert fields["even_and_odd_headers"] is True
