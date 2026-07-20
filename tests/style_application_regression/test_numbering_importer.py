"""Tests for numbering_importer.py — fail-fast validation and no font injection."""

import pytest
import re
import xml.etree.ElementTree as ET
from pathlib import Path

from spec_formatter.style_application.numbering_importer import (
    build_numbering_import_plan,
    inject_numbering_into_xml,
    import_numbering,
    extract_used_num_ids_from_styles,
)


# ---------------------------------------------------------------------------
# Helpers — minimal XML fixtures
# ---------------------------------------------------------------------------

def _make_arch_styles_xml(styles):
    """Build a minimal styles.xml from a list of (styleId, numId_or_None)."""
    blocks = []
    for sid, num_id in styles:
        numpr = ""
        if num_id is not None:
            numpr = f'<w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="{num_id}"/></w:numPr></w:pPr>'
        blocks.append(
            f'<w:style w:type="paragraph" w:styleId="{sid}">'
            f'{numpr}'
            f'</w:style>'
        )
    return f'<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">{"".join(blocks)}</w:styles>'


def _make_registry(abstract_nums=None, nums=None):
    """Build a minimal arch_template_registry with numbering data."""
    reg = {}
    if abstract_nums is not None or nums is not None:
        reg["numbering"] = {
            "abstract_nums": abstract_nums or [],
            "nums": nums or [],
        }
    return reg


def _make_abstract_num(abstract_num_id, rpr_xml=""):
    """Build a minimal abstractNum dict entry."""
    xml = (
        f'<w:abstractNum w:abstractNumId="{abstract_num_id}">'
        f'<w:nsid w:val="AABB0011"/>'
        f'<w:lvl w:ilvl="0">{rpr_xml}<w:start w:val="1"/>'
        f'<w:numFmt w:val="decimal"/></w:lvl>'
        f'</w:abstractNum>'
    )
    return {"abstractNumId": abstract_num_id, "xml": xml}


def _make_num(num_id, abstract_num_id):
    """Build a minimal num dict entry."""
    xml = (
        f'<w:num w:numId="{num_id}">'
        f'<w:abstractNumId w:val="{abstract_num_id}"/>'
        f'</w:num>'
    )
    return {"numId": num_id, "abstractNumId": abstract_num_id, "xml": xml}


MINIMAL_TARGET_NUMBERING = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    '<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
    '<w:abstractNum w:abstractNumId="0"><w:nsid w:val="00000001"/>'
    '<w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/></w:lvl>'
    '</w:abstractNum>'
    '<w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num>'
    '</w:numbering>'
)


CT_NS = "http://schemas.openxmlformats.org/package/2006/content-types"
PKG_REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"


def _write_minimal_package_wiring(root: Path):
    (root / "word" / "_rels").mkdir(parents=True, exist_ok=True)
    (root / "[Content_Types].xml").write_text(
        f'<Types xmlns="{CT_NS}">'
        '<Default Extension="xml" ContentType="application/xml"/>'
        '<Default Extension="rels" '
        'ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
        '</Types>',
        encoding="utf-8",
    )
    (root / "word" / "_rels" / "document.xml.rels").write_text(
        f'<Relationships xmlns="{PKG_REL_NS}">'
        '<Relationship Id="rId4" '
        'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" '
        'Target="styles.xml"/>'
        '</Relationships>',
        encoding="utf-8",
    )


# ---------------------------------------------------------------------------
# Test: no font injection in inject_numbering_into_xml
# ---------------------------------------------------------------------------

class TestNoFontInjection:
    def test_inject_numbering_preserves_architect_rpr(self):
        """Verify that inject_numbering_into_xml does NOT modify <w:rPr> blocks."""
        rpr = '<w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/><w:sz w:val="24"/></w:rPr>'
        abstract_xml = (
            f'<w:abstractNum w:abstractNumId="10">'
            f'<w:nsid w:val="AABB0011"/>'
            f'<w:lvl w:ilvl="0">{rpr}<w:start w:val="1"/>'
            f'<w:numFmt w:val="decimal"/></w:lvl>'
            f'</w:abstractNum>'
        )
        num_xml = (
            '<w:num w:numId="10">'
            '<w:abstractNumId w:val="10"/>'
            '</w:num>'
        )

        result = inject_numbering_into_xml(
            MINIMAL_TARGET_NUMBERING,
            [{"xml": abstract_xml}],
            [{"xml": num_xml}],
        )

        # The original rPr must appear verbatim — no Arial injection
        assert rpr in result
        assert "Arial" not in result

    def test_inject_numbering_preserves_rpr_without_fonts(self):
        """rPr blocks without rFonts should NOT get Arial injected."""
        rpr = '<w:rPr><w:b/></w:rPr>'
        abstract_xml = (
            f'<w:abstractNum w:abstractNumId="10">'
            f'<w:nsid w:val="AABB0011"/>'
            f'<w:lvl w:ilvl="0">{rpr}<w:start w:val="1"/>'
            f'<w:numFmt w:val="decimal"/></w:lvl>'
            f'</w:abstractNum>'
        )

        result = inject_numbering_into_xml(
            MINIMAL_TARGET_NUMBERING,
            [{"xml": abstract_xml}],
            [],
        )

        assert rpr in result
        assert "Arial" not in result

    def test_inject_numbering_no_rpr_no_font_added(self):
        """abstractNum with NO rPr at all — no rPr or font should be added."""
        abstract_xml = (
            '<w:abstractNum w:abstractNumId="10">'
            '<w:nsid w:val="AABB0011"/>'
            '<w:lvl w:ilvl="0"><w:start w:val="1"/>'
            '<w:numFmt w:val="decimal"/></w:lvl>'
            '</w:abstractNum>'
        )

        result = inject_numbering_into_xml(
            MINIMAL_TARGET_NUMBERING,
            [{"xml": abstract_xml}],
            [],
        )

        # Count rPr occurrences — should only be those in original target (if any)
        assert "Arial" not in result
        # The injected abstractNum should not have had rPr added
        assert abstract_xml in result


# ---------------------------------------------------------------------------
# Test: fail-fast on missing numId in registry
# ---------------------------------------------------------------------------

class TestBuildPlanFailFast:
    def test_raises_on_missing_num_id(self):
        """build_numbering_import_plan raises when a referenced numId is missing."""
        styles_xml = _make_arch_styles_xml([("CSILevel1", 99)])
        registry = _make_registry(
            abstract_nums=[_make_abstract_num(5)],
            nums=[_make_num(2, 5)],  # has numId=2 but style references 99
        )

        with pytest.raises(ValueError, match="missing required numId"):
            build_numbering_import_plan(
                registry,
                styles_xml,
                MINIMAL_TARGET_NUMBERING,
                ["CSILevel1"],
            )

    def test_raises_on_missing_abstract_num(self):
        """build_numbering_import_plan raises when a referenced abstractNumId is missing."""
        styles_xml = _make_arch_styles_xml([("CSILevel1", 2)])
        registry = _make_registry(
            abstract_nums=[],  # no abstractNums at all
            nums=[_make_num(2, 5)],  # numId=2 references abstractNum 5
        )

        with pytest.raises(ValueError, match="missing required abstractNum"):
            build_numbering_import_plan(
                registry,
                styles_xml,
                MINIMAL_TARGET_NUMBERING,
                ["CSILevel1"],
            )

    def test_raises_on_empty_nums_list(self):
        """Style references numId but registry nums list is empty."""
        styles_xml = _make_arch_styles_xml([("CSILevel1", 2)])
        registry = _make_registry(
            abstract_nums=[_make_abstract_num(5)],
            nums=[],  # empty — numId=2 can't be found
        )

        with pytest.raises(ValueError, match="missing required numId"):
            build_numbering_import_plan(
                registry,
                styles_xml,
                MINIMAL_TARGET_NUMBERING,
                ["CSILevel1"],
            )

    def test_succeeds_when_no_numbering_needed(self):
        """Empty plan returned when styles don't reference numbering."""
        styles_xml = _make_arch_styles_xml([("PlainStyle", None)])
        registry = _make_registry(
            abstract_nums=[_make_abstract_num(5)],
            nums=[{
                "numId": 2,
                "abstractNumId": 5,
                "xml": '<w:num w:numId="2" w16cid:durableId="1" xmlns:w16cid="http://schemas.microsoft.com/office/word/2016/wordml/cid"><w:abstractNumId w:val="5"/></w:num>',
            }],
        )

        plan = build_numbering_import_plan(
            registry,
            styles_xml,
            MINIMAL_TARGET_NUMBERING,
            ["PlainStyle"],
        )
        assert plan["abstract_nums_to_import"] == []
        assert plan["nums_to_import"] == []
        assert plan["style_numid_remap"] == {}

    def test_collision_safe_ids_when_hash_matches_existing(self, monkeypatch):
        styles_xml = _make_arch_styles_xml([("CSILevel1", 2)])
        registry = _make_registry(
            abstract_nums=[_make_abstract_num(5)],
            nums=[{
                "numId": 2,
                "abstractNumId": 5,
                "xml": '<w:num w:numId="2" w16cid:durableId="1" xmlns:w16cid="http://schemas.microsoft.com/office/word/2016/wordml/cid"><w:abstractNumId w:val="5"/></w:num>',
            }],
        )
        target_numbering = (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
            '<w:abstractNum w:abstractNumId="0"><w:nsid w:val="COLLIDE1"/></w:abstractNum>'
            '<w:num w:numId="1" w16cid:durableId="999" xmlns:w16cid="http://schemas.microsoft.com/office/word/2016/wordml/cid"/>'
            '</w:numbering>'
        )

        seq_nsid = iter(["COLLIDE1", "UNIQUE99"])
        seq_durable = iter(["999", "12345"])
        monkeypatch.setattr("spec_formatter.style_application.numbering_importer._generate_unique_nsid", lambda _xml: next(seq_nsid))
        monkeypatch.setattr("spec_formatter.style_application.numbering_importer._generate_unique_durable_id", lambda _xml: next(seq_durable))

        plan = build_numbering_import_plan(
            registry,
            styles_xml,
            target_numbering,
            ["CSILevel1"],
        )
        assert 'w:nsid w:val="UNIQUE99"' in plan["abstract_nums_to_import"][0]["xml"]
        assert 'w16cid:durableId="12345"' in plan["nums_to_import"][0]["xml"]


# ---------------------------------------------------------------------------
# Test: import_numbering creates missing numbering.xml and package wiring
# ---------------------------------------------------------------------------

class TestImportNumberingPartHandling:
    def test_creates_numbering_part_and_wiring_when_needed(self, tmp_path):
        """A valid unnumbered target gains numbering.xml and both OPC declarations."""
        _write_minimal_package_wiring(tmp_path)
        word_dir = tmp_path / "word"

        styles_xml = _make_arch_styles_xml([("CSILevel1", 2)])
        registry = _make_registry(
            abstract_nums=[_make_abstract_num(5)],
            nums=[_make_num(2, 5)],
        )
        log = []

        remap = import_numbering(
            target_extract_dir=tmp_path,
            arch_template_registry=registry,
            arch_styles_xml=styles_xml,
            style_ids_to_import=["CSILevel1"],
            log=log,
        )

        assert remap["CSILevel1"] == {"old_numId": 2, "new_numId": 1}
        numbering_root = ET.fromstring((word_dir / "numbering.xml").read_bytes())
        assert len(numbering_root.findall("{*}abstractNum")) == 1
        assert len(numbering_root.findall("{*}num")) == 1

        content_types_root = ET.fromstring((tmp_path / "[Content_Types].xml").read_bytes())
        numbering_overrides = [
            node
            for node in content_types_root.findall(f"{{{CT_NS}}}Override")
            if node.attrib.get("PartName") == "/word/numbering.xml"
        ]
        assert len(numbering_overrides) == 1
        assert numbering_overrides[0].attrib["ContentType"].endswith(
            "wordprocessingml.numbering+xml"
        )

        rels_root = ET.fromstring(
            (word_dir / "_rels" / "document.xml.rels").read_bytes()
        )
        numbering_rels = [
            node
            for node in rels_root.findall(f"{{{PKG_REL_NS}}}Relationship")
            if node.attrib.get("Type", "").endswith("/numbering")
        ]
        assert len(numbering_rels) == 1
        assert numbering_rels[0].attrib["Id"] == "rId5"
        assert numbering_rels[0].attrib["Target"] == "numbering.xml"
        assert any("Created word/numbering.xml" in message for message in log)

        # Re-importing may add definitions, but must never duplicate the OPC
        # content-type declaration or document relationship.
        import_numbering(
            target_extract_dir=tmp_path,
            arch_template_registry=registry,
            arch_styles_xml=styles_xml,
            style_ids_to_import=["CSILevel1"],
            log=[],
        )
        content_types_root = ET.fromstring((tmp_path / "[Content_Types].xml").read_bytes())
        rels_root = ET.fromstring(
            (word_dir / "_rels" / "document.xml.rels").read_bytes()
        )
        assert sum(
            node.attrib.get("PartName") == "/word/numbering.xml"
            for node in content_types_root.findall(f"{{{CT_NS}}}Override")
        ) == 1
        assert sum(
            node.attrib.get("Type", "").endswith("/numbering")
            for node in rels_root.findall(f"{{{PKG_REL_NS}}}Relationship")
        ) == 1

    def test_reuses_predeclared_numbering_wiring_when_part_was_missing(self, tmp_path):
        _write_minimal_package_wiring(tmp_path)
        content_types_path = tmp_path / "[Content_Types].xml"
        content_types_path.write_text(
            content_types_path.read_text(encoding="utf-8").replace(
                "</Types>",
                '<Override PartName="/word/numbering.xml" '
                'ContentType="application/xml"/></Types>',
            ),
            encoding="utf-8",
        )
        rels_path = tmp_path / "word" / "_rels" / "document.xml.rels"
        rels_path.write_text(
            rels_path.read_text(encoding="utf-8").replace(
                "</Relationships>",
                '<Relationship Id="rId9" '
                'Type="http://schemas.openxmlformats.org/officeDocument/2006/'
                'relationships/numbering" Target="numbering.xml"/></Relationships>',
            ),
            encoding="utf-8",
        )

        import_numbering(
            target_extract_dir=tmp_path,
            arch_template_registry=_make_registry(
                abstract_nums=[_make_abstract_num(5)],
                nums=[_make_num(2, 5)],
            ),
            arch_styles_xml=_make_arch_styles_xml([("CSILevel1", 2)]),
            style_ids_to_import=["CSILevel1"],
            log=[],
        )

        content_types_root = ET.fromstring(content_types_path.read_bytes())
        rels_root = ET.fromstring(rels_path.read_bytes())
        overrides = [
            node
            for node in content_types_root.findall(f"{{{CT_NS}}}Override")
            if node.attrib.get("PartName") == "/word/numbering.xml"
        ]
        relationships = [
            node
            for node in rels_root.findall(f"{{{PKG_REL_NS}}}Relationship")
            if node.attrib.get("Type", "").endswith("/numbering")
        ]
        assert len(overrides) == 1
        assert overrides[0].attrib["ContentType"].endswith(
            "wordprocessingml.numbering+xml"
        )
        assert len(relationships) == 1
        assert relationships[0].attrib["Id"] == "rId9"

    def test_succeeds_when_no_numbering_xml_not_needed(self, tmp_path):
        """import_numbering returns {} when no styles need numbering, even without numbering.xml."""
        word_dir = tmp_path / "word"
        word_dir.mkdir()

        styles_xml = _make_arch_styles_xml([("PlainStyle", None)])
        registry = _make_registry(
            abstract_nums=[_make_abstract_num(5)],
            nums=[_make_num(2, 5)],
        )
        log = []

        result = import_numbering(
            target_extract_dir=tmp_path,
            arch_template_registry=registry,
            arch_styles_xml=styles_xml,
            style_ids_to_import=["PlainStyle"],
            log=log,
        )
        assert result == {}

    def test_raises_when_registry_has_no_numbering_but_needed(self, tmp_path):
        """import_numbering raises when registry has no numbering data but styles need it."""
        word_dir = tmp_path / "word"
        word_dir.mkdir()
        (word_dir / "numbering.xml").write_text(MINIMAL_TARGET_NUMBERING, encoding="utf-8")

        styles_xml = _make_arch_styles_xml([("CSILevel1", 2)])
        registry = {}  # no numbering key at all
        log = []

        with pytest.raises(ValueError, match="no numbering data"):
            import_numbering(
                target_extract_dir=tmp_path,
                arch_template_registry=registry,
                arch_styles_xml=styles_xml,
                style_ids_to_import=["CSILevel1"],
                log=log,
            )

    def test_raises_when_numbering_empty_but_needed(self, tmp_path):
        """Registry has numbering key but empty lists — fail-fast when styles need it."""
        word_dir = tmp_path / "word"
        word_dir.mkdir()
        (word_dir / "numbering.xml").write_text(MINIMAL_TARGET_NUMBERING, encoding="utf-8")

        styles_xml = _make_arch_styles_xml([("CSILevel1", 2)])
        registry = _make_registry(abstract_nums=[], nums=[])
        log = []

        with pytest.raises(ValueError, match="empty numbering definitions"):
            import_numbering(
                target_extract_dir=tmp_path,
                arch_template_registry=registry,
                arch_styles_xml=styles_xml,
                style_ids_to_import=["CSILevel1"],
                log=log,
            )


# ---------------------------------------------------------------------------
# Test: happy path — full remap with no font injection
# ---------------------------------------------------------------------------

class TestHappyPath:
    def test_import_happy_path_no_font_injection(self, tmp_path):
        """Full remap works correctly and no font XML is injected."""
        word_dir = tmp_path / "word"
        word_dir.mkdir()
        (word_dir / "numbering.xml").write_text(MINIMAL_TARGET_NUMBERING, encoding="utf-8")

        rpr = '<w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri"/><w:sz w:val="18"/></w:rPr>'
        styles_xml = _make_arch_styles_xml([("CSILevel1", 2)])
        registry = _make_registry(
            abstract_nums=[_make_abstract_num(5, rpr_xml=rpr)],
            nums=[_make_num(2, 5)],
        )
        log = []

        remap = import_numbering(
            target_extract_dir=tmp_path,
            arch_template_registry=registry,
            arch_styles_xml=styles_xml,
            style_ids_to_import=["CSILevel1"],
            log=log,
        )

        # Verify remap was produced
        assert "CSILevel1" in remap
        assert remap["CSILevel1"]["old_numId"] == 2
        assert remap["CSILevel1"]["new_numId"] > 1  # remapped to avoid collision

        # Verify the written numbering.xml has no Arial injection
        result_xml = (word_dir / "numbering.xml").read_text(encoding="utf-8")
        assert "Arial" not in result_xml
        # Original font should be preserved
        assert "Calibri" in result_xml

    def test_nsid_regenerated(self, tmp_path):
        """nsid values must be regenerated to avoid collisions with source doc."""
        word_dir = tmp_path / "word"
        word_dir.mkdir()
        (word_dir / "numbering.xml").write_text(MINIMAL_TARGET_NUMBERING, encoding="utf-8")

        styles_xml = _make_arch_styles_xml([("CSILevel1", 2)])
        registry = _make_registry(
            abstract_nums=[_make_abstract_num(5)],
            nums=[_make_num(2, 5)],
        )

        import_numbering(
            target_extract_dir=tmp_path,
            arch_template_registry=registry,
            arch_styles_xml=styles_xml,
            style_ids_to_import=["CSILevel1"],
            log=[],
        )

        result_xml = (word_dir / "numbering.xml").read_text(encoding="utf-8")
        # The original nsid "AABB0011" should have been replaced with a new random value
        nsid_matches = re.findall(r'<w:nsid\s+w:val="([^"]+)"', result_xml)
        # There should be at least 2 nsid values (original target + imported)
        assert len(nsid_matches) >= 2
        # The imported one should NOT be "AABB0011" (regenerated)
        imported_nsids = [n for n in nsid_matches if n != "00000001"]  # original target nsid
        assert all(n != "AABB0011" for n in imported_nsids)

    def test_new_numid_avoids_collision(self, tmp_path):
        """Remapped numId must be greater than existing target max."""
        word_dir = tmp_path / "word"
        word_dir.mkdir()
        (word_dir / "numbering.xml").write_text(MINIMAL_TARGET_NUMBERING, encoding="utf-8")

        styles_xml = _make_arch_styles_xml([("CSILevel1", 2)])
        registry = _make_registry(
            abstract_nums=[_make_abstract_num(5)],
            nums=[_make_num(2, 5)],
        )

        remap = import_numbering(
            target_extract_dir=tmp_path,
            arch_template_registry=registry,
            arch_styles_xml=styles_xml,
            style_ids_to_import=["CSILevel1"],
            log=[],
        )

        # Target has numId=1, so the new numId must be > 1
        new_num_id = remap["CSILevel1"]["new_numId"]
        assert new_num_id > 1

def test_inject_numbering_into_xml_when_target_has_no_num():
    target = (
        '<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
        '<w:abstractNum w:abstractNumId="0"/>'
        '</w:numbering>'
    )
    abstract = '<w:abstractNum w:abstractNumId="9"/>'
    num = '<w:num w:numId="3"><w:abstractNumId w:val="9"/></w:num>'
    out = inject_numbering_into_xml(target, [{"xml": abstract}], [{"xml": num, "new_abstract_id": 9}])
    assert out.index(abstract) < out.index(num)


def test_imported_nums_reference_existing_abstract_nums_after_injection():
    target = '<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"></w:numbering>'
    with pytest.raises(ValueError, match="missing abstractNumId"):
        inject_numbering_into_xml(target, [], [{"xml": '<w:num w:numId="4"><w:abstractNumId w:val="88"/></w:num>', "new_abstract_id": 88}])


def test_inject_numbering_merges_source_namespace_declarations():
    target = (
        '<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
        '</w:numbering>'
    )
    source = (
        '<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
        'xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml" '
        'xmlns:w16cid="http://schemas.microsoft.com/office/word/2016/wordml/cid">'
        '</w:numbering>'
    )
    abstract = (
        '<w:abstractNum w:abstractNumId="9" w15:restartNumberingAfterBreak="0">'
        '<w:lvl w:ilvl="0"/></w:abstractNum>'
    )
    num = (
        '<w:num w:numId="10" w16cid:durableId="123">'
        '<w:abstractNumId w:val="9"/></w:num>'
    )
    out = inject_numbering_into_xml(
        target,
        [{"xml": abstract}],
        [{"xml": num, "new_abstract_id": 9}],
        source_numbering_xml=source,
    )
    ET.fromstring(out)
    assert "xmlns:w15=" in out
    assert "xmlns:w16cid=" in out


def test_numbering_plan_includes_based_on_dependency_numpr():
    styles_xml = (
        '<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
        '<w:style w:type="paragraph" w:styleId="BaseList">'
        '<w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="7"/></w:numPr></w:pPr>'
        '</w:style>'
        '<w:style w:type="paragraph" w:styleId="DerivedList">'
        '<w:basedOn w:val="BaseList"/>'
        '</w:style>'
        '</w:styles>'
    )
    registry = {
        "numbering": {
            "abstract_nums": [{
                "abstractNumId": 3,
                "xml": '<w:abstractNum w:abstractNumId="3"><w:nsid w:val="ABCDEF01"/></w:abstractNum>',
            }],
            "nums": [{
                "numId": 7,
                "abstractNumId": 3,
                "xml": '<w:num w:numId="7"><w:abstractNumId w:val="3"/></w:num>',
            }],
        }
    }

    plan = build_numbering_import_plan(
        registry,
        styles_xml,
        MINIMAL_TARGET_NUMBERING,
        ["DerivedList"],
    )

    assert len(plan["abstract_nums_to_import"]) == 1
    assert len(plan["nums_to_import"]) == 1
    assert "BaseList" in plan["style_numid_remap"]


def test_numbering_plan_imports_direct_role_pattern_and_preserves_level():
    registry = {
        "numbering": {
            "abstract_nums": [{
                "abstractNumId": 3,
                "xml": (
                    '<w:abstractNum w:abstractNumId="3"><w:nsid w:val="ABCDEF01"/>'
                    '<w:lvl w:ilvl="2"><w:numFmt w:val="lowerLetter"/>'
                    '<w:lvlText w:val="%3."/></w:lvl></w:abstractNum>'
                ),
            }],
            "nums": [{
                "numId": 7,
                "abstractNumId": 3,
                "xml": (
                    '<w:num w:numId="7"><w:abstractNumId w:val="3"/>'
                    '<w:lvlOverride w:ilvl="2"><w:startOverride w:val="4"/>'
                    '</w:lvlOverride></w:num>'
                ),
            }],
        }
    }
    role_specs = {
        "SUBSUBPARAGRAPH": {
            "style_id": "DirectRole",
            "numbering_provenance": "direct_numpr",
            "numbering_pattern": {
                "numId": "7",
                "ilvl": "2",
                "abstractNumId": "3",
                "startOverride": "4",
                "numFmt": "lowerLetter",
                "lvlText": "%3.",
            },
        }
    }
    styles_xml = (
        '<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
        '<w:style w:type="paragraph" w:styleId="DirectRole"/></w:styles>'
    )

    plan = build_numbering_import_plan(
        registry,
        styles_xml,
        MINIMAL_TARGET_NUMBERING,
        ["DirectRole"],
        role_specs=role_specs,
        roles_to_apply=["SUBSUBPARAGRAPH"],
    )

    remap = plan["role_numpr_remap"]["SUBSUBPARAGRAPH"]
    assert remap["old_numId"] == 7
    assert remap["new_numId"] != 7
    assert remap["ilvl"] == 2


def test_numbering_plan_rejects_role_pattern_that_disagrees_with_definition():
    registry = {
        "numbering": {
            "abstract_nums": [{
                "abstractNumId": 3,
                "xml": (
                    '<w:abstractNum w:abstractNumId="3"><w:lvl w:ilvl="0">'
                    '<w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/>'
                    '</w:lvl></w:abstractNum>'
                ),
            }],
            "nums": [{
                "numId": 7,
                "abstractNumId": 3,
                "xml": '<w:num w:numId="7"><w:abstractNumId w:val="3"/></w:num>',
            }],
        }
    }
    role_specs = {
        "PARAGRAPH": {
            "style_id": "DirectRole",
            "numbering_provenance": "direct_numpr",
            "numbering_pattern": {"numId": "7", "ilvl": "0", "lvlText": "WRONG"},
        }
    }
    styles_xml = (
        '<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
        '<w:style w:type="paragraph" w:styleId="DirectRole"/></w:styles>'
    )
    with pytest.raises(ValueError, match="does not match"):
        build_numbering_import_plan(
            registry,
            styles_xml,
            MINIMAL_TARGET_NUMBERING,
            ["DirectRole"],
            role_specs=role_specs,
            roles_to_apply=["PARAGRAPH"],
        )


def test_numbering_plan_imports_additional_header_footer_numid():
    registry = {
        "numbering": {
            "abstract_nums": [_make_abstract_num(3)],
            "nums": [_make_num(7, 3)],
        }
    }
    styles_xml = (
        '<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
        '</w:styles>'
    )
    plan = build_numbering_import_plan(
        registry,
        styles_xml,
        MINIMAL_TARGET_NUMBERING,
        [],
        additional_num_ids=[7],
    )
    assert plan["num_id_remap"][7] == plan["nums_to_import"][0]["new_id"]
