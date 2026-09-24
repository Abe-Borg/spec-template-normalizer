"""Probe: Format-only style import against a current-Word docDefaults block.

Evidence for WI-01 of the DOCX Method Hardening plan. Word 365 writes
``<w14:ligatures w14:val="standardContextual"/>`` into
``docDefaults/rPrDefault/rPr`` of every new document. The bundle's portable
stylesheet carries the architect's styles verbatim, so that element reaches
the target style importer, whose Format-only materialization parses each
run-property block with only the ``w`` namespace declared.

Run from the repository root::

    python docs/docx_method_hardening/probes/probe_style_import_w14.py

Before WI-01 the ``format_only`` row fails with ``unbound prefix`` while the
``csi_to_canadian`` row succeeds. After WI-01 both rows must print ``OK`` and
the output stylesheet must parse. Keep this file runnable; WI-01 turns it into
a regression test.
"""

from __future__ import annotations

import sys
import tempfile
import xml.etree.ElementTree as ET
from pathlib import Path

REPO_ROOT = Path(__file__).resolve().parents[3]
sys.path.insert(0, str(REPO_ROOT))

from spec_formatter.style_application.core.style_import import (  # noqa: E402
    import_arch_styles_into_target,
)

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
W14 = "http://schemas.microsoft.com/office/word/2010/wordml"
LIGATURES = '<w14:ligatures w14:val="standardContextual"/>'

ARCHITECT_PORTABLE_STYLES = (
    f'<w:styles xmlns:w="{W}" xmlns:w14="{W14}">'
    "<w:docDefaults><w:rPrDefault><w:rPr>"
    f'<w:rFonts w:ascii="Aptos"/><w:sz w:val="24"/>{LIGATURES}'
    "</w:rPr></w:rPrDefault>"
    '<w:pPrDefault><w:pPr><w:spacing w:after="160"/></w:pPr></w:pPrDefault></w:docDefaults>'
    '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>'
    '<w:style w:type="paragraph" w:styleId="CSI_Part__ARCH"><w:name w:val="CSI Part"/>'
    '<w:basedOn w:val="Normal"/><w:pPr><w:keepNext/></w:pPr><w:rPr><w:b/></w:rPr></w:style>'
    "</w:styles>"
)

# A target whose stylesheet root declares only the main namespace, as files
# produced by older Word versions and by third-party generators do.
TARGET_STYLES = (
    f'<w:styles xmlns:w="{W}">'
    "<w:docDefaults><w:rPrDefault><w:rPr>"
    '<w:rFonts w:ascii="Times New Roman"/></w:rPr></w:rPrDefault></w:docDefaults>'
    '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>'
    "</w:styles>"
)


def main() -> int:
    failures = 0
    for label, body_roots in (("format_only", {"CSI_Part__ARCH"}), ("csi_to_canadian", None)):
        extract_dir = Path(tempfile.mkdtemp(prefix=f"probe-w14-{label}-"))
        (extract_dir / "word").mkdir()
        (extract_dir / "word" / "styles.xml").write_text(TARGET_STYLES, encoding="utf-8")
        try:
            result = import_arch_styles_into_target(
                target_extract_dir=extract_dir,
                arch_styles_xml=ARCHITECT_PORTABLE_STYLES,
                needed_style_ids=["CSI_Part__ARCH"],
                log=[],
                style_numid_remap={},
                format_only_body_style_ids=body_roots,
                shell_style_ids=set(),
                namespace_seed="probe",
            )
            ET.fromstring((extract_dir / "word" / "styles.xml").read_text(encoding="utf-8"))
        except Exception as exc:  # noqa: BLE001 - the probe reports, it does not judge
            failures += 1
            print(f"{label:16} FAIL {type(exc).__name__}: {str(exc)[:120]}")
        else:
            print(f"{label:16} OK   {result.style_id_map}")
    return 1 if failures else 0


if __name__ == "__main__":
    sys.exit(main())
