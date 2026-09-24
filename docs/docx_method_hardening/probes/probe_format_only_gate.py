"""Probe: which body mutations does the final Format-only gate accept?

Evidence for WI-02 of the DOCX Method Hardening plan. It builds a minimal
source DOCX, applies one mutation per paragraph to a copy, and runs the
engine's own end-of-pipeline gate (``verify_phase2_invariants`` under the
``format_only`` policy) against source versus mutated output.

Run from the repository root::

    python docs/docx_method_hardening/probes/probe_format_only_gate.py

Before WI-02 every row but the control prints ``ACCEPTED``. After WI-02 every
row must print ``REJECTED``. Keep this file runnable; WI-02 turns each row into
a regression test.
"""

from __future__ import annotations

import sys
import tempfile
import zipfile
from pathlib import Path

REPO_ROOT = Path(__file__).resolve().parents[3]
sys.path.insert(0, str(REPO_ROOT))

from spec_formatter.style_application.phase2_invariants import (  # noqa: E402
    verify_phase2_invariants,
)

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"

CASES = [
    # (label, source paragraph, mutated paragraph, what Word would show)
    (
        "run-level w:tab dropped beside a space",
        '<w:p><w:r><w:t xml:space="preserve">Section </w:t><w:tab/><w:t>Title</w:t></w:r></w:p>',
        '<w:p><w:r><w:t xml:space="preserve">Section </w:t><w:t>Title</w:t></w:r></w:p>',
        "tab stop alignment lost",
    ),
    (
        "tracked-deleted text (w:delText) emptied",
        '<w:p><w:r><w:t xml:space="preserve">Old </w:t></w:r>'
        '<w:del w:id="1" w:author="A" w:date="2026-01-01T00:00:00Z">'
        "<w:r><w:delText>deleted words</w:delText></w:r></w:del>"
        "<w:r><w:t>kept</w:t></w:r></w:p>",
        '<w:p><w:r><w:t xml:space="preserve">Old </w:t></w:r>'
        '<w:del w:id="1" w:author="A" w:date="2026-01-01T00:00:00Z">'
        "<w:r><w:delText></w:delText></w:r></w:del>"
        "<w:r><w:t>kept</w:t></w:r></w:p>",
        "reject-all no longer restores the author's words",
    ),
    (
        "xml:space=preserve removed from a run ending in a space",
        '<w:p><w:r><w:t xml:space="preserve">Trailing </w:t></w:r><w:r><w:t>text</w:t></w:r></w:p>',
        "<w:p><w:r><w:t>Trailing </w:t></w:r><w:r><w:t>text</w:t></w:r></w:p>",
        "Word drops the space on open: 'Trailingtext'",
    ),
    (
        "non-breaking spaces replaced by plain spaces",
        "<w:p><w:r><w:t>SECTION 21 13 13</w:t></w:r></w:p>",
        "<w:p><w:r><w:t>SECTION 21 13 13</w:t></w:r></w:p>",
        "section number can now wrap across lines",
    ),
    (
        "double space collapsed",
        "<w:p><w:r><w:t>Double  space</w:t></w:r></w:p>",
        "<w:p><w:r><w:t>Double space</w:t></w:r></w:p>",
        "one character lost",
    ),
    (
        "w:softHyphen dropped",
        "<w:p><w:r><w:t>Fire</w:t><w:softHyphen/><w:t>proofing</w:t></w:r></w:p>",
        "<w:p><w:r><w:t>Fire</w:t><w:t>proofing</w:t></w:r></w:p>",
        "optional hyphenation point lost",
    ),
    (
        "w:br dropped beside a space",
        '<w:p><w:r><w:t xml:space="preserve">line one </w:t><w:br/><w:t>line two</w:t></w:r></w:p>',
        '<w:p><w:r><w:t xml:space="preserve">line one </w:t><w:t>line two</w:t></w:r></w:p>',
        "manual line break lost",
    ),
    (
        "CONTROL: a visible word changed",
        "<w:p><w:r><w:t>Text</w:t></w:r></w:p>",
        "<w:p><w:r><w:t>Changed</w:t></w:r></w:p>",
        "(rejected today; must stay rejected)",
    ),
]


def write_docx(path: Path, body: str) -> None:
    parts = {
        "[Content_Types].xml": (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
            '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
            '<Default Extension="xml" ContentType="application/xml"/>'
            '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
            '<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>'
            "</Types>"
        ),
        "_rels/.rels": (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            f'<Relationship Id="rId1" Type="{R}/officeDocument" Target="word/document.xml"/>'
            "</Relationships>"
        ),
        "word/_rels/document.xml.rels": (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            f'<Relationship Id="rId1" Type="{R}/styles" Target="styles.xml"/>'
            "</Relationships>"
        ),
        "word/styles.xml": (
            f'<w:styles xmlns:w="{W}"><w:docDefaults><w:rPrDefault><w:rPr/></w:rPrDefault>'
            "<w:pPrDefault><w:pPr/></w:pPrDefault></w:docDefaults>"
            '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>'
            "</w:styles>"
        ),
        "word/document.xml": (
            f'<w:document xmlns:w="{W}" xmlns:r="{R}"><w:body>{body}<w:sectPr/></w:body></w:document>'
        ),
    }
    with zipfile.ZipFile(path, "w") as package:
        for name, text in parts.items():
            package.writestr(name, text)


def main() -> int:
    workspace = Path(tempfile.mkdtemp(prefix="probe-format-only-gate-"))
    accepted = 0
    for index, (label, source_paragraph, mutated_paragraph, effect) in enumerate(CASES):
        source = workspace / f"source-{index}.docx"
        output = workspace / f"output-{index}.docx"
        write_docx(source, source_paragraph)
        write_docx(output, mutated_paragraph)
        new_document_xml = zipfile.ZipFile(output).read("word/document.xml")
        try:
            verify_phase2_invariants(
                source, new_document_xml, new_docx=output, conversion_mode="format_only"
            )
        except Exception as exc:  # noqa: BLE001 - the probe reports, it does not judge
            verdict, note = "REJECTED", str(exc).splitlines()[0][:90]
        else:
            verdict, note = "ACCEPTED", effect
            if not label.startswith("CONTROL"):
                accepted += 1
        print(f"{verdict:8}  {label:55}  {note}")
    print(f"\n{accepted} silent mutation(s) accepted by the Format-only gate.")
    return 1 if accepted else 0


if __name__ == "__main__":
    sys.exit(main())
