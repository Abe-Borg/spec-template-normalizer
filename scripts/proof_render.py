#!/usr/bin/env python3
"""Compare where words actually land in two rendered DOCX files.

The engine's invariants are XML truth: they prove what the markup says. This
proves what a reader sees, which is a different claim and occasionally a
different answer. Use it before trusting a formatting change on live work.

    python scripts/proof_render.py reference.docx candidate.docx

It renders both documents with LibreOffice, extracts per-word bounding boxes
with ``pdftotext -bbox``, matches words by page and reading order, and reports
the largest horizontal and vertical displacements. A change that preserves
content but moves it -- the class that motivated the geometry invariant -- shows
up here as a non-zero dx.

**This is a proofing aid, not an oracle, and deliberately not a CI gate.**
LibreOffice does not implement pStyle-linked numbering levels: a style whose
``w:numPr`` omits an explicit ``w:ilvl`` renders at level 0 there and at its
real level in Word. Where the two disagree about a numbered document, Word is
right. Read a non-zero result as "look at this in Word", never as a verdict.

Requires ``soffice`` and ``pdftotext`` on PATH. It fails loudly when they are
missing rather than reporting a clean run it never performed -- a proofing tool
that passes because it did nothing is worse than no proofing tool.
"""

from __future__ import annotations

import argparse
import shutil
import subprocess
import sys
import tempfile
import xml.etree.ElementTree as ET
from dataclasses import dataclass
from pathlib import Path

#: Twips per point, for reporting in the same unit as w:ind values.
_POINTS_PER_TWIP = 1 / 20


@dataclass(frozen=True)
class Word:
    page: int
    order: int
    text: str
    x: float
    y: float


def _require(tool: str) -> str:
    path = shutil.which(tool)
    if path is None:
        sys.exit(
            f"error: {tool!r} is not on PATH. This check renders documents; it "
            "cannot report a clean result without actually rendering them."
        )
    return path


def _render_pdf(docx: Path, out_dir: Path) -> Path:
    subprocess.run(
        [
            _require("soffice"),
            "--headless",
            f"-env:UserInstallation=file://{out_dir / 'profile'}",
            "--convert-to",
            "pdf",
            "--outdir",
            str(out_dir),
            str(docx),
        ],
        check=True,
        capture_output=True,
        timeout=600,
    )
    pdf = out_dir / f"{docx.stem}.pdf"
    if not pdf.is_file():
        sys.exit(f"error: LibreOffice produced no PDF for {docx.name}")
    return pdf


def _words(pdf: Path) -> list[Word]:
    result = subprocess.run(
        [_require("pdftotext"), "-bbox", str(pdf), "-"],
        check=True,
        capture_output=True,
        timeout=300,
    )
    root = ET.fromstring(result.stdout)
    namespace = {"x": "http://www.w3.org/1999/xhtml"}
    words: list[Word] = []
    for page_number, page in enumerate(root.iter(f"{{{namespace['x']}}}page")):
        for order, word in enumerate(page.iter(f"{{{namespace['x']}}}word")):
            words.append(
                Word(
                    page=page_number,
                    order=order,
                    text=(word.text or "").strip(),
                    x=float(word.attrib["xMin"]),
                    y=float(word.attrib["yMin"]),
                )
            )
    return words


def _compare(reference: list[Word], candidate: list[Word], limit: int) -> int:
    by_key = {(w.page, w.order): w for w in candidate}
    displaced: list[tuple[float, float, Word, Word]] = []
    unmatched = 0
    for word in reference:
        other = by_key.get((word.page, word.order))
        if other is None or other.text != word.text:
            unmatched += 1
            continue
        dx, dy = other.x - word.x, other.y - word.y
        if abs(dx) > 0.5 or abs(dy) > 0.5:
            displaced.append((dx, dy, word, other))

    print(f"reference words: {len(reference)}   candidate words: {len(candidate)}")
    if unmatched:
        print(
            f"warning: {unmatched} words did not match by position and text; "
            "the documents may differ in content, not only in layout"
        )
    if not displaced:
        print("no word moved by more than 0.5pt")
        return 0

    displaced.sort(key=lambda item: -max(abs(item[0]), abs(item[1])))
    print(f"\n{len(displaced)} words moved. Largest {min(limit, len(displaced))}:\n")
    print(f"{'page':>5} {'dx pt':>9} {'dy pt':>9} {'dx twips':>9}  word")
    for dx, dy, word, _other in displaced[:limit]:
        print(
            f"{word.page:>5} {dx:>9.2f} {dy:>9.2f} {dx / _POINTS_PER_TWIP:>9.0f}  "
            f"{word.text[:40]}"
        )
    print(
        "\nA non-zero dx on numbered paragraphs is the signature of indentation "
        "lost with its numbering. Confirm in Word before acting on it."
    )
    return 1


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    parser.add_argument("reference", type=Path, help="the document as it should render")
    parser.add_argument("candidate", type=Path, help="the document to check")
    parser.add_argument(
        "--limit", type=int, default=20, help="how many displaced words to list"
    )
    args = parser.parse_args()

    for path in (args.reference, args.candidate):
        if not path.is_file():
            sys.exit(f"error: no such file: {path}")

    with tempfile.TemporaryDirectory() as tmp:
        out = Path(tmp)
        reference_pdf = _render_pdf(args.reference, out / "reference")
        candidate_pdf = _render_pdf(args.candidate, out / "candidate")
        return _compare(_words(reference_pdf), _words(candidate_pdf), args.limit)


if __name__ == "__main__":
    raise SystemExit(main())
