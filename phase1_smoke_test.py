#!/usr/bin/env python3
"""Offline smoke test for the supported immutable Phase 1 bundle pipeline.

Usage:
    python phase1_smoke_test.py TEMPLATE.docx instructions.json [OUTPUT_DIR]

The supplied instructions replace the live classifier call.  All normal
runtime validation, environment capture, audit generation, checksum checks,
and atomic publication still run.
"""

from __future__ import annotations

import json
import sys
from pathlib import Path
from typing import Any, Dict

from phase1_bundle import validate_bundle_directory
from phase1_pipeline import run_phase1


def run() -> None:
    if len(sys.argv) not in {3, 4}:
        print(
            "Usage: python phase1_smoke_test.py "
            "<template.docx> <instructions.json> [output_dir]"
        )
        raise SystemExit(2)

    source_docx = Path(sys.argv[1])
    instruction_path = Path(sys.argv[2])
    output_root = Path(sys.argv[3]) if len(sys.argv) == 4 else source_docx.parent
    if not source_docx.is_file():
        raise FileNotFoundError(source_docx)
    if not instruction_path.is_file():
        raise FileNotFoundError(instruction_path)

    instructions: Dict[str, Any] = json.loads(
        instruction_path.read_text(encoding="utf-8")
    )

    def supplied_classifier(**_kwargs: Any) -> Dict[str, Any]:
        return instructions

    result = run_phase1(
        source_docx=source_docx,
        output_root=output_root,
        api_key="",
        classifier=supplied_classifier,
        progress=print,
    )
    validate_bundle_directory(
        result.bundle_dir,
        expected_source_sha256=result.source_sha256,
        reject_unlisted=True,
    )
    print(
        "Phase 1 smoke test: PASS — "
        f"{result.handled_paragraphs}/{result.classifiable_paragraphs} candidate paragraphs; "
        f"bundle={result.bundle_dir}"
    )


if __name__ == "__main__":
    run()
