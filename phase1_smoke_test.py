#!/usr/bin/env python3
"""
Phase 1 smoke test:
- extracts DOCX and builds slim bundle
- applies instructions
- validates arch_style_registry.json schema
- validates no header/footer and sectPr drift (indirectly via apply-instructions invariants)

Usage:
  python phase1_smoke_test.py MySpec.docx instructions.json
"""

from __future__ import annotations
import json
import sys
from pathlib import Path
import time

from docx_decomposer import (
    extract_docx,
    build_slim_bundle,
    apply_instructions,
    emit_arch_style_registry,
)
from phase1_validator import validate_style_registry


def run() -> None:
    if len(sys.argv) != 3:
        print("Usage: python phase1_smoke_test.py <MySpec.docx> <instructions.json>")
        sys.exit(2)

    docx = Path(sys.argv[1])
    instr = Path(sys.argv[2])
    if not docx.exists():
        raise FileNotFoundError(docx)
    if not instr.exists():
        raise FileNotFoundError(instr)

    stamp = time.strftime("%Y%m%d_%H%M%S")
    extract_dir = Path(f"{docx.stem}_extracted__smoke__{stamp}")

    # Extract and build slim bundle
    extract_docx(docx, extract_dir)
    bundle = build_slim_bundle(extract_dir)
    (extract_dir / "slim_bundle.json").write_text(
        json.dumps(bundle, indent=2), encoding="utf-8"
    )

    # Load instructions and apply
    instructions = json.loads(instr.read_text(encoding="utf-8"))
    apply_instructions(extract_dir, instructions)
    emit_arch_style_registry(extract_dir, docx.name, instructions)

    reg = extract_dir / "arch_style_registry.json"

    if not reg.exists():
        raise FileNotFoundError(f"Expected registry at: {reg}")

    data = json.loads(reg.read_text(encoding="utf-8"))
    validate_style_registry(data)
    print("✅ Phase 1 smoke test: PASS")


if __name__ == "__main__":
    run()
