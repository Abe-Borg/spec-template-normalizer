#!/usr/bin/env python3
"""
Phase 1 smoke test:
- runs normalize-slim
- runs apply-instructions
- validates arch_style_registry.json schema
- validates no header/footer and sectPr drift (indirectly via apply-instructions invariants)

Usage:
  python phase1_smoke_test.py MySpec.docx instructions.json
"""

from __future__ import annotations
import json
import subprocess
import sys
from pathlib import Path
import time



def validate_registry(path: Path) -> None:
    data = json.loads(path.read_text(encoding="utf-8"))

    if not isinstance(data, dict):
        raise ValueError("registry must be a JSON object")

    if data.get("version") != 1:
        raise ValueError("registry.version must be 1")

    if not isinstance(data.get("source_docx"), str) or not data["source_docx"]:
        raise ValueError("registry.source_docx must be a non-empty string")

    roles = data.get("roles")
    if not isinstance(roles, dict):
        raise ValueError("registry.roles must be an object")

    # SectionID is optional per the JSON schema — not all spec sections
    # have a separate section-number paragraph (some combine ID + title).
    required_roles = {
        "SectionTitle",
        "PART",
        "ARTICLE",
        "PARAGRAPH",
        "SUBPARAGRAPH",
        "SUBSUBPARAGRAPH",
    }
    missing = required_roles - set(roles.keys())
    if missing:
        raise ValueError(f"registry.roles missing: {sorted(missing)}")

    for role, spec in roles.items():
        if not isinstance(spec, dict):
            raise ValueError(f"roles['{role}'] must be an object")
        if not isinstance(spec.get("style_id"), str) or not spec["style_id"]:
            raise ValueError(f"roles['{role}'].style_id must be a non-empty string")
        if not isinstance(spec.get("exemplar_paragraph_index"), int) or spec["exemplar_paragraph_index"] < 0:
            raise ValueError(f"roles['{role}'].exemplar_paragraph_index must be a non-negative int")
        if "style_name" in spec and spec["style_name"] is not None and not isinstance(spec["style_name"], str):
            raise ValueError(f"roles['{role}'].style_name must be a string or null")


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

    # IMPORTANT: assumes you renamed/overwrote docx_decomposer.py with lean Phase-1 script
    exe = [sys.executable, "docx_decomposer.py", str(docx)]

    stamp = time.strftime("%Y%m%d_%H%M%S")
    extract_dir = Path(f"{docx.stem}_extracted__smoke__{stamp}")

    subprocess.check_call(exe + ["--extract-dir", str(extract_dir), "--normalize-slim"])
    subprocess.check_call(exe + ["--extract-dir", str(extract_dir), "--apply-instructions", str(instr)])

    reg = extract_dir / "arch_style_registry.json"

    if not reg.exists():
        raise FileNotFoundError(f"Expected registry at: {reg}")

    validate_registry(reg)
    print("✅ Phase 1 smoke test: PASS")


if __name__ == "__main__":
    run()
