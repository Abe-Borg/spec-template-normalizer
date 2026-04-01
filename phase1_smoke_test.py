#!/usr/bin/env python3
"""
Phase 1 smoke test:
- extracts DOCX and builds slim bundle
- applies instructions
- validates both arch_style_registry.json and arch_template_registry.json
- validates cross-registry consistency (style IDs exist in template registry)
- validates no header/footer and sectPr drift (indirectly via apply-instructions invariants)

Usage:
  python phase1_smoke_test.py MySpec.docx instructions.json
"""

from __future__ import annotations
import json
import shutil
import sys
from pathlib import Path
import time

from docx_decomposer import (
    extract_docx,
    build_slim_bundle,
    apply_instructions,
    build_style_registry_dict,
)
from arch_env_extractor import extract_arch_template_registry
from phase1_validator import validate_phase1_contracts


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

    # Build both registries in memory (mirrors gui.py pipeline)
    style_registry = build_style_registry_dict(
        extract_dir,
        docx.name,
        instructions,
        pre_apply_bundle=bundle,
    )
    template_registry = extract_arch_template_registry(extract_dir, docx)

    # Validate both registries + cross-registry consistency
    validate_phase1_contracts(style_registry, template_registry)

    # Write both to disk
    reg_path = extract_dir / "arch_style_registry.json"
    reg_path.write_text(json.dumps(style_registry, indent=2), encoding="utf-8")

    env_path = extract_dir / "arch_template_registry.json"
    env_path.write_text(json.dumps(template_registry, indent=2), encoding="utf-8")

    raw_styles_src = extract_dir / "word" / "styles.xml"
    raw_settings_src = extract_dir / "word" / "settings.xml"
    raw_styles_dst = extract_dir / "arch_styles_raw.xml"
    raw_settings_dst = extract_dir / "arch_settings_raw.xml"

    if raw_styles_src.exists():
        shutil.copy2(raw_styles_src, raw_styles_dst)
        print(f"Preserved raw styles.xml as {raw_styles_dst.name}")
    if raw_settings_src.exists():
        shutil.copy2(raw_settings_src, raw_settings_dst)
        print(f"Preserved raw settings.xml as {raw_settings_dst.name}")

    # Verify round-trip from disk
    for path in (reg_path, env_path, raw_styles_dst, raw_settings_dst):
        if not path.exists():
            raise FileNotFoundError(f"Expected artifact at: {path}")
        if path.suffix == ".json":
            data = json.loads(path.read_text(encoding="utf-8"))
            if not isinstance(data, dict) or not data:
                raise ValueError(f"Artifact is empty or not a JSON object: {path}")

    print("Phase 1 smoke test: PASS (both registries validated)")


if __name__ == "__main__":
    run()
