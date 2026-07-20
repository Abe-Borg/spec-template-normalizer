"""Namespaced access to the architect-template analysis engine.

The legacy engine remains at the repository root for compatibility.  Put that
root first only while importing it so a neighboring checkout with same-named
modules cannot satisfy its absolute imports.
"""

from pathlib import Path
import sys


_PROJECT_ROOT = Path(__file__).resolve().parents[1]
_project_root_text = str(_PROJECT_ROOT)
_inserted_project_root = not sys.path or sys.path[0] != _project_root_text
if _inserted_project_root:
    sys.path.insert(0, _project_root_text)

try:
    from phase1_bundle import (
        BundleManifest,
        load_bundle_manifest,
        sha256_file,
        validate_bundle_directory,
    )
    from phase1_pipeline import (
        DEFAULT_MODEL,
        PIPELINE_VERSION,
        Phase1Result,
        run_phase1,
    )
finally:
    if _inserted_project_root:
        try:
            sys.path.remove(_project_root_text)
        except ValueError:  # pragma: no cover - defensive import cleanup
            pass

__all__ = [
    "BundleManifest",
    "DEFAULT_MODEL",
    "PIPELINE_VERSION",
    "Phase1Result",
    "load_bundle_manifest",
    "run_phase1",
    "sha256_file",
    "validate_bundle_directory",
]
