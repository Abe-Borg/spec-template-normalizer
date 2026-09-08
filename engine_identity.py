"""Committed fingerprint of the architect-analysis engine.

The profile cache key already covers the source hash, the pipeline version,
the classifier model, and the prompt hashes, but not the code that turns a
classification into a profile: the response schema and repair logic in
``llm_classifier.py``, the text-signal rules in ``paragraph_rules.py``, the
shell capture policy in ``arch_env_extractor.py``, the profile builder in
``docx_decomposer.py``, and the cross-contract validator in
``phase1_validator.py``. A change to any of them used to be protected only
by a hand-bumped ``PIPELINE_VERSION``.

Hashing those files at runtime cannot work in the frozen Windows build (the
modules live inside the PyInstaller archive), so the digest is committed here
and ``tests/test_engine_identity.py`` recomputes it from the checkout: the
test fails whenever one of the files changes without this constant being
updated, and ``python engine_identity.py`` prints the new value.
"""

from __future__ import annotations

import hashlib
from pathlib import Path
from typing import Iterable, Optional

#: Source files whose behaviour shapes every published profile, relative to
#: the repository root. Keep sorted; the digest depends on this order.
ENGINE_SOURCE_FILES: tuple[str, ...] = (
    "arch_env_extractor.py",
    "docx_decomposer.py",
    "llm_classifier.py",
    "paragraph_rules.py",
    "phase1_validator.py",
)

#: First 16 hex digits of the SHA-256 over ENGINE_SOURCE_FILES. Update with
#: ``python engine_identity.py`` whenever one of those files changes.
ENGINE_SOURCE_DIGEST = "08daea8f8fe2c6d4"

ENGINE_SOURCE_DIGEST_LENGTH = 16


def compute_engine_source_digest(
    root: Optional[Path] = None,
    files: Iterable[str] = ENGINE_SOURCE_FILES,
) -> str:
    """Recompute the digest from a source checkout.

    Line endings are normalised so a Windows checkout with ``autocrlf``
    produces the same digest as a POSIX checkout.
    """

    base = Path(root) if root is not None else Path(__file__).resolve().parent
    digest = hashlib.sha256()
    for name in files:
        data = (base / name).read_bytes().replace(b"\r\n", b"\n")
        digest.update(name.encode("utf-8"))
        digest.update(b"\0")
        digest.update(str(len(data)).encode("ascii"))
        digest.update(b"\0")
        digest.update(data)
        digest.update(b"\0")
    return digest.hexdigest()[:ENGINE_SOURCE_DIGEST_LENGTH]


if __name__ == "__main__":  # pragma: no cover - developer convenience
    print(compute_engine_source_digest())
