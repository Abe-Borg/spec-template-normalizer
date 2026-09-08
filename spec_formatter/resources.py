"""Where the application's shipped resources live, frozen or from a checkout.

The architect prompts (``master_prompt.txt``, ``run_instruction_prompt.txt``),
the phase-2 prompts under ``spec_formatter/style_application/core/prompts``,
``LICENSE`` and ``THIRD_PARTY_NOTICES.md`` are data files, not modules. In a
checkout they sit under the repository root. In the PyInstaller one-folder
build (``packaging/windows/specification-formatter.spec``) they are copied
into the bundle directory that PyInstaller exposes as ``sys._MEIPASS`` (the
app's ``_internal`` folder), and the same layout is kept below it, so one
root resolves every resource in both worlds.

Every module that needs a shipped file resolves it through :func:`resource_root`
instead of ``Path(__file__)`` arithmetic, which only works in the frozen app by
coincidence of PyInstaller's module layout. The frozen ``--selfcheck`` reads
the prompts through the same helper, so a bundle that lost a data file fails
in CI rather than on a user's first run.
"""

from __future__ import annotations

import sys
from pathlib import Path

_REPO_ROOT = Path(__file__).resolve().parents[1]

#: Prompt files the architect analysis reads from :func:`architect_prompt_dir`.
ARCHITECT_PROMPT_FILES: tuple[str, ...] = (
    "master_prompt.txt",
    "run_instruction_prompt.txt",
)

#: Prompt files the target classifier reads from :func:`target_prompt_dir`.
TARGET_PROMPT_FILES: tuple[str, ...] = (
    "phase2_master_prompt.txt",
    "phase2_run_instruction.txt",
)


def is_frozen() -> bool:
    """Return True inside a PyInstaller bundle."""

    return bool(getattr(sys, "frozen", False)) and hasattr(sys, "_MEIPASS")


def resource_root() -> Path:
    """Return the directory that holds the shipped resource files."""

    if is_frozen():
        return Path(getattr(sys, "_MEIPASS")).resolve()
    return _REPO_ROOT


def architect_prompt_dir() -> Path:
    """Directory holding the architect-analysis prompt files."""

    return resource_root()


def target_prompt_dir() -> Path:
    """Directory holding the target-classification prompt files."""

    return resource_root() / "spec_formatter" / "style_application" / "core" / "prompts"


def resource_path(*parts: str) -> Path:
    """Return ``resource_root()`` joined with *parts* (for LICENSE and notices)."""

    return resource_root().joinpath(*parts)


__all__ = [
    "ARCHITECT_PROMPT_FILES",
    "TARGET_PROMPT_FILES",
    "architect_prompt_dir",
    "is_frozen",
    "resource_path",
    "resource_root",
    "target_prompt_dir",
]
