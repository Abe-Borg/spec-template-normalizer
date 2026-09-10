"""Coverage for the geometry-proofing tool's comparison logic.

Rendering needs LibreOffice and Poppler and is therefore not exercised here.
The comparison is, because it is the part that can be wrong quietly: a
displacement check that matches nothing reports a clean run it never performed,
which is worse than no check at all.
"""

from __future__ import annotations

import importlib.util
import sys
from pathlib import Path

import pytest

_MODULE_PATH = Path(__file__).resolve().parent.parent / "scripts" / "proof_render.py"
_spec = importlib.util.spec_from_file_location("proof_render", _MODULE_PATH)
proof_render = importlib.util.module_from_spec(_spec)
# Registered before execution because @dataclass resolves its own module by
# name while the class body runs.
sys.modules["proof_render"] = proof_render
_spec.loader.exec_module(proof_render)

Word = proof_render.Word


def _row(order: int, text: str, x: float, y: float = 100.0) -> Word:
    return Word(page=0, order=order, text=text, x=x, y=y)


def test_identical_layout_reports_success(capsys) -> None:
    words = [_row(0, "PART", 72.0), _row(1, "GENERAL", 108.0)]
    assert proof_render._compare(words, list(words), limit=5) == 0
    assert "no word moved" in capsys.readouterr().out


def test_horizontal_displacement_is_reported(capsys) -> None:
    """The signature of indentation lost with its numbering."""

    reference = [_row(0, "A.", 36.0), _row(1, "Section", 62.95)]
    candidate = [_row(0, "A.", 0.0), _row(1, "Section", 36.0)]
    assert proof_render._compare(reference, candidate, limit=5) == 1
    out = capsys.readouterr().out
    assert "2 words moved" in out
    # Reported in twips too, because that is the unit w:ind is written in.
    assert "-720" in out


def test_sub_half_point_jitter_is_not_reported(capsys) -> None:
    reference = [_row(0, "PART", 72.0)]
    candidate = [_row(0, "PART", 72.3)]
    assert proof_render._compare(reference, candidate, limit=5) == 0


def test_unmatched_words_are_surfaced_not_silently_skipped(capsys) -> None:
    """A word that cannot be matched must not read as a word that did not move.

    Without this the tool reports "no word moved by more than 0.5pt" on two
    documents that share no text at all.
    """

    reference = [_row(0, "PART", 72.0), _row(1, "GENERAL", 108.0)]
    candidate = [_row(0, "PART", 72.0), _row(1, "PRODUCTS", 108.0)]
    proof_render._compare(reference, candidate, limit=5)
    assert "1 words did not match" in capsys.readouterr().out


def test_missing_renderer_exits_rather_than_passing() -> None:
    """A proofing tool that passes because it did nothing is the failure mode."""

    with pytest.raises(SystemExit) as raised:
        proof_render._require("definitely-not-a-real-binary-name")
    assert "cannot report a clean result" in str(raised.value)
