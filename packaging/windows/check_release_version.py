"""Fail a release build unless the git tag matches the package version literal.

The shipped app reports ``spec_formatter/__init__.py::__version__`` -- the value
the updater compares against the manifest. This repo keeps a SINGLE source of
truth (there is no ``pyproject.toml``), so the guard only has to confirm that one
literal equals the tag. A tag that drifts from ``__version__`` would ship an
installer stuck in a perpetual "update available" loop: the published manifest
version would be the tag, but the installed app would keep reporting the stale
``__version__``.

Pure standard library; reads the file with ``ast`` without importing the
package, so it runs before any ``pip install``.

Usage:  python packaging/windows/check_release_version.py --tag v1.2.3
"""
from __future__ import annotations

import argparse
import ast
import pathlib
import sys

_ROOT = pathlib.Path(__file__).resolve().parents[2]
_INIT_REL = "spec_formatter/__init__.py"


def init_version(root: pathlib.Path = _ROOT) -> str:
    src = (root / "spec_formatter" / "__init__.py").read_text(encoding="utf-8")
    for node in ast.walk(ast.parse(src)):
        if isinstance(node, ast.Assign) and any(
            isinstance(t, ast.Name) and t.id == "__version__" for t in node.targets
        ):
            value = node.value
            if isinstance(value, ast.Constant) and isinstance(value.value, str):
                return value.value
    raise SystemExit(f"could not find a string __version__ in {_INIT_REL}")


def check(tag: str, *, root: pathlib.Path = _ROOT) -> list:
    """Return a list of mismatch messages (empty when the tag matches)."""
    tag_version = tag[1:] if tag.startswith("v") else tag
    problems = []
    init = init_version(root)
    if init != tag_version:
        problems.append(f"{_INIT_REL} __version__ {init!r} != tag {tag_version!r}")
    return problems


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="Guard tag == package version literal.")
    parser.add_argument("--tag", required=True, help="git tag, e.g. v1.2.3")
    args = parser.parse_args(argv)

    tag_version = args.tag[1:] if args.tag.startswith("v") else args.tag
    print(f"tag={tag_version} {_INIT_REL}={init_version()}")
    problems = check(args.tag)
    if problems:
        for problem in problems:
            print("ERROR:", problem, file=sys.stderr)
        print(
            f"Bump __version__ in {_INIT_REL} to match the tag, then re-tag.",
            file=sys.stderr,
        )
        return 1
    print("release version guard: OK")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
