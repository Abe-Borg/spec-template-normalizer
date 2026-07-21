"""Frozen-app entry point for the Windows PyInstaller build.

PyInstaller freezes a *script*, not a module, so this thin wrapper calls the
GUI's ``main()`` (the repo uses a flat layout, so ``gui.py`` sits at the repo
root and is importable as ``gui``). It also adds two headless flags the release
workflow uses to smoke-test the frozen executable without opening a window:

    SpecificationFormatter.exe --version     print the version and exit
    SpecificationFormatter.exe --selfcheck   import the app's heavy modules --
                                             proving PyInstaller bundled every
                                             hidden import -- and exit 0 (non-zero
                                             on any import error)

The GUI build is windowed (``console=False``), so ``sys.stdout`` may be ``None``
in the frozen app; ``_emit`` writes results to the file named by
``SPEC_FORMATTER_SELFCHECK_OUT`` (set by CI) as well as printing when it can, so
the smoke step can read the outcome regardless.
"""
from __future__ import annotations

import os
import sys

SELFCHECK_OUT_ENV = "SPEC_FORMATTER_SELFCHECK_OUT"


def _emit(message: str) -> None:
    try:
        if sys.stdout is not None:
            print(message)
    except Exception:
        pass
    out = os.environ.get(SELFCHECK_OUT_ENV)
    if out:
        try:
            with open(out, "w", encoding="utf-8") as fh:
                fh.write(message + "\n")
        except OSError:
            pass


def _print_version() -> int:
    import spec_formatter

    _emit(spec_formatter.__version__)
    return 0


def _selfcheck() -> int:
    try:
        import spec_formatter
        from spec_formatter import pipeline  # noqa: F401 - proves the engine froze
        from spec_formatter import updates  # noqa: F401 - proves the updater froze
        from spec_formatter import secrets  # noqa: F401 - proves the keyring wrapper froze
        import gui  # noqa: F401 - pulls customtkinter
        import keyring.backends.Windows  # noqa: F401 - Windows Credential Manager backend
    except Exception:
        import traceback

        _emit("SELFCHECK FAILED:\n" + traceback.format_exc())
        return 1
    _emit(f"SpecificationFormatter {spec_formatter.__version__} selfcheck ok")
    return 0


def main() -> int:
    args = sys.argv[1:]
    if "--version" in args:
        return _print_version()
    if "--selfcheck" in args:
        return _selfcheck()
    from gui import main as gui_main

    gui_main()
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
