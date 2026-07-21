# -*- mode: python ; coding: utf-8 -*-
"""PyInstaller one-folder build for the Specification Formatter Windows app.

Build (on Windows, from the repo root) with:

    pip install -r requirements.txt -r requirements-build.txt
    pyinstaller packaging/windows/specification-formatter.spec --noconfirm --clean

Output: ``dist/SpecificationFormatter/`` (a folder containing
``SpecificationFormatter.exe`` plus its bundled interpreter and dependencies).
``packaging/windows/installer.iss`` wraps that folder into
``SpecificationFormatterSetup.exe``.

The spec is exec'd by PyInstaller with globals like ``Analysis``/``PYZ``/``EXE``/
``COLLECT``/``SPECPATH`` injected -- that is why linters flag "undefined name"
here.

One-folder (not one-file) is deliberate: it starts faster, updates more
reliably, and trips antivirus far less than a self-extracting one-file exe --
and the Inno Setup installer makes it a normal double-click "install" for the
user regardless.
"""
import os

from PyInstaller.utils.hooks import collect_all, collect_submodules, copy_metadata

# The repo uses a flat layout: gui.py and the phase1_* modules live at the repo
# root (two levels up from this spec), alongside the spec_formatter package.
# pathex must include the root so ``import gui`` / ``import spec_formatter``
# resolve when PyInstaller analyses app_entry.py.
_REPO_ROOT = os.path.abspath(os.path.join(SPECPATH, "..", ".."))

datas = []
binaries = []
hiddenimports = []

# customtkinter ships theme JSON assets that must be collected or the app
# renders wrong at runtime.
_d, _b, _h = collect_all("customtkinter")
datas += _d
binaries += _b
hiddenimports += _h

# keyring resolves its backend (Windows Credential Manager) dynamically; bundle
# every backend plus the metadata it reads to enumerate them via entry points.
hiddenimports += collect_submodules("keyring.backends")
hiddenimports += ["keyring.backends.Windows"]

# The anthropic SDK is imported *inside functions* (a deferred import) across the
# codebase, so PyInstaller's static analysis of the entry script won't discover
# it. Pull the whole package in, plus its metadata (the SDK reads its own version
# via importlib.metadata to build the request User-Agent).
hiddenimports += collect_submodules("anthropic")
for _dist in ("anthropic", "keyring"):
    try:
        datas += copy_metadata(_dist)
    except Exception:
        # A missing dist here is non-fatal -- the app still runs.
        pass

# The app package: submodules + package data (the phase2 prompt files under
# spec_formatter/style_application/core/prompts).
_d, _b, _h = collect_all("spec_formatter")
datas += _d
binaries += _b
hiddenimports += _h

# Root-level modules (flat layout) that the pipeline may reach through deferred
# imports -- bundle them explicitly so a frozen run can never miss one.
hiddenimports += [
    "gui",
    "phase1_pipeline",
    "phase1_bundle",
    "phase1_validator",
    "docx_decomposer",
    "llm_classifier",
    "paragraph_rules",
    "arch_env_extractor",
    "ooxml_text",
]

# Prompt payloads that ship at the bundle root (they live at the repo root, not
# inside the package, so collect_all above does not pick them up).
datas += [
    (os.path.join(_REPO_ROOT, "master_prompt.txt"), "."),
    (os.path.join(_REPO_ROOT, "run_instruction_prompt.txt"), "."),
    (
        os.path.join(_REPO_ROOT, "spec_formatter", "style_application", "core", "prompts"),
        os.path.join("spec_formatter", "style_application", "core", "prompts"),
    ),
]

a = Analysis(
    [os.path.join(SPECPATH, "app_entry.py")],
    pathex=[_REPO_ROOT],
    binaries=binaries,
    datas=datas,
    hiddenimports=hiddenimports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    # Test-only / dev-only packages must never be pulled into the shipped app.
    excludes=["pytest", "_pytest", "coverage"],
    noarchive=False,
)

pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name="SpecificationFormatter",
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,
    console=False,  # windowed GUI app -- no console window behind it
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon=None,  # drop an .ico here (icon="app.ico") once one exists
)

coll = COLLECT(
    exe,
    a.binaries,
    a.datas,
    strip=False,
    upx=False,
    upx_exclude=[],
    name="SpecificationFormatter",
)
