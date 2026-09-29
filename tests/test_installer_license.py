"""The Windows installer describes the license and requires the user to accept it.

``packaging/windows/installer.iss`` is compiled only on the Windows release
runner, so nothing else notices if the License Agreement page is dropped,
skipped, or preselected. These tests read the script as text and hold it to
three facts:

* ``LicenseFile`` points at this repository's ``LICENSE``, the PolyForm
  Noncommercial text, in a form Inno Setup displays correctly;
* the page's own label describes that license in plain words, including that
  commercial use needs a separate license;
* nothing in the script skips the page or preselects "I accept", and the
  in-app updater launches the installer interactively, so an update shows the
  page too.

Inno Setup itself does the rest: with ``LicenseFile`` set, "I do not accept" is
preselected and Next stays disabled until the user picks "I accept".
"""

from __future__ import annotations

import re
import sys
from pathlib import Path, PureWindowsPath

from spec_formatter import updates

REPO_ROOT = Path(__file__).resolve().parents[1]
INSTALLER_DIR = REPO_ROOT / "packaging" / "windows"
INSTALLER_SCRIPT = INSTALLER_DIR / "installer.iss"

_SECTION = re.compile(r"^\[(?P<name>[A-Za-z]+)\]\s*$")
_ENTRY = re.compile(r"^(?P<key>[A-Za-z0-9_.]+)\s*=\s*(?P<value>.*)$")
_DEFINE = re.compile(r'^\s*#define\s+(?P<name>\w+)\s+"(?P<value>[^"]*)"\s*$')
_INLINE_DEFINE = re.compile(r"\{#(?P<name>\w+)\}")


def _script_text() -> str:
    return INSTALLER_SCRIPT.read_text(encoding="ascii")


def _defines(text: str) -> dict[str, str]:
    return {
        match["name"]: match["value"]
        for match in map(_DEFINE.match, text.splitlines())
        if match
    }


def _section_entries(text: str, section: str) -> list[tuple[str, str]]:
    """Every ``key=value`` line of *section*, preprocessor defines expanded."""

    defines = _defines(text)

    def expand(value: str) -> str:
        def replace(match: re.Match[str]) -> str:
            name = match["name"]
            assert name in defines, f"undefined preprocessor name {name!r}"
            return defines[name]

        return _INLINE_DEFINE.sub(replace, value)

    entries: list[tuple[str, str]] = []
    current = None
    for raw in text.splitlines():
        line = raw.strip()
        if not line or line.startswith(";"):
            continue
        header = _SECTION.match(line)
        if header:
            current = header["name"].lower()
            continue
        if current != section.lower():
            continue
        entry = _ENTRY.match(line)
        if entry:
            entries.append((entry["key"], expand(entry["value"].strip())))
    return entries


def _single_value(text: str, section: str, key: str) -> str:
    values = [
        value
        for name, value in _section_entries(text, section)
        if name.lower() == key.lower()
    ]
    assert len(values) == 1, f"expected one [{section}] {key}, found {len(values)}"
    return values[0]


def test_installer_shows_this_repositorys_license():
    license_file = _single_value(_script_text(), "Setup", "LicenseFile")
    # Relative LicenseFile paths resolve against the script's directory (Inno's
    # default SourceDir), and the script writes them Windows-style.
    resolved = (INSTALLER_DIR / Path(*PureWindowsPath(license_file).parts)).resolve()
    assert resolved == (REPO_ROOT / "LICENSE").resolve()


def test_license_file_is_the_polyform_noncommercial_text_inno_can_display():
    raw = (REPO_ROOT / "LICENSE").read_bytes()
    # Inno treats a file starting with {\rtf as rich text and reads any other
    # BOM-less file in the Windows ANSI code page; ASCII reads the same in all
    # of them, so a curly quote or dash can never turn into mojibake.
    assert not raw.startswith(b"{\\rtf")
    text = raw.decode("ascii")
    first_line = text.splitlines()[0]
    assert first_line.startswith("Required Notice: Copyright 2025 Abraham Borg")
    assert "# PolyForm Noncommercial License 1.0.0" in text
    assert "## Acceptance" in text
    assert "## Noncommercial Purposes" in text
    assert "## No Liability" in text


def test_license_page_describes_the_license_in_plain_words():
    label = _single_value(_script_text(), "Messages", "LicenseLabel3")
    assert label.isascii()
    # "%" introduces Inno message placeholders (%n, %1); the label uses none.
    assert "%" not in label
    lowered = label.lower()
    assert "Specification Formatter" in label
    assert "PolyForm Noncommercial License 1.0.0" in label
    assert "noncommercial purpose" in lowered
    assert "commercial use" in lowered
    assert "separate license" in lowered
    assert "without any warranty" in lowered
    assert "accept" in lowered


def test_nothing_skips_the_license_page_or_accepts_it_for_the_user():
    text = _script_text()
    # Comments may name the page; only script lines and [Code] can act on it.
    # ";" starts a comment in the sections and "//" in Pascal [Code].
    code = "\n".join(
        line
        for line in text.splitlines()
        if not line.lstrip().startswith((";", "//"))
    )
    # wpLicense is the page's ID for ShouldSkipPage and friends, and the two
    # radios are the only way code could preselect "I accept".
    for forbidden in ("wpLicense", "LicenseAcceptedRadio", "LicenseNotAcceptedRadio"):
        assert forbidden.lower() not in code.lower(), forbidden
    messages = {name.lower() for name, _ in _section_entries(text, "Messages")}
    # The radio captions keep Inno's defaults, so "accept" always reads as such.
    assert not messages & {"licenseaccepted", "licensenotaccepted"}


def test_updater_launches_the_installer_interactively_on_windows(tmp_path, monkeypatch):
    installer = tmp_path / "SpecificationFormatterSetup.exe"
    calls = []
    monkeypatch.setattr(sys, "platform", "win32")
    monkeypatch.setattr(
        updates.os,
        "startfile",
        lambda *args, **kwargs: calls.append((args, kwargs)),
        raising=False,
    )

    updates.spawn_installer(installer)

    # No /SILENT or /VERYSILENT: the wizard, license page included, is shown.
    assert calls == [((str(installer),), {})]


def test_updater_launches_the_installer_interactively_elsewhere(tmp_path, monkeypatch):
    installer = tmp_path / "SpecificationFormatterSetup.exe"
    calls = []
    monkeypatch.setattr(sys, "platform", "linux")
    monkeypatch.setattr(
        updates.subprocess,
        "Popen",
        lambda *args, **kwargs: calls.append((args, kwargs)),
    )

    updates.spawn_installer(installer)

    assert calls == [(([str(installer)],), {})]
