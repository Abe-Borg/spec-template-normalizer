"""Generate THIRD_PARTY_NOTICES.md from a resolved dependency set.

Attribution licences (MIT, BSD, Apache-2.0, PSF, MPL) require the copyright
notice and licence text to travel with a *binary* distribution. Hand-maintaining
that list does not work: ``requirements.txt`` pins only the direct dependencies,
so the transitive packages pip actually installs -- and PyInstaller therefore
freezes into the app -- silently go unrecorded.

This script resolves the real runtime closure from the ``requirements.txt``
roots, evaluating environment markers so platform-conditional dependencies are
included or excluded correctly, then extracts each distribution's *own* licence
file rather than a generic licence template.

Two sources are supported:

* installed distributions (default) -- what the Windows release job has after
  ``pip install -r requirements.txt``. This is the authoritative mode: it
  describes exactly what is about to be frozen.
* ``--wheel-dir`` -- a directory of downloaded wheels, which allows the notices
  for a target platform to be regenerated from any machine::

      pip download -r requirements.txt --only-binary :all: \\
          --platform win_amd64 --python-version 3.11 -d wheels
      python packaging/windows/generate_third_party_notices.py \\
          --wheel-dir wheels --target-platform win32

Build-time-only tooling (PyInstaller and friends) is deliberately not part of
the closure: it is not shipped. The one exception is PyInstaller's bootloader,
which *is* embedded in the executable and is covered by the exception
reproduced in the static preamble below.
"""

from __future__ import annotations

import argparse
import glob
import io
import os
import re
import sys
import zipfile
from typing import Dict, List, Optional, Sequence, Tuple

REPO_ROOT = os.path.abspath(os.path.join(os.path.dirname(__file__), "..", ".."))

# dist-info members that carry a licence, in the layouts wheels actually use:
# PEP 639 ``licenses/`` subdirectories and the older flat files.
_LICENSE_MEMBER = re.compile(
    r"\.dist-info/(licenses/.+|(LICENSE|LICENCE|COPYING|NOTICE|AUTHORS)[^/]*)$",
    re.IGNORECASE,
)

# Licence text is reproduced verbatim, but a handful of projects ship enormous
# AUTHORS files; cap any single file so one dependency cannot swamp the notices.
_MAX_LICENSE_CHARS = 20000


def _canonical(name: str) -> str:
    return re.sub(r"[-_.]+", "-", name).strip().lower()


class Package:
    """One distribution's identity and verbatim licence text."""

    def __init__(self, name: str, version: str, license_id: str) -> None:
        self.name = name
        self.version = version
        self.license_id = license_id
        self.license_files: List[Tuple[str, str]] = []

    @property
    def key(self) -> str:
        return _canonical(self.name)


def _license_id(metadata: str) -> str:
    """Best available SPDX-ish identifier from a METADATA blob."""
    expr = re.search(r"^License-Expression:\s*(.+)$", metadata, re.M)
    if expr:
        return expr.group(1).strip()

    classifiers = re.findall(r"^Classifier:\s*License\s*::\s*(.+)$", metadata, re.M)
    if classifiers:
        # "OSI Approved :: MIT License" -> "MIT License"
        return "; ".join(c.split("::")[-1].strip() for c in classifiers)

    legacy = re.search(r"^License:\s*(.+)$", metadata, re.M)
    if legacy:
        value = legacy.group(1).strip()
        # Some projects dump their whole licence into this field.
        if value and len(value) < 60 and "\n" not in value:
            return value
    return "see licence text below"


def _requirements(metadata: str) -> List[str]:
    return re.findall(r"^Requires-Dist:\s*(.+)$", metadata, re.M)


def _marker_env(target_platform: Optional[str], python_version: str) -> Dict[str, str]:
    from packaging.markers import default_environment

    env = dict(default_environment())
    env["python_version"] = python_version
    if target_platform == "win32":
        env.update(
            {
                "sys_platform": "win32",
                "platform_system": "Windows",
                "os_name": "nt",
                "platform_machine": "AMD64",
                "implementation_name": "cpython",
                "platform_python_implementation": "CPython",
            }
        )
    return env


def _read_roots(requirements_path: str) -> List[str]:
    lines: List[str] = []
    with io.open(requirements_path, encoding="utf-8") as handle:
        for raw in handle:
            line = raw.split("#", 1)[0].strip()
            if line and not line.startswith("-"):
                lines.append(line)
    return lines


def _collect_from_wheels(wheel_dir: str) -> Dict[str, Tuple[Package, List[str]]]:
    found: Dict[str, Tuple[Package, List[str]]] = {}
    for wheel in sorted(glob.glob(os.path.join(wheel_dir, "*.whl"))):
        with zipfile.ZipFile(wheel) as archive:
            names = archive.namelist()
            meta_names = [n for n in names if n.endswith(".dist-info/METADATA")]
            if not meta_names:
                continue
            metadata = archive.read(meta_names[0]).decode("utf-8", "replace")
            name = re.search(r"^Name:\s*(.+)$", metadata, re.M).group(1).strip()
            version = re.search(r"^Version:\s*(.+)$", metadata, re.M).group(1).strip()
            package = Package(name, version, _license_id(metadata))
            for member in sorted(names):
                if _LICENSE_MEMBER.search(member):
                    text = archive.read(member).decode("utf-8", "replace")
                    package.license_files.append(
                        (os.path.basename(member), text[:_MAX_LICENSE_CHARS])
                    )
            found[package.key] = (package, _requirements(metadata))
    return found


def _collect_from_installed() -> Dict[str, Tuple[Package, List[str]]]:
    import importlib.metadata as importlib_metadata

    found: Dict[str, Tuple[Package, List[str]]] = {}
    failures: List[str] = []

    for dist in importlib_metadata.distributions():
        name = ""
        try:
            # Deliberately NOT ``str(dist.metadata)``: that re-serializes the
            # headers through the email package, which raises for any project
            # whose METADATA carries a long multi-line ``Description`` header
            # (certifi and anthropic among them). Reading the raw part keeps
            # every distribution in the closure.
            metadata = dist.read_text("METADATA") or dist.read_text("PKG-INFO") or ""
            name = dist.metadata["Name"] or ""
            if not name or not metadata:
                continue
            package = Package(name, dist.version or "", _license_id(metadata))
            for entry in dist.files or []:
                member = str(entry).replace(os.sep, "/")
                if not _LICENSE_MEMBER.search(member):
                    continue
                text = None
                try:
                    with io.open(
                        str(dist.locate_file(entry)),
                        encoding="utf-8",
                        errors="replace",
                    ) as handle:
                        text = handle.read()
                except Exception:
                    try:
                        text = dist.read_text(member.split("/", 1)[-1])
                    except Exception:
                        text = None
                if text is None:
                    continue
                package.license_files.append(
                    (os.path.basename(member), text[:_MAX_LICENSE_CHARS])
                )
            found[package.key] = (package, _requirements(metadata))
        except Exception as exc:  # pragma: no cover - defensive
            failures.append("%s (%s)" % (name or "<unknown>", exc.__class__.__name__))

    if failures:
        # Never silent: a dropped distribution is undisclosed shipped code.
        print(
            "warning: could not read metadata for: %s" % ", ".join(failures),
            file=sys.stderr,
        )
    return found


def _resolve_closure(
    available: Dict[str, Tuple[Package, List[str]]],
    roots: Sequence[str],
    env: Dict[str, str],
) -> Tuple[List[Package], List[str]]:
    from packaging.requirements import Requirement

    keep: Dict[str, Package] = {}
    missing: List[str] = []
    stack: List[str] = []

    for raw in roots:
        try:
            requirement = Requirement(raw)
        except Exception:
            continue
        if requirement.marker is None or requirement.marker.evaluate(env):
            stack.append(_canonical(requirement.name))

    while stack:
        key = stack.pop()
        if key in keep:
            continue
        entry = available.get(key)
        if entry is None:
            if key not in missing:
                missing.append(key)
            continue
        package, requires = entry
        keep[key] = package
        for raw in requires:
            try:
                requirement = Requirement(raw)
            except Exception:
                continue
            marker = requirement.marker
            if marker is not None:
                # ``extra`` markers guard optional features that pip did not
                # install, so they are not part of the shipped closure.
                if "extra" in str(marker) or not marker.evaluate(env):
                    continue
            stack.append(_canonical(requirement.name))

    return [keep[k] for k in sorted(keep)], sorted(missing)


PREAMBLE = """# Third-Party Notices

<!-- GENERATED FILE -- do not edit by hand.
     Regenerate with packaging/windows/generate_third_party_notices.py; the
     Windows release workflow regenerates it from the real build environment
     before PyInstaller runs, so the shipped copy always matches the shipped
     code. -->

Specification Formatter is distributed under the PolyForm Noncommercial License
1.0.0 (see `LICENSE`). That license covers this project's own source code only.

The application, and the Windows installer built from it, also include
third-party components licensed by their respective copyright holders under the
terms reproduced below. Those terms govern those components, not this project's
license.

None of these components imposes a copyleft obligation on this project's own
source code, so this project is free to be licensed as it is. The obligations
that do apply are listed under "Obligations".

## Obligations

**Attribution (MIT, BSD, Apache-2.0, PSF).** The copyright notice and license
text must accompany binary distributions. Each package's own license file is
reproduced verbatim below, so the actual copyright lines -- not a generic
template -- ship with the application.

**certifi (MPL-2.0).** The Mozilla Public License 2.0 is a file-level copyleft.
It reaches only certifi's own files, never this project's source. certifi is
bundled unmodified, so the obligation is to preserve its notice and to make its
source available; the unmodified source is published at
<https://pypi.org/project/certifi/>. Modifying certifi's files would require
releasing those modified files under MPL-2.0.

**Apache-2.0 components.** If one is ever modified, the modified files must
carry prominent change notices, and any `NOTICE` file the component ships must
be reproduced.

**PyInstaller (GPL-2.0-or-later with the Bootloader Exception).** PyInstaller is
a build tool and is not itself shipped, but the bootloader it embeds into the
frozen executable is. That bootloader is covered by an explicit exception
permitting distribution of the combined executable under any terms:

```
In addition to the permissions in the GNU General Public License, the
authors give you unlimited permission to link or embed compiled bootloader
and related files into combinations with other programs, and to distribute
those combinations without any restriction coming from the use of those
files. (The General Public License restrictions do apply in other respects;
for example, they cover modification of the files, and distribution when
not linked into a combined executable.)
```

The GNU General Public License text is not reproduced here because the exception
removes the GPL's restrictions from the shipped combination; it continues to
govern modification and redistribution of PyInstaller itself, which this project
does not do.

The application also bundles a CPython interpreter and the Tcl/Tk libraries
backing Tkinter. CPython is distributed under the Python Software Foundation
License; Tcl/Tk under a BSD-style license. The Windows installer is produced
with Inno Setup, whose license permits building installers for commercial and
noncommercial applications.
"""


def render(packages: Sequence[Package], platform_label: str) -> str:
    out: List[str] = [PREAMBLE]
    out.append("## Bundled components\n")
    out.append(
        "The %d packages below are the complete runtime closure of "
        "`requirements.txt` for %s, including transitive dependencies.\n"
        % (len(packages), platform_label)
    )
    out.append("| Package | Version | License |")
    out.append("|---|---|---|")
    for package in packages:
        out.append(
            "| `%s` | %s | %s |" % (package.name, package.version, package.license_id)
        )
    out.append("\n## License texts\n")
    out.append(
        "Reproduced verbatim from each distribution's own packaged license "
        "file.\n"
    )
    for package in packages:
        out.append("### %s %s\n" % (package.name, package.version))
        if not package.license_files:
            out.append(
                "_This distribution ships no license file. Declared license: "
                "%s._\n" % package.license_id
            )
            continue
        for filename, text in package.license_files:
            out.append("<details><summary><code>%s</code></summary>\n" % filename)
            out.append("```")
            out.append(text.strip())
            out.append("```")
            out.append("</details>\n")
    return "\n".join(out) + "\n"


def main(argv: Optional[Sequence[str]] = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument(
        "--wheel-dir",
        help="Read distributions from downloaded wheels instead of the "
        "installed environment.",
    )
    parser.add_argument(
        "--requirements",
        default=os.path.join(REPO_ROOT, "requirements.txt"),
        help="Requirements file whose entries seed the closure.",
    )
    parser.add_argument(
        "--target-platform",
        choices=["win32", "current"],
        default="current",
        help="Platform to evaluate environment markers for.",
    )
    parser.add_argument(
        "--python-version",
        default="%d.%d" % sys.version_info[:2],
        help="Python version to evaluate environment markers for.",
    )
    parser.add_argument(
        "--output",
        default=os.path.join(REPO_ROOT, "THIRD_PARTY_NOTICES.md"),
        help="File to write.",
    )
    args = parser.parse_args(argv)

    if args.wheel_dir:
        available = _collect_from_wheels(args.wheel_dir)
    else:
        available = _collect_from_installed()
    if not available:
        print("error: no distributions found", file=sys.stderr)
        return 1

    env = _marker_env(
        None if args.target_platform == "current" else args.target_platform,
        args.python_version,
    )
    packages, missing = _resolve_closure(
        available, _read_roots(args.requirements), env
    )
    if missing:
        print(
            "error: no distribution found for: %s\n"
            "The closure is incomplete, which would ship undisclosed code."
            % ", ".join(missing),
            file=sys.stderr,
        )
        return 1

    label = (
        "Windows (CPython %s)" % args.python_version
        if args.target_platform == "win32"
        else "this platform (CPython %s)" % args.python_version
    )
    with io.open(args.output, "w", encoding="utf-8", newline="\n") as handle:
        handle.write(render(packages, label))

    without = [p.name for p in packages if not p.license_files]
    print("wrote %s (%d packages)" % (args.output, len(packages)))
    if without:
        print("warning: no packaged license file for: %s" % ", ".join(without))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
