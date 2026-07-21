"""Unified architect-template specification formatter.

The public façade is loaded lazily so the independently namespaced style
application engine can be imported without importing the template engine.
"""

from typing import Any

# Single source of truth for the app version. The frozen Windows build reports
# this value, the in-app updater compares it against the release manifest, and
# packaging/windows/check_release_version.py guards it against the git tag.
# Bump this (only) when cutting a release; see docs/RELEASE_WINDOWS.md.
__version__ = "1.0.0"

__all__ = [
    "CSI_TO_CANADIAN",
    "FORMAT_ONLY",
    "FormatRunResult",
    "TargetFormatResult",
    "TemplateProfile",
    "default_template_cache_dir",
    "format_specifications",
]


def __getattr__(name: str) -> Any:
    if name in __all__:
        from . import pipeline

        return getattr(pipeline, name)
    raise AttributeError(name)
