"""Unified architect-template specification formatter.

The public façade is loaded lazily so the independently namespaced style
application engine can be imported without importing the template engine.
"""

from typing import Any

__all__ = [
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
