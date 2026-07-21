"""App-specific filesystem locations for persistent state and config.

Mirrors the platform strategy in
:func:`spec_formatter.pipeline.default_template_cache_dir`, but targets a
*config/state* directory (persisted, not a disposable cache): the updater's
once-a-day throttle marker, the "skip this version" flag, and downloaded
installers all live here.

    Windows -> ``%LOCALAPPDATA%\\SpecificationFormatter``
    POSIX   -> ``$XDG_CONFIG_HOME/specification-formatter`` or
               ``~/.config/specification-formatter``

The path is computed lazily and is NOT created here; callers that write (the
updater's ``save_state`` / ``download_installer``) create it on demand, so
simply resolving the path never touches the filesystem.
"""

from __future__ import annotations

import os
from pathlib import Path

# Windows uses the product PascalCase name (matching the existing
# ``%LOCALAPPDATA%\\SpecificationFormatter\\TemplateCache`` cache dir); POSIX
# uses the kebab-case slug (matching ``specification-formatter/template-cache``).
APP_DIR_NAME = "SpecificationFormatter"
POSIX_DIR_NAME = "specification-formatter"


def default_config_dir() -> Path:
    """Return the per-user config/state directory (not created here)."""

    local_app_data = os.environ.get("LOCALAPPDATA")
    if local_app_data:
        return Path(local_app_data) / APP_DIR_NAME
    xdg_config = os.environ.get("XDG_CONFIG_HOME")
    if xdg_config:
        return Path(xdg_config) / POSIX_DIR_NAME
    return Path.home() / ".config" / POSIX_DIR_NAME
