"""Who made Specification Formatter and how to reach them.

The About dialog in ``gui.py`` reads these values, and
``packaging/windows/installer.iss`` repeats them as preprocessor defines
(the installer is compiled by Inno Setup, which cannot import Python).
``tests/test_installer_license.py`` holds the two copies equal.
"""

from __future__ import annotations

AUTHOR_NAME = "Abraham Borg"
AUTHOR_EMAIL = "abrahamborg@gmail.com"
AUTHOR_LINKEDIN_URL = "https://www.linkedin.com/in/abrahamborg"

COPYRIGHT_NOTICE = f"Copyright 2025 {AUTHOR_NAME}"


def author_mailto_url() -> str:
    """The ``mailto:`` link the About dialog opens for :data:`AUTHOR_EMAIL`."""

    return f"mailto:{AUTHOR_EMAIL}"


def linkedin_display_text() -> str:
    """:data:`AUTHOR_LINKEDIN_URL` without its scheme and ``www.``, for display."""

    return AUTHOR_LINKEDIN_URL.removeprefix("https://").removeprefix("www.")
