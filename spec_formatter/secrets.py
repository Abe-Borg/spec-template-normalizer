"""Optional persistence of the user's Anthropic API key via the OS keyring.

On Windows the key is stored in Credential Manager (keyring's Windows backend).
Everything degrades gracefully: if no keyring backend is available, ``keyring``
is not installed, or any backend call fails, the helpers behave as though
nothing is stored. The app then still works with an in-memory / ``ANTHROPIC_API_KEY``
env-var key -- persistence is a convenience, never a requirement.

``keyring`` is imported lazily inside each function (mirroring the deferred
``import anthropic`` style elsewhere in the codebase) so importing this module
stays cheap and never fails when the backend is missing.
"""

from __future__ import annotations

KEYRING_SERVICE = "SpecificationFormatter"
KEYRING_USERNAME = "anthropic-api-key"


def load_api_key() -> str:
    """Return the stored API key, or ``""`` when none is stored / no backend."""

    try:
        import keyring

        value = keyring.get_password(KEYRING_SERVICE, KEYRING_USERNAME)
    except Exception:
        return ""
    return value or ""


def save_api_key(key: str) -> bool:
    """Persist ``key`` to the OS keyring. Returns ``True`` on success.

    An empty/blank key clears any stored value instead of saving it.
    """

    key = (key or "").strip()
    if not key:
        return clear_api_key()
    try:
        import keyring

        keyring.set_password(KEYRING_SERVICE, KEYRING_USERNAME, key)
        return True
    except Exception:
        return False


def clear_api_key() -> bool:
    """Delete any stored API key. Returns ``True`` if it is now absent."""

    try:
        import keyring

        try:
            keyring.delete_password(KEYRING_SERVICE, KEYRING_USERNAME)
        except Exception:
            # Already absent, or the backend has nothing to delete.
            pass
        return True
    except Exception:
        return False
