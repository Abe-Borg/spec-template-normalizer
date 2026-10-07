"""Target classifier defaults shared by the public API and desktop app."""

HAIKU_TARGET_MODEL = "claude-haiku-5-5"
DEFAULT_TARGET_MODEL = HAIKU_TARGET_MODEL

TARGET_MODEL_CHOICES: tuple[tuple[str, str], ...] = (
    ("Haiku 5.5", HAIKU_TARGET_MODEL),
    ("Sonnet 5.5", "claude-sonnet-5-5"),
)
