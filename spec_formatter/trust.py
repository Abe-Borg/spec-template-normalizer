"""The "Why Trust It?" topic and the "I'm not convinced" dossier, as data.

Purpose: the two trust surfaces of the desktop app. Level 1 is the short
**Why Trust It?** topic behind the help button of that name. Level 2 is the
long dossier its closing button opens, which names the mechanism behind every
claim, action by action. ``trust_ui.py`` draws both in the window, and
``render_markdown()`` produces ``docs/TRUST.md`` from the same text.

Reader: a working engineer who puts their name on the formatted
specification, plus whoever approves installing the app. The text uses plain
words and defines a technical term the first time it appears.

If the implementation changes, this changes with it. A trust document that
has drifted from the code is worse than none.

Every number, model ID, limit, host name, path and setting the copy quotes
lives in :data:`FACTS` as a hand-written literal. ``tests/test_trust_surface.py``
asserts each one equals the constant or literal in the code that makes it true,
and fails if a digit appears in the copy that did not come from here. Change
the code and that test fails until the copy is updated. ``docs/TRUST_CLAIMS.md``
is the ledger behind every claim. ``python -m spec_formatter.trust`` prints
the Markdown rendering; ``docs/TRUST.md`` must equal it.

This module imports nothing from the engine, so it costs nothing to load.
"""

from __future__ import annotations

from dataclasses import dataclass
from typing import Iterator, Union

# ---------------------------------------------------------------------------
# Pinned facts. Hand-written on purpose: a value read from the code would
# follow a change silently, and the copy around it would not.
# ---------------------------------------------------------------------------

FACTS: dict[str, object] = {
    # Network
    "anthropic_host": "api.anthropic.com",
    "anthropic_base_url_env": "ANTHROPIC_BASE_URL",
    "api_key_env": "ANTHROPIC_API_KEY",
    "update_host": "github.com",
    "update_manifest_url": (
        "https://github.com/abe-borg/spec-template-normalizer/releases/latest/download/latest.json"
    ),
    "releases_url": "https://github.com/abe-borg/spec-template-normalizer/releases/latest",
    "update_user_agent": "SpecificationFormatter-Updater",
    "update_check_delay_s": 1.5,
    "update_check_interval_days": 1,
    "update_disable_env": "SPEC_FORMATTER_DISABLE_UPDATE_CHECK",
    "update_url_env": "SPEC_FORMATTER_UPDATE_URL",
    "update_manifest_timeout_s": 8,
    "update_manifest_max_kib": 64,
    "installer_timeout_s": 60,
    "installer_max_mib": 512,
    "console_keys_url": "https://console.anthropic.com/settings/keys",
    # Models and request settings
    "architect_model": "claude-opus-5-5",
    "architect_model_label": "Opus 5.5",
    "target_model_default": "claude-haiku-5-5",
    "target_model_default_label": "Haiku 5.5",
    "target_model_alt": "claude-sonnet-5-5",
    "target_model_alt_label": "Sonnet 5.5",
    "effort": "high",
    "thinking_type": "adaptive",
    "max_output_tokens": 128_000,
    "target_effort_env": "SPEC_FORMATTER_TARGET_EFFORT",
    "target_ignore_reason": "non_csi_content",
    "connect_timeout_s": 5,
    "read_timeout_s": 600,
    "read_timeout_min": 10,
    "target_transport_retries": 2,
    "retry_after_cap_s": 120,
    "backoff_first_s": 2,
    "backoff_second_s": 4,
    "target_regenerations": 2,
    "target_total_attempts": 3,
    "architect_transport_attempts": 3,
    "architect_response_attempts": 2,
    "architect_patch_attempts": 3,
    "architect_requests": 4,
    "architect_sends_per_request": 6,
    "chunk_tokens": 80_000,
    "chunk_paragraphs": 300,
    "chunk_overlap": 20,
    "chunk_workers": 6,
    "max_concurrent_requests": 4,
    "max_concurrent_requests_ceiling": 64,
    "max_concurrent_requests_env": "SPEC_FORMATTER_MAX_CONCURRENT_REQUESTS",
    # What is sent
    "architect_text_chars": 200,
    "target_text_chars": 200,
    "neighbour_text_chars": 80,
    # Cost and size guards
    "architect_input_cap": 150_000,
    "target_paragraph_cap": 2_000,
    "target_paragraph_cap_env": "SPEC_FORMATTER_MAX_TARGET_PARAGRAPHS",
    "haiku_prompt_budget": 95_000,
    "haiku_price_step": 100_000,
    "package_entries": 10_000,
    "package_total_mib": 512,
    "package_part_mib": 128,
    "compression_ratio": 1_000,
    "hf_media_mib": 16,
    "hf_media_total_mib": 64,
    "gui_workers": 3,
    "api_max_workers": 6,
    # Data at rest
    "cache_dir_windows": r"%LOCALAPPDATA%\SpecificationFormatter\TemplateCache",
    "config_dir_windows": r"%LOCALAPPDATA%\SpecificationFormatter",
    "updates_subdir": "updates",
    "update_state_file": "update_check.json",
    "cache_keep": 2,
    "keyring_service": "SpecificationFormatter",
    "keyring_user": "anthropic-api-key",
}

#: Numbers that may appear in the copy without being a fact, each with the
#: reason. Anything else with a digit in it must come from :data:`FACTS`.
NAMED_NUMBERS: dict[str, str] = {
    "256": "the name of the SHA-256 hash algorithm",
    "1": "the CSI marker examples PART 1 and 1. and the value 1 in an example",
    "1.1": "the Canadian/CSI marker example 1.1",
}


def _n(key: str) -> str:
    """A pinned fact as the copy prints it (integers with thousands commas)."""

    value = FACTS[key]
    if isinstance(value, bool):  # pragma: no cover - no boolean facts
        raise TypeError(key)
    if isinstance(value, int):
        return f"{value:,}"
    return str(value)


# ---------------------------------------------------------------------------
# Building blocks. ``trust_ui.py`` has one drawing primitive per block type,
# and ``render_markdown`` one rendering rule, so every section looks alike.
# ---------------------------------------------------------------------------


@dataclass(frozen=True)
class Paragraph:
    text: str


@dataclass(frozen=True)
class Bullets:
    items: tuple[str, ...]
    numbered: bool = False


@dataclass(frozen=True)
class KeyPoints:
    """Bold lead phrase plus explanation, one per item."""

    items: tuple[tuple[str, str], ...]


@dataclass(frozen=True)
class Table:
    columns: tuple[str, ...]
    rows: tuple[tuple[str, ...], ...]


@dataclass(frozen=True)
class Contrast:
    """Two columns side by side (one above the other when narrow)."""

    left_title: str
    left_items: tuple[str, ...]
    right_title: str
    right_items: tuple[str, ...]


@dataclass(frozen=True)
class Note:
    title: str
    text: str


@dataclass(frozen=True)
class ExactClaim:
    term: str
    claim: str
    not_claimed: str


@dataclass(frozen=True)
class Links:
    title: str
    items: tuple[tuple[str, str], ...]


@dataclass(frozen=True)
class DataFlowDiagram:
    """Drawn by the UI; ``description`` restates it in words for everyone."""

    description: str
    machine_label: str
    machine_items: tuple[str, ...]
    destinations: tuple[tuple[str, str, bool], ...]  # (label, arrow text, dashed)
    legend: str


#: The five rows of every runtime card, in order. The identical rows are the
#: point: they let a reader compare cards instead of reading essays.
RUNTIME_CARD_ROWS: tuple[str, ...] = (
    "You do",
    "What runs",
    "What is sent",
    "AI involved",
    "Bounded by",
)


@dataclass(frozen=True)
class RuntimeCard:
    title: str
    you_do: str
    what_runs: str
    what_is_sent: str
    ai_involved: str
    bounded_by: str
    automatic: bool = False

    def rows(self) -> tuple[tuple[str, str], ...]:
        return tuple(
            zip(
                RUNTIME_CARD_ROWS,
                (
                    self.you_do,
                    self.what_runs,
                    self.what_is_sent,
                    self.ai_involved,
                    self.bounded_by,
                ),
            )
        )


@dataclass(frozen=True)
class RuntimeCards:
    cards: tuple[RuntimeCard, ...]


Block = Union[
    Paragraph, Bullets, KeyPoints, Table, Contrast, Note, ExactClaim, Links,
    DataFlowDiagram, RuntimeCards,
]


@dataclass(frozen=True)
class Section:
    anchor: str
    kicker: str
    title: str
    blocks: tuple[Block, ...]
    #: The short label in the dossier's contents rail.
    rail: str = ""

    @property
    def rail_label(self) -> str:
        return self.rail or self.title


@dataclass(frozen=True)
class TrustPoint:
    title: str
    body: str


# ---------------------------------------------------------------------------
# Level 1: "Why Trust It?"
# ---------------------------------------------------------------------------

LEVEL1_TITLE = "Why Trust It?"

LEVEL1_LEAD = (
    "The whole design assumes a qualified person checks every output, so the "
    "app never edits your originals, never quietly guesses, and withholds any "
    "file it cannot verify."
)

LEVEL1_POINTS: tuple[TrustPoint, ...] = (
    TrustPoint(
        "Your original files are never edited",
        "The app copies your template and each target to a private snapshot, "
        "confirms the copy matches the original by its SHA-256 fingerprint, "
        "and works only on the copy. Each original is checked again before "
        "its output is published; if it changed during the run, nothing is "
        "published for it.",
    ),
    TrustPoint(
        "Rules decide first, AI only the rest",
        "A paragraph whose role is proven by Word's own numbering or by "
        "unmistakable text such as PART 1 - GENERAL is decided by fixed rules "
        "on your computer, and the AI cannot overrule them. Only the "
        "remaining paragraphs are sent. When rules settle every paragraph of "
        "a target, no request is made for it.",
    ),
    TrustPoint(
        "The AI picks labels, never your words",
        "For each paragraph it is shown, the AI may only answer which CSI role "
        "it has (PART, ARTICLE, PARAGRAPH and so on) or that it is not CSI "
        "content. Formatting is copied from the template's own paragraphs, "
        "and no model writes the text of your specification.",
    ),
    TrustPoint(
        "Nothing is silently dropped or filled in",
        "Every visible body paragraph outside tables, text boxes and drawings "
        "must end up styled or explicitly ignored, exactly once. A gap, a duplicate, or an answer outside the allowed "
        "roles fails that target; it is never patched from the paragraph "
        "next door.",
    ),
    TrustPoint(
        "A file that fails a check is withheld",
        "Before publishing, the app proves the body text is unchanged except "
        "where the mode said it would change, and that numbering, tracked "
        "changes, indentation and every untouched part of the file survived. "
        "A failed check withholds that one output; your other targets still "
        "finish.",
    ),
    TrustPoint(
        "One thing runs without you, and it is named",
        f"About {_n('update_check_delay_s')} seconds after launch, at most "
        f"once a day, the app asks {_n('update_host')} for a small version "
        "file. It sends no document data, no key and no version number. Set "
        f"{_n('update_disable_env')} to turn it off. Every AI request, "
        "retries included, happens only inside a run you started.",
    ),
    TrustPoint(
        "Your key and your spend stay yours",
        "Requests go straight from your computer to Anthropic "
        f"({_n('anthropic_host')}) with your own key and are billed to your "
        "Anthropic account; no server of this app's sits in between. The key is held "
        "in memory, or in Windows Credential Manager if you tick Remember, "
        "and is removed from every log and record.",
    ),
    TrustPoint(
        "Every run leaves a record you can check",
        "Each run gets its own timestamped folder with run.log, an audit per "
        "target, diagnostics.jsonl and run.json: fingerprints of every input "
        "and output, the models used, and the tokens counted. None of it "
        "holds your paragraph text or your key.",
    ),
)

LEVEL1_CLOSING_HEADING = "Not convinced?"
LEVEL1_CLOSING_TEXT = (
    "Fair. The points above are claims — here is the mechanism behind each "
    "one, action by action, plus how to audit any of it yourself."
)
LEVEL1_BUTTON = "I'm not convinced — show me exactly what runs →"


# ---------------------------------------------------------------------------
# Level 2: the dossier
# ---------------------------------------------------------------------------

DOSSIER_TITLE = "Exactly what runs"
DOSSIER_SUBTITLE = (
    "For the person who puts their name on the output. This describes "
    "Specification Formatter as it ships, with its default settings; where a "
    "setting can change a behaviour, the setting is named."
)

_ANTHROPIC = _n("anthropic_host")
_GITHUB = _n("update_host")
_ARCH = f"{_n('architect_model_label')} ({_n('architect_model')})"
_HAIKU = f"{_n('target_model_default_label')} ({_n('target_model_default')})"
_SONNET = f"{_n('target_model_alt_label')} ({_n('target_model_alt')})"
_SETTINGS = (
    f"adaptive thinking, effort “{_n('effort')}”, at most "
    f"{_n('max_output_tokens')} output tokens, no tools"
)


_SHORT_ANSWER = Section(
    "short-answer",
    "Start here",
    "The short answer",
    (
        Paragraph(
            "Specification Formatter applies an architect's Word template to "
            "your specifications, or converts their numbering between CSI and "
            "Canadian CSC PageFormat. Three commitments hold the rest of this "
            "document together."
        ),
        KeyPoints(
            (
                (
                    "Every part of the output has a declared origin.",
                    "Your words and, in Format only, your numbering come from "
                    "your target. Formatting comes from the template's own "
                    "paragraphs or from the app's built-in Canadian scheme. "
                    "The AI contributes role labels and, for a template, the "
                    "names of the styles it generates. One gap: the records "
                    "say how many paragraphs the AI labelled, not which ones.",
                ),
                (
                    "Nothing runs that you did not start, with one named exception.",
                    "The exception is the daily update check, which sends no "
                    "document data and can be switched off. Template analysis, "
                    "AI labelling and every retry happen only inside a run you "
                    "started with FORMAT SPECS.",
                ),
                (
                    "Every automated judgment is checked; the rule-based ones can be re-derived.",
                    "Rules give the same answer every time for the same file. "
                    "AI answers are not guaranteed to repeat, so each one is "
                    "confined to a fixed list of roles, checked for complete "
                    "coverage, and then held to the same verification gates "
                    "as everything else before a file is published.",
                ),
            )
        ),
    ),
    rail="The short answer",
)


_PROVENANCE = Section(
    "provenance",
    "Provenance",
    "Where the output comes from",
    (
        Table(
            ("Origin", "What it actually is", "How you can tell"),
            (
                (
                    "Your target specification",
                    "All body text apart from the list numbers a conversion "
                    "changes, plus tables, footnotes, comments, images and "
                    "tracked changes. In Format only, also its list numbers, "
                    "levels and restarts.",
                    "Format only proves body text and numbering unchanged "
                    "before publishing. Every part of the file outside the "
                    "mode's remit is checked byte for byte. In Format only, "
                    "Word's Review › Compare shows no text changes in the "
                    "body.",
                ),
                (
                    "The architect template (two template modes)",
                    "Paragraph styles copied from the template's own example "
                    "paragraphs, document defaults, theme, page layout, and "
                    "headers and footers. In Canadian mode, also its numbering.",
                    "run.json records the template's SHA-256 fingerprint. "
                    "run.log lists every part of the file that changed, was "
                    "added or was removed.",
                ),
                (
                    "The built-in Canadian scheme (no-template conversion)",
                    "A nine-level PART / 1.1 / .1 numbering list generated "
                    "from constants inside the app and checked against the "
                    "same rules a template must pass. No styles, fonts or "
                    "headers are applied.",
                    "run.json says numbering_scheme: builtin_csc and records "
                    "the scheme's fingerprint.",
                ),
                (
                    "Fixed rules in the app",
                    "Most role decisions, both Canadian conversions, the "
                    "numbers written back as typed CSI markers, and the "
                    "section number and title placed in imported headers.",
                    "run.log: “Built slim bundle: N unresolved + M "
                    "deterministic”. Deterministic means decided by rules.",
                ),
                (
                    "AI, through Anthropic",
                    "A role label for each paragraph the rules could not "
                    "settle. For a newly analyzed template, which paragraph "
                    "is the example of each role, and the display names of "
                    "the CSI_ styles it creates; those names appear in Word's "
                    "Styles pane in the two template modes.",
                    "run.log says “Classifying unresolved paragraphs with "
                    "Anthropic” or “Anthropic request skipped”. run.json "
                    "records the models, effort and tokens. The audit lists "
                    "every paragraph's final role, but not whether a rule or "
                    "the AI chose it.",
                ),
            ),
        ),
        Note(
            "Things that are not happening",
            "There is no hidden library or corpus: the only reference material "
            "shipped is the classifiers' prompt files and the built-in "
            "numbering scheme, and nothing is fetched from the internet to "
            "write your output. There is no server, account or sign-in "
            "belonging to this app, so nothing is shared between users. The "
            "app does no training of its own; Anthropic's Commercial Terms "
            "state that “Anthropic may not train models on Customer Content "
            "from Services.” There is no telemetry, crash reporting or "
            "analytics. The AI never sees a whole document, an image, a "
            "header or footer, or a file name.",
        ),
    ),
    rail="Where it comes from",
)


_ENGINE = Section(
    "engine",
    "Engine",
    "What the engine actually is",
    (
        Table(
            ("Job", "Model or component", "Why"),
            (
                (
                    "Analyze a template the app has not seen before",
                    f"Anthropic {_ARCH}, on Anthropic's servers: {_SETTINGS}.",
                    "The most capable model, for a job done once per template "
                    "version; the result is cached and reused for free.",
                ),
                (
                    "Label target paragraphs that rules could not settle",
                    f"Anthropic {_HAIKU} by default, or {_SONNET} from the "
                    f"Target AI model menu, on Anthropic's servers: {_SETTINGS}.",
                    "Haiku is the lower-cost default; Sonnet is there for "
                    "documents where you want the larger model.",
                ),
                (
                    "Measure a template's first request, and each Haiku "
                    "request, before sending it",
                    "Anthropic's token-counting service, on Anthropic's servers.",
                    "Enforces the cost limits described under Money. It "
                    "produces no answer and is not billed as model use.",
                ),
                (
                    "Everything else: copying, unpacking, rules, conversion, "
                    "styling, checking, packaging, publishing",
                    "This app's own code, on your computer.",
                    "Deterministic: the same input gives the same result, with "
                    "no AI and no network.",
                ),
                (
                    "Remember your key (optional)",
                    "Windows Credential Manager, on your computer.",
                    "Only when you tick Remember.",
                ),
                (
                    "Check for updates",
                    f"Python's standard web client, talking to {_GITHUB}.",
                    "So you can get fixes; it can be switched off.",
                ),
            ),
        ),
        Paragraph(
            "The settings in plain words. Adaptive thinking lets the model "
            "decide how much to reason before it answers; effort “"
            f"{_n('effort')}” asks it to be thorough. The output limit stops "
            "a runaway answer. No temperature or other randomness setting is "
            "sent, and the model is offered no tools, so it can only reply "
            "with text. That reply must fit a fixed JSON schema: a "
            "machine-checkable form listing exactly which fields and values "
            "are allowed."
        ),
    ),
    rail="The engine",
)


_BOUNDARY = Section(
    "boundary",
    "The boundary",
    "What leaves your computer, and where it goes",
    (
        DataFlowDiagram(
            description=(
                "Data-flow diagram, in words. Your computer holds your "
                "original files, the working copies, the outputs and run "
                "records, the template cache and, if you chose Remember, your "
                f"key. One solid arrow leaves it: to Anthropic at {_ANTHROPIC}, "
                "only during a run that needs AI, carrying paragraph excerpts "
                "and your key. Two dashed arrows are automatic or optional: "
                f"to {_GITHUB} at most once a day for a small update file, "
                "with no document data; and from GitHub only when you click "
                "Download & Install, for the installer."
            ),
            machine_label="Your computer",
            machine_items=(
                "Your originals (read only)",
                "Working copies (temporary)",
                "Outputs and run records",
                "Template cache",
                "Key in Credential Manager (optional)",
            ),
            destinations=(
                (
                    f"Anthropic — {_ANTHROPIC}",
                    "Only in a run that needs AI: paragraph excerpts + key",
                    False,
                ),
                (
                    f"GitHub — {_GITHUB}",
                    "Automatic, once a day at most: update file, no document data",
                    True,
                ),
                (
                    "GitHub release download",
                    "Only when you click Download & Install: the installer",
                    True,
                ),
            ),
            legend="Solid: only in a run that needs AI. Dashed: automatic or on request.",
        ),
        Contrast(
            "Leaves your machine",
            (
                f"To {_ANTHROPIC}, only during a run that needs AI: up to "
                f"{_n('target_text_chars')} characters of each paragraph the "
                f"rules could not settle, up to {_n('neighbour_text_chars')} "
                "characters of the paragraphs before and after it, its style "
                "name and numbering and formatting hints, the role list and "
                "the app's prompt. For a new template, every template "
                f"paragraph (up to {_n('architect_text_chars')} characters "
                "each) with its style and numbering catalogs. Your key "
                "travels in the request header.",
                "Before a template's first request and before each Haiku "
                "request, the same content to Anthropic's token-counting "
                "service on the same host. Template follow-ups and Sonnet "
                "requests are not counted first.",
                "With every Anthropic request, the standard headers of "
                "Anthropic's software library: its version, your operating "
                "system, processor type and Python version.",
                f"To {_GITHUB}, automatically at most once a day and whenever "
                "you click Check for Updates: one request for a small file "
                f"named latest.json, identified as {_n('update_user_agent')}. "
                "Like any web request it shows GitHub your IP address.",
                "From GitHub, only when you click Download & Install: the "
                "installer named in that file.",
                "Through your own browser or mail program, only when you click "
                "a link: Anthropic's console, the releases page, the author's "
                "LinkedIn page or email address.",
            ),
            "Stays on your machine",
            (
                "Your original template and targets, which are only read.",
                "Whole documents, tables, images, headers and footers, "
                "comments and footnotes.",
                "File names and folder paths.",
                "Every output and every run record.",
                "The template cache and the update-check state.",
                "Your key, if you chose Remember.",
                "All rule-based decisions, conversions and checks.",
            ),
        ),
        Note(
            "Settings that can move this boundary",
            f"The environment variable {_n('anthropic_base_url_env')}, which "
            "Anthropic's library honours, sends every AI request, excerpts and "
            f"key included, to another host instead of {_ANTHROPIC}. "
            f"{_n('update_url_env')} sends the update check elsewhere (it "
            "must still be HTTPS). The app never sets either one and does not "
            "show them in the window. If your IT department set one, your "
            "traffic goes where they pointed it. Separately, both Anthropic's "
            "library and the update check follow your computer's proxy "
            "settings, so on a network with a proxy the same requests pass "
            "through it on their way to the same hosts.",
        ),
    ),
    rail="What leaves your computer",
)


_CARDS: tuple[RuntimeCard, ...] = (
    RuntimeCard(
        "Opening the app",
        you_do="Start Specification Formatter.",
        what_runs=(
            "The window is built. The key field is filled from the "
            f"{_n('api_key_env')} environment variable if it is set, otherwise "
            "from Windows Credential Manager if you chose Remember earlier. "
            "The update check (next card) is scheduled."
        ),
        what_is_sent="Nothing.",
        ai_involved="None.",
        bounded_by=(
            "If Credential Manager cannot be read, the key field stays empty "
            "and nothing else changes."
        ),
    ),
    RuntimeCard(
        "Automatic update check",
        you_do=(
            f"Nothing. It runs about {_n('update_check_delay_s')} seconds "
            "after the window opens, if the last check was a day or more ago."
        ),
        what_runs=(
            "A background task downloads latest.json, reads the version, "
            "installer address and SHA-256 fingerprint in it, and compares "
            "the version with yours on your computer. It writes the time of "
            f"the check to {_n('update_state_file')}. An Update available "
            "window appears only if the version is newer and you have not "
            "skipped it."
        ),
        what_is_sent=(
            f"One HTTPS request to {_n('update_manifest_url')} (GitHub "
            "forwards it to its download servers). No document data, key or "
            "version number."
        ),
        ai_involved="None.",
        bounded_by=(
            f"{_n('update_manifest_timeout_s')}-second timeout; reads at most "
            f"{_n('update_manifest_max_kib')} KiB; refuses anything but HTTPS, "
            "including after redirects. A failure is silent apart from the "
            f"footer. {_n('update_disable_env')} turns it off."
        ),
        automatic=True,
    ),
    RuntimeCard(
        "Check for Updates",
        you_do="Click Check for Updates in the footer.",
        what_runs="The same check as the automatic one, then a message with the result.",
        what_is_sent=f"The same single request to {_GITHUB}.",
        ai_involved="None.",
        bounded_by=(
            "The same timeout, size cap and HTTPS rule. If checks are switched "
            "off, it says so and sends nothing."
        ),
    ),
    RuntimeCard(
        "Download & Install, Later, Skip this Version",
        you_do="Choose one of the buttons in the Update available window.",
        what_runs=(
            "Download & Install is refused while a run is active. Otherwise "
            "the installer is downloaded to a temporary name in "
            f"{_n('config_dir_windows')}\\{_n('updates_subdir')}, its SHA-256 "
            "fingerprint is checked against latest.json, and only then is it "
            "renamed. You are asked once more; yes launches the installer and "
            "closes the app. Later closes the window. Skip this Version "
            "stops the automatic prompt for that version. The View this "
            "release on GitHub link opens the releases page in your browser."
        ),
        what_is_sent=(
            "Download & Install: one HTTPS request for the installer address "
            "named in latest.json. Later and Skip: nothing."
        ),
        ai_involved="None.",
        bounded_by=(
            f"At most {_n('installer_max_mib')} MiB; "
            f"{_n('installer_timeout_s')}-second timeout per network read; "
            "HTTPS only. A fingerprint mismatch deletes the file, so it can "
            "never be run. The installer shows the license page again. "
            "Downloaded installers are not deleted afterwards."
        ),
    ),
    RuntimeCard(
        "Choose File (architect template)",
        you_do="Click Choose File and pick a .docx.",
        what_runs=(
            "Windows' file dialog. The path is remembered; the file is not "
            "opened until you click FORMAT SPECS."
        ),
        what_is_sent="Nothing.",
        ai_involved="None.",
        bounded_by="Disabled in the two no-template modes and during a run.",
    ),
    RuntimeCard(
        "Add Files, Add Folder, Clear",
        you_do="Add target files or a folder of them, or clear the list.",
        what_runs=(
            "The preview lists the .docx files directly inside a folder (not "
            "in its subfolders), leaving out Word lock files, this app's own "
            "outputs, and the architect template. The output folder defaults "
            "to a Formatted Specs folder beside the first file, or inside the "
            "first folder."
        ),
        what_is_sent="Nothing.",
        ai_involved="None.",
        bounded_by="Only the folder listing is read. Locked during a run.",
    ),
    RuntimeCard(
        "Choose the output mode and Target AI model",
        you_do=(
            "Pick one of the four modes, and "
            f"{_n('target_model_default_label')} or {_n('target_model_alt_label')}."
        ),
        what_runs=(
            "The two no-template modes grey out the template field. The model "
            "choice is stored for the next run."
        ),
        what_is_sent="Nothing.",
        ai_involved="None. (This choice decides which model a later run uses.)",
        bounded_by="Both are locked while a run is active.",
    ),
    RuntimeCard(
        "Output folder",
        you_do="Type a folder or click Choose Folder.",
        what_runs="The path is remembered for the next run.",
        what_is_sent="Nothing.",
        ai_involved="None.",
        bounded_by=(
            "Every run writes into a new subfolder of it; nothing already "
            "there is changed."
        ),
    ),
    RuntimeCard(
        "API key, Show, Remember, How do I get one?",
        you_do="Paste a key, tick Show or Remember, or click How do I get one?",
        what_runs=(
            "Show reveals the key in the field. Ticking Remember saves the "
            "key to Windows Credential Manager straight away (service "
            f"{_n('keyring_service')}, user name {_n('keyring_user')}) and "
            "again when a run starts; unticking deletes it. How do I get one? "
            "opens a help window whose console link opens in your browser."
        ),
        what_is_sent=(
            "Nothing. The console link opens "
            f"{_n('console_keys_url')} in your browser only when you click it."
        ),
        ai_involved="None.",
        bounded_by=(
            "If Credential Manager refuses the key, Remember unticks itself "
            "and a message says the key will not be remembered."
        ),
    ),
    RuntimeCard(
        "FORMAT SPECS: checking inputs",
        you_do="Click FORMAT SPECS.",
        what_runs=(
            "Checks that each input is a saved .docx, not a Word lock file, "
            "not one of this app's outputs, and that the template is not "
            "also a target. Checks the output folder can be written. Only "
            "then creates the run folder <UTC time>_<mode>_<run id> and "
            "locks every input on the screen."
        ),
        what_is_sent="Nothing.",
        ai_involved="None.",
        bounded_by=(
            "A problem here stops the run before the run folder is created "
            "and before anything is sent."
        ),
    ),
    RuntimeCard(
        "Template step: a template already analyzed",
        you_do="Run a template mode with a template the app has analyzed before.",
        what_runs=(
            "The template's SHA-256 fingerprint is computed and the cache is "
            "searched for a profile with the same fingerprint, app engine, "
            "model, prompts and format version. Every file in that profile "
            "is checked against its recorded fingerprint, then it is used."
        ),
        what_is_sent="Nothing.",
        ai_involved="None.",
        bounded_by=(
            "A damaged or mismatched profile is ignored, never used. If the "
            "template changes during the run, the run stops."
        ),
        automatic=True,
    ),
    RuntimeCard(
        "Template step: a new or changed template",
        you_do="Run a template mode with a template the app has not analyzed.",
        what_runs=(
            "The template is copied, unpacked within size limits and read "
            "paragraph by paragraph. Its size is estimated and then measured "
            "with Anthropic's token counter. One request asks the model to map "
            "CSI roles to example paragraphs; up to "
            f"{_n('architect_patch_attempts')} follow-up requests cover any "
            "paragraphs it missed. Fixed rules then correct roles that "
            "strong text signals contradict, coverage must be complete, "
            "styles are copied from the example paragraphs, the page layout "
            "and headers are captured, and a fingerprinted profile is saved "
            f"to the cache. Older profiles of that template beyond the newest "
            f"{_n('cache_keep')} are deleted. In Canadian mode the template's "
            "numbering is checked before any target starts."
        ),
        what_is_sent=(
            f"To {_ANTHROPIC}: every template paragraph (up to "
            f"{_n('architect_text_chars')} characters each) with its style, "
            "numbering and formatting hints, the template's style and "
            "numbering catalogs, fingerprints of its headers and footers, "
            "and the app's two template prompts. The same content goes to "
            "the token counter first."
        ),
        ai_involved=f"{_ARCH}: {_SETTINGS}.",
        bounded_by=(
            f"Refused before any request above {_n('architect_input_cap')} "
            "input tokens. These limits apply to each request, and there are "
            f"at most {_n('architect_requests')} (the first plus up to "
            f"{_n('architect_patch_attempts')} follow-ups): up to "
            f"{_n('architect_transport_attempts')} tries for network or "
            f"server errors, with {_n('backoff_first_s')}- and "
            f"{_n('backoff_second_s')}-second waits, and the whole request "
            f"repeated for an unusable answer, {_n('architect_response_attempts')} "
            f"times in all, so up to {_n('architect_sends_per_request')} sends "
            "per request, plus one more without the answer format if Anthropic "
            "cannot compile it. A refusal is final. Connection timeout "
            f"{_n('connect_timeout_s')} seconds, answer timeout "
            f"{_n('read_timeout_min')} minutes. A failure stops the whole run "
            "before any target, with records written; tokens already used are "
            "billed and recorded."
        ),
        automatic=True,
    ),
    RuntimeCard(
        "Template step: the no-template modes",
        you_do="Run Convert CSI to Canadian - no template, or Convert Canadian back to CSI.",
        what_runs=(
            "For the first, the built-in nine-level Canadian scheme is "
            "generated from constants and checked against the same rules a "
            "template's numbering must pass. For the second, nothing is "
            "needed: numbers are written as text. Nothing is cached."
        ),
        what_is_sent="Nothing.",
        ai_involved="None.",
        bounded_by=(
            "A template supplied through the Python interface is refused "
            "rather than ignored."
        ),
        automatic=True,
    ),
    RuntimeCard(
        "Each target: copy, unpack and rule-based labelling",
        you_do="Nothing more; this runs for every target.",
        what_runs=(
            f"Up to {_n('gui_workers')} targets are processed at once. Each "
            "is copied to a short temporary folder and fingerprinted, "
            "unpacked within the size limits, and read paragraph by "
            "paragraph with its styles and numbering. Fixed rules label "
            "every paragraph whose role Word's numbering or unmistakable "
            "text proves. What is left is counted."
        ),
        what_is_sent="Nothing.",
        ai_involved="None.",
        bounded_by=(
            f"Unpacking: at most {_n('package_entries')} entries, "
            f"{_n('package_total_mib')} MiB in all, {_n('package_part_mib')} "
            f"MiB per part, compression no higher than "
            f"{_n('compression_ratio')} to one. A target with more than "
            f"{_n('target_paragraph_cap')} unlabelled paragraphs fails here, "
            f"before any request ({_n('target_paragraph_cap_env')} changes "
            "the limit). A failure affects only this target."
        ),
        automatic=True,
    ),
    RuntimeCard(
        "Each target: AI labelling of the rest",
        you_do="Nothing more; this happens only if some paragraphs are still unlabelled.",
        what_runs=(
            "Long targets are split into chunks of about "
            f"{_n('chunk_tokens')} estimated tokens or "
            f"{_n('chunk_paragraphs')} paragraphs, overlapping by "
            f"{_n('chunk_overlap')} paragraphs, up to {_n('chunk_workers')} "
            "at once. Each Haiku request is measured with Anthropic's token "
            "counter first (Sonnet requests are not). The answer is streamed "
            "back and checked: only paragraphs that "
            "were sent, each exactly once, each with an allowed role. If two "
            "chunks disagree in their overlap, those paragraphs are asked "
            "once more. Finally every paragraph must have exactly one answer."
        ),
        what_is_sent=(
            f"To {_ANTHROPIC}: for each unlabelled paragraph, up to "
            f"{_n('target_text_chars')} characters of its text, up to "
            f"{_n('neighbour_text_chars')} characters of its neighbours, its "
            "style name and numbering and formatting hints; the role list; "
            "the app's target prompt; your key in the header. Table "
            "paragraphs are never sent."
        ),
        ai_involved=(
            f"{_HAIKU} by default, or {_SONNET} if you chose it: {_SETTINGS}. "
            f"{_n('target_effort_env')} can change the effort; the app never "
            "sets it."
        ),
        bounded_by=(
            "Each Haiku request must measure under "
            f"{_n('haiku_prompt_budget')} tokens, or the target fails "
            f"before it is sent. At most {_n('max_concurrent_requests')} "
            "requests are open at once across the whole run. For each chunk, "
            f"network, rate limit and server errors get {_n('target_transport_retries')} "
            f"retries (rate-limit waits capped at {_n('retry_after_cap_s')} "
            f"seconds); an unusable answer gets {_n('target_regenerations')} "
            "more tries; a refusal is final. No key means the target fails "
            "here. A failure affects only this target; tokens already used "
            "are billed and recorded."
        ),
        automatic=True,
    ),
    RuntimeCard(
        "Each target: apply, verify, publish",
        you_do="Nothing more.",
        what_runs=(
            "The mode is applied. Format only brings in the template's "
            "styles without replacing any style your file already has, and "
            "applies its page layout, headers and footers while keeping your "
            "own numbering. Convert CSI to Canadian removes typed CSI markers "
            "and puts the converted paragraphs on the Canadian list (the "
            "template's or the built-in one); with a template it also applies "
            "the template's styles, layout, headers and footers. Convert "
            "Canadian back to CSI writes each number into the text. Then "
            "every check runs on the finished file: body text "
            "run by run, numbering, tracked changes by author and kind, "
            "indentation, styles read back, and every untouched part of the "
            "file byte for byte. The original is fingerprinted again, and "
            "only then is the file copied into the run folder in one step "
            "and its fingerprint confirmed."
        ),
        what_is_sent="Nothing.",
        ai_involved="None.",
        bounded_by=(
            "Any failed check means no output for this target, with an error "
            "code, a fixed explanation and, where known, the paragraph's "
            "location. Your "
            "other targets continue."
        ),
        automatic=True,
    ),
    RuntimeCard(
        "Writing the run records",
        you_do="Nothing more.",
        what_runs=(
            "One audit.json per target, then run.log, diagnostics.jsonl and, "
            "last, run.json. Temporary files are removed and a message box "
            "reports the result."
        ),
        what_is_sent="Nothing.",
        ai_involved="None.",
        bounded_by=(
            "Written even when some or all targets failed. If writing them "
            "fails after outputs were published, a run.json saying so is "
            "still written."
        ),
        automatic=True,
    ),
    RuntimeCard(
        "Cancel",
        you_do="Click Cancel during a run.",
        what_runs=(
            "A stop signal is set. No new target or request starts, waits "
            "between retries end at once, and an answer still arriving is cut "
            "off at its next piece. Targets not yet published are marked "
            "cancelled. The run records are written."
        ),
        what_is_sent="Nothing new.",
        ai_involved="None started.",
        bounded_by=(
            "Outputs already published stay. Tokens of answers that finished "
            "are billed and recorded. An answer cut off midway is recorded as "
            "unknown, not zero, and Anthropic may still bill what it produced."
        ),
    ),
    RuntimeCard(
        "Closing the window during a run",
        you_do="Close the window while a run is active.",
        what_runs=(
            "You are asked whether to cancel. Yes cancels as above and closes "
            "the window once the records are saved; No keeps the run going."
        ),
        what_is_sent="Nothing new.",
        ai_involved="None started.",
        bounded_by="The same as Cancel.",
    ),
    RuntimeCard(
        "Open Output Folder",
        you_do="Click Open Output Folder.",
        what_runs="File Explorer opens the folder of the last run.",
        what_is_sent="Nothing.",
        ai_involved="None.",
        bounded_by="Nothing in the folder is changed.",
    ),
    RuntimeCard(
        "How to Use, How It Works, Why Trust It?, About",
        you_do="Click a help button at the top of the window.",
        what_runs=(
            "A window of text opens. Links in About (email, LinkedIn, "
            "releases) open in your browser or mail program when clicked."
        ),
        what_is_sent="Nothing, unless you click a link.",
        ai_involved="None.",
        bounded_by="Escape closes the topmost window only.",
    ),
    RuntimeCard(
        "Installing and uninstalling",
        you_do="Run SpecificationFormatterSetup.exe, or uninstall from Windows Settings.",
        what_runs=(
            "You must accept the license before anything is installed. It "
            "installs for your user only, with a Start menu shortcut and an "
            "optional desktop icon, and offers to start the app at the end. "
            "Uninstalling removes the program files."
        ),
        what_is_sent="Nothing.",
        ai_involved="None.",
        bounded_by=(
            "No startup entry, service or scheduled task is added. "
            "Uninstalling leaves the template cache, downloaded installers "
            "and any remembered key; delete them by hand (see Security)."
        ),
    ),
)


_RUNTIME = Section(
    "runtime",
    "Runtime",
    "What runs when you click",
    (
        Paragraph(
            "One card per action, including what happens on its own. Every "
            "card answers the same five questions in the same order, so you "
            "can compare them. Cards marked automatic run without a click of "
            "their own."
        ),
        RuntimeCards(_CARDS),
    ),
    rail="What runs when you click",
)


_BLAST_RADIUS = Section(
    "blast-radius",
    "Blast radius",
    "What the AI may touch, and what it cannot",
    (
        Table(
            ("Capability", "What it can do", "Constraint"),
            (
                (
                    "Label a target paragraph",
                    "Give one role to each paragraph it was sent.",
                    "The role must be one of the template's role names. Answers "
                    "about paragraphs it was not sent, duplicates and unknown "
                    "roles are rejected. It cannot change a rule's answer. "
                    "Every paragraph needs exactly one answer.",
                ),
                (
                    "Mark a target paragraph as not CSI",
                    "Leave a paragraph exactly as it is.",
                    f"Only one reason is allowed: {_n('target_ignore_reason')}. "
                    "An ignored paragraph receives no edits at all.",
                ),
                (
                    "Choose a template's example paragraphs",
                    "Say which template paragraph shows each role's formatting.",
                    "The formatting itself is copied from that paragraph, never "
                    "written by the model. Strong text signals overrule it, and "
                    "the style each one is based on is taken from the template.",
                ),
                (
                    "Name the generated styles",
                    "Choose the display name of each CSI_ style it creates.",
                    "The name is escaped as plain text. The style's ID is fixed "
                    "by the app.",
                ),
                (
                    "Write notes about a template",
                    "Add free-text remarks to its answer.",
                    "Kept only in the template cache's audit file; never in "
                    "your output or the run records.",
                ),
                (
                    "Tools",
                    "None.",
                    "No tools are offered, so the model can only reply with text "
                    "in the required form.",
                ),
            ),
        ),
        Paragraph(
            "What the AI categorically cannot do: open, read or write any "
            "file; reach the network or any address; run code; see a whole "
            "document, a file name, an image, a header or a footer; set a "
            "formatting value; change a word of your text; or decide a number. "
            "Numbers written by the Canadian-to-CSI conversion are counted by "
            "the app from Word's own list."
        ),
    ),
    rail="What the AI may touch",
)


_TERMS = Section(
    "terms",
    "Definitions",
    "What the key terms mean, exactly",
    (
        ExactClaim(
            "“Never modified” / “read only”",
            "The app opens your template and targets only to copy them. All "
            "work happens on the copy, and each original is fingerprinted "
            "again before its output is published.",
            "It does not stop another program, such as Word or a sync client, "
            "from changing the file meanwhile. It notices and stops instead.",
        ),
        ExactClaim(
            "“Verified”",
            "The output passed every automated check for its mode: complete "
            "coverage, body text run by run except predicted changes, "
            "numbering, tracked changes, indentation, styles read back, and "
            "every untouched part byte for byte.",
            "It does not mean the right role was chosen, that every page "
            "looks right in Word, or that the specification is technically "
            "correct.",
        ),
        ExactClaim(
            "“Unchanged” (Format only)",
            "Every body paragraph keeps its exact text, including tabs, "
            "breaks, spaces, hyphens and tracked deletions, and its effective "
            "list numbering.",
            "Headers and footers are replaced by the template's. The page can "
            "reflow because fonts, margins and defaults change. The file is "
            "not byte-identical.",
        ),
        ExactClaim(
            "“Deterministic” / “rules”",
            "Decided by fixed code on your computer: the same file always "
            "gives the same answer, with no AI and no network.",
            "It does not mean correct. A rule can be confidently wrong, and "
            "the AI cannot overrule it.",
        ),
        ExactClaim(
            "“No document text in the records”",
            "run.json, audit.json, run.log and diagnostics.jsonl carry counts, "
            "codes, fingerprints, model names and locations given as a "
            "section number and a heading count, never paragraph text, and "
            "your key "
            "is removed wherever it could appear.",
            "File names and folder paths are recorded. The template cache "
            f"does keep the first {_n('architect_text_chars')} characters of "
            "each template paragraph.",
        ),
        ExactClaim(
            "“Cached”",
            "A template analysis is reused only when the template's "
            "fingerprint, the app's engine, the model, the prompts and the "
            "format version all match, and every file in it checks out.",
            "The cache never expires. The app does not encrypt it; like your "
            "other files it is protected only by your Windows account.",
        ),
        ExactClaim(
            "“Fail closed”",
            "When the app cannot prove a step, it stops that target and "
            "publishes nothing for it, rather than guess.",
            "It does not mean you get a partial result. A refused or failed "
            "target still costs whatever requests were already made.",
        ),
        ExactClaim(
            "“Tokens counted” / “usage”",
            "The counts Anthropic reported in each response, totalled per "
            "target and per run. A missing count is marked unknown.",
            "It is not a price and not your invoice; Anthropic's bill is the "
            "authority.",
        ),
    ),
    rail="Key terms, exactly",
)


_NO_AI = Section(
    "no-ai",
    "No AI here",
    "The parts with no AI in them",
    (
        Bullets(
            (
                "Copying your files to private snapshots and fingerprinting them.",
                "Unpacking and checking every Word file against the size and "
                "safety limits.",
                "Reading paragraphs, styles, numbering, headers and footers.",
                "The rules that label paragraphs from Word's numbering and "
                "unmistakable text.",
                "Choosing which paragraphs need the AI at all.",
                "Checking every AI answer.",
                "Copying formatting from the template's example paragraphs.",
                "Importing styles under private names.",
                "Applying page layout, headers and footers.",
                "Both Canadian conversions, including every number written "
                "as text.",
                "The built-in Canadian numbering scheme.",
                "Every verification check and the decision to publish.",
                "Publishing outputs and writing run records.",
                "Checking and reusing cached template analyses.",
                "The update check and installer fingerprint check.",
            )
        ),
        Note(
            "With the network disconnected",
            "These still work: the whole window and its help; both "
            "no-template modes; and a template mode with a template already "
            "in the cache. In each case, a target works when rules label every "
            "paragraph. What cannot work is analyzing a new template and "
            "labelling paragraphs the rules could not settle: those targets "
            "fail after their retries, with a record saying so. The update "
            "check fails silently.",
        ),
    ),
    rail="The parts with no AI",
)


_SECURITY = Section(
    "security",
    "Security",
    "Security and privacy, mechanism by mechanism",
    (
        Table(
            ("Concern", "How it is handled"),
            (
                (
                    "Your API key at rest",
                    "Held in memory by default. With Remember, stored in "
                    "Windows Credential Manager under "
                    f"{_n('keyring_service')} / {_n('keyring_user')}; unticking "
                    "deletes it. Never written to a file by the app.",
                ),
                (
                    "Your API key in logs and records",
                    "Replaced wherever it could appear in run.log, run.json, "
                    "audit.json, diagnostics.jsonl and on-screen errors.",
                ),
                (
                    "A malicious or damaged Word file",
                    "Treated as untrusted. Paths that escape the file, "
                    "duplicates and links are refused; size is capped at "
                    f"{_n('package_entries')} entries, "
                    f"{_n('package_total_mib')} MiB, {_n('package_part_mib')} "
                    "MiB per part and a compression ratio of "
                    f"{_n('compression_ratio')} to one. Header and footer "
                    f"images: {_n('hf_media_mib')} MiB each, "
                    f"{_n('hf_media_total_mib')} MiB in all.",
                ),
                (
                    "XML tricks (entity expansion)",
                    "Any DOCTYPE or ENTITY declaration is refused before the "
                    "file is parsed, in any text encoding.",
                ),
                (
                    "Links inside a Word file",
                    "Recorded, never followed: no internet address, network "
                    "share or local path named by a document is opened.",
                ),
                (
                    "A tampered update",
                    "The version file must arrive over HTTPS, and the "
                    "installer must match its SHA-256 fingerprint before it can "
                    "be run. The app is not code-signed, so Windows "
                    "SmartScreen warns on first install.",
                ),
                (
                    "A tampered or stale template cache",
                    "Every cached file is checked against its recorded "
                    "fingerprint and the profile against the current engine; "
                    "anything that does not match is ignored.",
                ),
                (
                    "Listening for connections",
                    "None. The app runs no server and opens no port.",
                ),
                (
                    "Administrator rights",
                    "Not needed. The installer installs for your user only.",
                ),
                (
                    "Data left on disk",
                    f"Template cache: {_n('cache_dir_windows')}. Update state "
                    f"and downloaded installers: {_n('config_dir_windows')}. "
                    "Both are plain files protected only by your Windows "
                    "account, and both survive uninstalling. Delete the "
                    "folders to remove them; untick Remember first to delete "
                    "the key.",
                ),
                (
                    "Traffic redirected by settings",
                    f"{_n('anthropic_base_url_env')} and "
                    f"{_n('update_url_env')} can point traffic elsewhere; see "
                    "the boundary section.",
                ),
            ),
        ),
    ),
    rail="Security and privacy",
)


_MONEY = Section(
    "money",
    "Money",
    "What it costs, and who bills you",
    (
        KeyPoints(
            (
                (
                    "Who bills you.",
                    "Anthropic, directly, for model use on the API account "
                    "your key belongs to. The app has no subscription, no "
                    "server and no share of the spend.",
                ),
                (
                    "What is billed.",
                    "The input and output tokens of each model request, "
                    "including the model's thinking. Anthropic documents its "
                    "token-counting service as free.",
                ),
                (
                    "When nothing is spent.",
                    "A run spends nothing only when both hold: no template "
                    "needs a fresh analysis (it is already cached, or the "
                    "mode uses none), and rules label every paragraph of "
                    "every target. A new template is paid for even if rules "
                    "then settle every target; a cached one does not stop "
                    "targets from needing the AI. A run that fails its input "
                    "checks also spends nothing.",
                ),
                (
                    "What is estimated.",
                    "The app never shows a price. Its only estimates are "
                    "internal limits: a size estimate used to split long "
                    "targets, and Anthropic's token count, which Anthropic "
                    "itself calls an estimate.",
                ),
                (
                    "What is exact.",
                    "The token counts Anthropic reports in each response, "
                    "recorded per target in audit.json and per run in "
                    "run.json (diagnostics.usage). If a count never arrives it "
                    "is marked unknown and the total is a lower bound.",
                ),
                (
                    "Failed and stopped work.",
                    "Requests that finished are billed even if the target then "
                    "fails or you cancel, and they are counted in the records. "
                    "An answer cut off by Cancel is recorded as unknown.",
                ),
            )
        ),
        Bullets(
            (
                f"A template is refused before any request above "
                f"{_n('architect_input_cap')} input tokens.",
                f"A target with more than {_n('target_paragraph_cap')} "
                "paragraphs needing the AI is refused before any request.",
                f"Each Haiku request must measure under "
                f"{_n('haiku_prompt_budget')} tokens. The code's stated reason "
                f"is a price step at {_n('haiku_price_step')} tokens; check "
                "Anthropic's pricing page for current rates.",
                "Retries and follow-up requests are counted limits, listed on "
                "the cards above.",
                f"At most {_n('max_concurrent_requests')} requests run at once.",
            )
        ),
    ),
    rail="Money",
)


_HONESTY = Section(
    "honesty",
    "Honesty",
    "What this does not do",
    (
        Bullets(
            (
                "It does not check or revise technical content: codes, "
                "standards, references, units, terminology or requirements.",
                "The AI can give a paragraph the wrong role. The text and "
                "numbering checks still pass, so the paragraph gets the wrong "
                "style and nothing automated notices. Guard: rules decide "
                "first, roles are limited to a fixed list, and you review "
                "the output.",
                "A rule can be confidently wrong, and the AI cannot overrule "
                "it. Guard: rules fire only on narrow, unmistakable shapes "
                "and leave anything ambiguous to the AI.",
                "The records do not show which paragraphs the AI labelled, "
                "only how many. You cannot trace one AI decision to its "
                "paragraph after the run.",
                "AI answers can differ between two runs of the same file.",
                "The checks prove the file's structure, not its appearance. "
                "Applying a template's page layout and defaults reflows "
                "pages, so look at every page in Word.",
                "In the template modes your headers and footers are replaced "
                "by the template's; the section number and title are filled "
                "in only where the app can confirm them, otherwise the target "
                "fails.",
                "Tables, text boxes and drawings are not restyled.",
                "Canadian conversion refuses documents whose numbering it "
                "cannot prove, rather than guess. A round trip to Canadian "
                "and back drops a dash typed after PART n, so PART 1 - "
                "GENERAL returns as PART 1 GENERAL.",
                "The template cache keeps excerpts of every template it has "
                "analyzed, indefinitely, and downloaded installers pile up; "
                "uninstalling removes neither.",
                "The installer is not code-signed. The window has a dark "
                "theme only, and the shipped build is for Windows.",
                "Token counts are not an invoice.",
            )
        ),
    ),
    rail="What it does not do",
)


_CHECK = Section(
    "check",
    "The point",
    "Check it yourself",
    (
        Paragraph("From least effort to most. Each says what you will see."),
        Bullets(
            (
                "Click Open Output Folder and open run.log in Notepad. You "
                "will see, for each target, either “Anthropic request "
                "skipped” or “Classifying unresolved paragraphs with "
                "Anthropic”, the counts of paragraphs decided by rules and "
                "sent to the AI, and every part of the file that changed.",
                "Open the target's audit.json. You will see its final role "
                "for every paragraph, the counts, the tokens used and, for a "
                "failure, an error code with a location.",
                "Open run.json and search it for your key and for a sentence "
                "from your specification. Neither is there. You will see the "
                "model IDs, the effort and diagnostics.usage.",
                "In PowerShell, run Get-FileHash on your original and compare "
                "it with source.sha256 in the audit. They match, and they "
                "still match after the run.",
                "In Word, use Review › Compare on the original and the "
                "Format only output. You will see formatting changes and no "
                "inserted or deleted body text.",
                "Compare diagnostics.usage in run.json with the Usage page of "
                "the Anthropic Console for the same period. Use a key that "
                "nothing else uses and the totals should agree; the Console "
                "is the bill.",
                "Disconnect from the network and run a no-template mode, or a "
                "template you have used before. Targets that rules settle "
                "fully still finish; others fail with a record.",
                f"Set {_n('update_disable_env')}=1 and click Check for "
                "Updates. It says checks are turned off and sends nothing.",
                "Watch a firewall log or the Network tab of Windows Resource "
                f"Monitor. You will see {_ANTHROPIC} only during runs that "
                f"need AI, and {_GITHUB} once a day at most, at launch, and "
                "whenever you click Check for Updates.",
                "Read the source and run its tests: pip install -r "
                "requirements-dev.txt, then python -m pytest. The ledger in "
                "docs/TRUST_CLAIMS.md names the code behind every claim here.",
            ),
            numbered=True,
        ),
        Links(
            "Further reading",
            (
                ("Anthropic Privacy Policy", "https://www.anthropic.com/legal/privacy"),
                (
                    "Anthropic Commercial Terms of Service",
                    "https://www.anthropic.com/legal/commercial-terms",
                ),
                ("Anthropic Trust Center", "https://trust.anthropic.com"),
                (
                    "Anthropic: API and data retention",
                    "https://platform.claude.com/docs/en/manage-claude/api-and-data-retention",
                ),
                (
                    "Anthropic: token counting",
                    "https://platform.claude.com/docs/en/build-with-claude/token-counting",
                ),
                ("Specification Formatter releases", _n("releases_url")),
                (
                    "Specification Formatter source code",
                    "https://github.com/abe-borg/spec-template-normalizer",
                ),
            ),
        ),
        Note(
            "Still not convinced?",
            "Good. That is the right stance for anything that touches a "
            "document with your name on it. Run it on copies first. Start "
            "with Format only, the mode that proves your words did not "
            "change. Compare every page with the original. Treat every "
            "paragraph's style as a draft until you have looked at it, and "
            "keep the run folder with the deliverable so the record travels "
            "with the file.",
        ),
    ),
    rail="Check it yourself",
)


DOSSIER_SECTIONS: tuple[Section, ...] = (
    _SHORT_ANSWER,
    _PROVENANCE,
    _ENGINE,
    _BOUNDARY,
    _RUNTIME,
    _BLAST_RADIUS,
    _TERMS,
    _NO_AI,
    _SECURITY,
    _MONEY,
    _HONESTY,
    _CHECK,
)


# ---------------------------------------------------------------------------
# Inspection helpers (tests) and the Markdown rendering (docs/TRUST.md)
# ---------------------------------------------------------------------------


def runtime_cards() -> tuple[RuntimeCard, ...]:
    return tuple(
        card
        for section in DOSSIER_SECTIONS
        for block in section.blocks
        if isinstance(block, RuntimeCards)
        for card in block.cards
    )


def _block_strings(block: Block) -> Iterator[str]:
    if isinstance(block, Paragraph):
        yield block.text
    elif isinstance(block, Bullets):
        yield from block.items
    elif isinstance(block, KeyPoints):
        for lead, text in block.items:
            yield lead
            yield text
    elif isinstance(block, Table):
        yield from block.columns
        for row in block.rows:
            yield from row
    elif isinstance(block, Contrast):
        yield block.left_title
        yield from block.left_items
        yield block.right_title
        yield from block.right_items
    elif isinstance(block, Note):
        yield block.title
        yield block.text
    elif isinstance(block, ExactClaim):
        yield block.term
        yield block.claim
        yield block.not_claimed
    elif isinstance(block, Links):
        yield block.title
        for label, _url in block.items:
            yield label
    elif isinstance(block, DataFlowDiagram):
        yield block.description
        yield block.machine_label
        yield from block.machine_items
        for label, arrow, _dashed in block.destinations:
            yield label
            yield arrow
        yield block.legend
    elif isinstance(block, RuntimeCards):
        for card in block.cards:
            yield card.title
            for label, value in card.rows():
                yield label
                yield value
    else:  # pragma: no cover - every block type is listed above
        raise TypeError(type(block).__name__)


def iter_copy() -> Iterator[str]:
    """Every user-visible string of both levels (link URLs excluded)."""

    yield LEVEL1_TITLE
    yield LEVEL1_LEAD
    for point in LEVEL1_POINTS:
        yield point.title
        yield point.body
    yield LEVEL1_CLOSING_HEADING
    yield LEVEL1_CLOSING_TEXT
    yield LEVEL1_BUTTON
    yield DOSSIER_TITLE
    yield DOSSIER_SUBTITLE
    for section in DOSSIER_SECTIONS:
        yield section.kicker
        yield section.title
        for block in section.blocks:
            yield from _block_strings(block)


def external_links() -> tuple[str, ...]:
    """Every URL the trust surfaces may open: the further-reading box only."""

    return tuple(
        url
        for section in DOSSIER_SECTIONS
        for block in section.blocks
        if isinstance(block, Links)
        for _label, url in block.items
    )


def _md_cell(text: str) -> str:
    return text.replace("|", "\\|").replace("\n", " ")


def _render_block_md(block: Block) -> list[str]:
    if isinstance(block, Paragraph):
        return [block.text, ""]
    if isinstance(block, Bullets):
        marker = (lambda i: f"{i}.") if block.numbered else (lambda _i: "-")
        return [f"{marker(i)} {item}" for i, item in enumerate(block.items, 1)] + [""]
    if isinstance(block, KeyPoints):
        return [f"- **{lead}** {text}" for lead, text in block.items] + [""]
    if isinstance(block, Table):
        lines = [
            "| " + " | ".join(_md_cell(c) for c in block.columns) + " |",
            "|" + "---|" * len(block.columns),
        ]
        lines += ["| " + " | ".join(_md_cell(c) for c in row) + " |" for row in block.rows]
        return lines + [""]
    if isinstance(block, Contrast):
        lines = [f"**{block.left_title}**", ""]
        lines += [f"- {item}" for item in block.left_items]
        lines += ["", f"**{block.right_title}**", ""]
        lines += [f"- {item}" for item in block.right_items]
        return lines + [""]
    if isinstance(block, Note):
        return [f"> **{block.title}**", ">", f"> {block.text}", ""]
    if isinstance(block, ExactClaim):
        return [
            f"#### {block.term}",
            "",
            f"> **The exact claim:** {block.claim}",
            "",
            f"**What it does not claim:** {block.not_claimed}",
            "",
        ]
    if isinstance(block, Links):
        lines = [f"**{block.title}**", ""]
        lines += [f"- [{label}]({url})" for label, url in block.items]
        return lines + [""]
    if isinstance(block, DataFlowDiagram):
        lines = ["```text", f"[{block.machine_label}]"]
        lines += [f"   {item}" for item in block.machine_items]
        for label, arrow, dashed in block.destinations:
            line = "- - ->" if dashed else "----->"
            lines.append(f"   {line} {label}: {arrow}")
        lines += [f"   ({block.legend})", "```", "", f"*{block.description}*", ""]
        return lines
    if isinstance(block, RuntimeCards):
        lines: list[str] = []
        for number, card in enumerate(block.cards, 1):
            suffix = " *(automatic)*" if card.automatic else ""
            lines += [f"#### {number}. {card.title}{suffix}", ""]
            lines += ["| | |", "|---|---|"]
            lines += [f"| **{label}** | {_md_cell(value)} |" for label, value in card.rows()]
            lines.append("")
        return lines
    raise TypeError(type(block).__name__)  # pragma: no cover


def render_markdown() -> str:
    """The two levels as one Markdown document (``docs/TRUST.md``)."""

    lines = [
        "<!-- Generated by `python -m spec_formatter.trust`; do not edit by hand.",
        "     Edit spec_formatter/trust.py and regenerate. If the implementation",
        "     changes, this changes with it. A trust document that has drifted",
        "     from the code is worse than none. -->",
        "",
        "# Why Trust It?",
        "",
        "This is the text of the app's **Why Trust It?** help topic and of the "
        "dossier its button opens. The claims ledger behind it is "
        "[TRUST_CLAIMS.md](TRUST_CLAIMS.md).",
        "",
        f"**{LEVEL1_LEAD}**",
        "",
    ]
    for point in LEVEL1_POINTS:
        lines += [f"- **{point.title}.** {point.body}"]
    lines += [
        "",
        f"> **{LEVEL1_CLOSING_HEADING}** {LEVEL1_CLOSING_TEXT}",
        "",
        "---",
        "",
        f"## {DOSSIER_TITLE}",
        "",
        f"*{DOSSIER_SUBTITLE}*",
        "",
        "**Contents:** "
        + " · ".join(f"[{s.title}](#{s.anchor})" for s in DOSSIER_SECTIONS),
        "",
    ]
    for section in DOSSIER_SECTIONS:
        lines += [
            f'<a id="{section.anchor}"></a>',
            "",
            f"### {section.title}",
            "",
            f"*{section.kicker}*",
            "",
        ]
        for block in section.blocks:
            lines += _render_block_md(block)
    return "\n".join(lines).rstrip() + "\n"


if __name__ == "__main__":  # pragma: no cover - documentation generator
    import sys

    sys.stdout.reconfigure(encoding="utf-8")  # type: ignore[union-attr]
    sys.stdout.write(render_markdown())
