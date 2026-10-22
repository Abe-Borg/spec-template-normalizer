"""The trust surfaces are a contract with the code.

``spec_formatter/trust.py`` quotes numbers, model IDs, limits, hosts, paths
and settings. Each one is a hand-written entry in ``trust.FACTS``, and each is
asserted here against the constant, literal or behaviour in the code that
makes it true. Changing that code fails this module until the copy (and
``docs/TRUST_CLAIMS.md``) is updated in the same change.

These tests need no display and no ``tkinter``; the dialog behaviour is in
``tests/test_trust_ui.py``.
"""

from __future__ import annotations

import ast
import inspect
import os
import re
from pathlib import Path
from unittest import mock
from urllib.parse import urlsplit

import pytest

import docx_decomposer
import llm_classifier as architect_classifier
import phase1_pipeline
from spec_formatter import app_paths, model_config, pipeline, secrets, trust, updates
from spec_formatter.style_application import batch_runner
from spec_formatter.style_application import docx_decomposer as package_limits
from spec_formatter.style_application.core import classification
from spec_formatter.style_application.core import llm_classifier as target_classifier
from spec_formatter.style_application.core import registry

ROOT = Path(__file__).resolve().parents[1]


def _source(relative: str) -> str:
    return (ROOT / relative).read_text(encoding="utf-8")


def _one(pattern: str, text: str) -> str:
    """The single value *pattern* captures in *text*; every match must agree."""

    values = set(re.findall(pattern, text))
    assert len(values) == 1, (pattern, values)
    return values.pop()


def _gui_constant(name: str) -> object:
    # gui.py needs tkinter to import; read the literal instead.
    return ast.literal_eval(_one(rf"(?m)^{name} = (.+)$", _source("gui.py")))


def _sdk_client_source() -> str:
    import anthropic._client

    return inspect.getsource(anthropic._client)


def _architect_call_source() -> str:
    return inspect.getsource(architect_classifier._call_api)


def _target_classify_source() -> str:
    return inspect.getsource(target_classifier.classify_target_document)


def _label_for(model_id: str) -> str:
    """``claude-opus-5-5`` -> ``Opus 5.5``: the label the copy prints."""

    _vendor, family, *version = model_id.split("-")
    return f"{family.title()} {'.'.join(version)}"


def _windows_path(fn) -> str:
    with mock.patch.dict(os.environ, {"LOCALAPPDATA": "%LOCALAPPDATA%"}):
        return str(fn()).replace(os.sep, "\\")


def _exact_division(value: float, unit: int) -> int:
    assert value % unit == 0, (value, unit)
    return int(value // unit)


def _timeouts(source: str) -> tuple[float, float]:
    read, connect = _one(r"httpx\.Timeout\(([\d.]+), connect=([\d.]+)\)", source)
    return float(read), float(connect)


def _shared_timeouts() -> tuple[float, float]:
    both = {
        _timeouts(inspect.getsource(architect_classifier.classify_document)),
        _timeouts(_target_classify_source()),
    }
    assert len(both) == 1, both
    return both.pop()


def _shared_effort() -> str:
    target_default = inspect.signature(pipeline.format_specifications).parameters[
        "target_effort"
    ].default
    architect = _one(r'output_config = \{"effort": "(\w+)"\}', _architect_call_source())
    assert target_default == architect
    return architect


def _shared_thinking() -> str:
    architect = _one(r'thinking=\{"type": "(\w+)"\}', _architect_call_source())
    target = _one(r'"thinking": \{"type": "(\w+)"\}', _target_classify_source())
    assert architect == target
    return architect


def _shared_max_tokens() -> int:
    architect = inspect.signature(architect_classifier._call_api).parameters["max_tokens"].default
    target = int(_one(r'"max_tokens": (\d+)', _target_classify_source()))
    assert architect == target
    return target


class _Rate(Exception):
    pass


class _Conn(Exception):
    pass


class _Status(Exception):
    pass


_ERRORS = {"RateLimitError": _Rate, "APIConnectionError": _Conn, "APIStatusError": _Status}


def _target_model_alt() -> tuple[str, str]:
    others = [
        (label, model)
        for label, model in model_config.TARGET_MODEL_CHOICES
        if model != model_config.DEFAULT_TARGET_MODEL
    ]
    assert len(others) == 1, others
    return others[0]


def _ignore_reason() -> str:
    schema = target_classifier._classification_output_config(["PART"])["format"]["schema"]
    reasons = schema["properties"]["ignored_paragraphs"]["items"]["properties"]["reason"]["enum"]
    assert len(reasons) == 1
    return reasons[0]


#: Where each pinned fact comes from. Every entry is the code, not the docs.
PINS = {
    "anthropic_host": lambda: _one(r'base_url = f"https://([^"/]+)"', _sdk_client_source()),
    "anthropic_base_url_env": lambda: _one(r'os\.environ\.get\("(ANTHROPIC_BASE_URL)"\)', _sdk_client_source()),
    "api_key_env": lambda: _one(r'os\.environ\.get\("(\w+)", ""\)', _source("gui.py")),
    "update_host": lambda: urlsplit(updates._DEFAULT_MANIFEST_URL).hostname,
    "update_manifest_url": lambda: updates._DEFAULT_MANIFEST_URL,
    "releases_url": updates.releases_page_url,
    "update_user_agent": lambda: updates._USER_AGENT,
    "update_check_delay_s": lambda: _gui_constant("AUTO_UPDATE_CHECK_DELAY_MS") / 1000,
    "update_check_interval_days": lambda: updates.DEFAULT_MIN_INTERVAL_DAYS,
    "update_disable_env": lambda: updates.ENV_DISABLE,
    "update_url_env": lambda: updates.ENV_UPDATE_URL,
    "update_manifest_timeout_s": lambda: updates.DEFAULT_MANIFEST_TIMEOUT,
    "update_manifest_max_kib": lambda: _exact_division(updates.MAX_MANIFEST_BYTES, 1024),
    "installer_timeout_s": lambda: updates.DEFAULT_DOWNLOAD_TIMEOUT,
    "installer_max_mib": lambda: _exact_division(updates.MAX_INSTALLER_BYTES, 1024 * 1024),
    "console_keys_url": lambda: _one(r'"(https://console\.anthropic\.com/[^"]+)"', _source("gui.py")),
    "architect_model": lambda: phase1_pipeline.DEFAULT_MODEL,
    "architect_model_label": lambda: _label_for(phase1_pipeline.DEFAULT_MODEL),
    "target_model_default": lambda: model_config.DEFAULT_TARGET_MODEL,
    "target_model_default_label": lambda: dict(
        (model, label) for label, model in model_config.TARGET_MODEL_CHOICES
    )[model_config.DEFAULT_TARGET_MODEL],
    "target_model_alt": lambda: _target_model_alt()[1],
    "target_model_alt_label": lambda: _target_model_alt()[0],
    "effort": _shared_effort,
    "thinking_type": _shared_thinking,
    "max_output_tokens": _shared_max_tokens,
    "target_effort_env": lambda: _one(
        r'os\.environ\.get\("(SPEC_FORMATTER_TARGET_EFFORT)"', inspect.getsource(pipeline)
    ),
    "target_ignore_reason": _ignore_reason,
    "connect_timeout_s": lambda: _shared_timeouts()[1],
    "read_timeout_s": lambda: _shared_timeouts()[0],
    "read_timeout_min": lambda: _exact_division(_shared_timeouts()[0], 60),
    "target_transport_retries": lambda: target_classifier._TRANSPORT_RETRIES,
    "retry_after_cap_s": lambda: target_classifier._MAX_RETRY_AFTER_SECONDS,
    "backoff_first_s": lambda: target_classifier._transport_retry_delay(_Conn(), 0, _ERRORS),
    "backoff_second_s": lambda: target_classifier._transport_retry_delay(_Conn(), 1, _ERRORS),
    "target_regenerations": lambda: int(
        _one(r"max_regenerations = (\d+)", _target_classify_source())
    ),
    "target_total_attempts": lambda: int(
        _one(r"max_regenerations = (\d+)", _target_classify_source())
    ) + 1,
    "architect_transport_attempts": lambda: int(
        _one(r"while attempt < (\d+)", _architect_call_source())
    ),
    "architect_response_attempts": lambda: architect_classifier.DEFAULT_RESPONSE_ATTEMPTS,
    "architect_patch_attempts": lambda: inspect.signature(
        architect_classifier.classify_document
    ).parameters["max_patch_attempts"].default,
    "chunk_tokens": lambda: target_classifier._MAX_BUNDLE_TOKENS,
    "chunk_paragraphs": lambda: _chunk_paragraphs(),
    "chunk_overlap": lambda: target_classifier._CHUNK_OVERLAP,
    "chunk_workers": lambda: int(_one(r"min\(len\(chunks\), (\d+)\)", _target_classify_source())),
    "max_concurrent_requests": lambda: target_classifier._DEFAULT_MAX_CONCURRENT_REQUESTS,
    "max_concurrent_requests_ceiling": lambda: target_classifier._MAX_CONCURRENT_REQUESTS_CEILING,
    "max_concurrent_requests_env": lambda: target_classifier._MAX_CONCURRENT_REQUESTS_ENV,
    "architect_text_chars": lambda: _architect_text_chars(),
    "target_text_chars": lambda: int(
        _one(r'"text": cleaned_text\[:(\d+)\]', inspect.getsource(classification.build_phase2_slim_bundle))
    ),
    "neighbour_text_chars": lambda: int(
        _one(r'\["text"\]\[:(\d+)\]', inspect.getsource(classification.build_phase2_slim_bundle))
    ),
    "architect_input_cap": lambda: architect_classifier.MAX_SINGLE_PASS_INPUT_TOKENS,
    "target_paragraph_cap": lambda: batch_runner.DEFAULT_MAX_TARGET_PARAGRAPHS,
    "target_paragraph_cap_env": lambda: batch_runner._MAX_TARGET_PARAGRAPHS_ENV,
    "haiku_prompt_budget": lambda: target_classifier._HAIKU_MAX_PROMPT_TOKENS,
    "haiku_price_step": lambda: int(
        _one(
            r"below Haiku's ([\d,]+)-token pricing boundary",
            inspect.getsource(target_classifier),
        ).replace(",", "")
    ),
    "package_entries": lambda: package_limits.MAX_PACKAGE_ENTRIES,
    "package_total_mib": lambda: _exact_division(package_limits.MAX_PACKAGE_UNCOMPRESSED_BYTES, 1024 * 1024),
    "package_part_mib": lambda: _exact_division(package_limits.MAX_PACKAGE_PART_BYTES, 1024 * 1024),
    "compression_ratio": lambda: package_limits.MAX_COMPRESSION_RATIO,
    "hf_media_mib": lambda: _exact_division(registry.MAX_HEADER_FOOTER_MEDIA_BYTES, 1024 * 1024),
    "hf_media_total_mib": lambda: _exact_division(
        registry.MAX_HEADER_FOOTER_MEDIA_TOTAL_BYTES, 1024 * 1024
    ),
    "gui_workers": lambda: _gui_constant("DEFAULT_MAX_WORKERS"),
    "api_max_workers": lambda: pipeline._MAX_WORKERS,
    "cache_dir_windows": lambda: _windows_path(pipeline.default_template_cache_dir),
    "config_dir_windows": lambda: _windows_path(app_paths.default_config_dir),
    "updates_subdir": lambda: _one(r'default_config_dir\(\) / "(\w+)"', _source("gui.py")),
    "update_state_file": lambda: updates.STATE_FILENAME,
    "cache_keep": lambda: pipeline._PROFILE_CACHE_KEEP,
    "keyring_service": lambda: secrets.KEYRING_SERVICE,
    "keyring_user": lambda: secrets.KEYRING_USERNAME,
}


def _chunk_paragraphs() -> int:
    source = inspect.getsource(target_classifier._split_bundle_into_chunks)
    whole = int(_one(r"len\(paragraphs\) <= (\d+)", source))
    per_chunk = int(_one(r"min\(start \+ (\d+),", source))
    assert whole == per_chunk
    return whole


def _architect_text_chars() -> int:
    source = inspect.getsource(docx_decomposer.build_slim_bundle)
    limits = {
        int(value)
        for match in re.findall(r"raw_text\[:(\d+)\]|len\(raw_text\) > (\d+)", source)
        for value in match
        if value
    }
    assert len(limits) == 1, limits
    return limits.pop()


def test_every_fact_has_a_pin_and_every_pin_a_fact():
    assert set(PINS) == set(trust.FACTS)


@pytest.mark.parametrize("key", sorted(trust.FACTS))
def test_fact_matches_code(key):
    assert PINS[key]() == trust.FACTS[key], (
        f"{key}: the code says {PINS[key]()!r}, the trust copy says "
        f"{trust.FACTS[key]!r}. Update spec_formatter/trust.py, docs/TRUST.md "
        "and docs/TRUST_CLAIMS.md together."
    )


def test_every_fact_is_in_the_claims_ledger():
    ledger = (ROOT / "docs" / "TRUST_CLAIMS.md").read_text(encoding="utf-8")
    missing = [key for key in trust.FACTS if f"`{key}`" not in ledger]
    assert not missing


def _numbers_allowed() -> set[str]:
    allowed = set(trust.NAMED_NUMBERS)
    for value in trust.FACTS.values():
        if isinstance(value, bool):
            continue
        if isinstance(value, int):
            allowed.update({str(value), f"{value:,}"})
        elif isinstance(value, float):
            allowed.add(f"{value:g}")
        else:
            allowed.update(re.findall(r"\d+(?:[.,]\d+)*", str(value)))
    return allowed


def test_every_number_in_the_copy_is_pinned():
    allowed = _numbers_allowed()
    stray = {
        (number, text[:70])
        for text in trust.iter_copy()
        for number in re.findall(r"\d+(?:[.,]\d+)*", text)
        if number not in allowed
    }
    assert not stray, (
        "A number in the trust copy is not in trust.FACTS (or NAMED_NUMBERS): "
        f"{sorted(stray)}"
    )


def test_level1_shape():
    assert 6 <= len(trust.LEVEL1_POINTS) <= 8
    assert len(trust.LEVEL1_LEAD.split(".")) == 2  # one sentence
    for point in trust.LEVEL1_POINTS:
        assert 3 <= len(point.title.split()) <= 9, point.title
        sentences = [part for part in re.split(r"(?<=[.;])\s", point.body) if part]
        assert 1 <= len(sentences) <= 4, point.title
    assert trust.LEVEL1_CLOSING_HEADING == "Not convinced?"
    assert trust.LEVEL1_CLOSING_TEXT == (
        "Fair. The points above are claims — here is the mechanism behind each "
        "one, action by action, plus how to audit any of it yourself."
    )
    assert trust.LEVEL1_BUTTON == "I'm not convinced — show me exactly what runs →"
    words = sum(len(text.split()) for text in (
        trust.LEVEL1_LEAD,
        *(p.title + " " + p.body for p in trust.LEVEL1_POINTS),
        trust.LEVEL1_CLOSING_TEXT,
    ))
    assert words <= 500  # readable in under two minutes


def test_level1_names_what_runs_without_the_user_and_its_switch():
    automatic = [p for p in trust.LEVEL1_POINTS if "without you" in p.title]
    assert len(automatic) == 1
    assert trust.FACTS["update_disable_env"] in automatic[0].body
    assert trust.FACTS["update_host"] in automatic[0].body


def test_dossier_sections_in_order_with_unique_anchors():
    anchors = [section.anchor for section in trust.DOSSIER_SECTIONS]
    assert anchors == [
        "short-answer", "provenance", "engine", "boundary", "runtime",
        "blast-radius", "terms", "no-ai", "security", "money", "honesty", "check",
    ]
    assert len(set(anchors)) == len(anchors)
    for section in trust.DOSSIER_SECTIONS:
        assert section.kicker and section.title and section.blocks
        assert len(section.rail_label) <= 28


def test_provenance_and_engine_tables_have_the_required_columns():
    tables = {
        section.anchor: [b for b in section.blocks if isinstance(b, trust.Table)]
        for section in trust.DOSSIER_SECTIONS
    }
    assert tables["provenance"][0].columns == ("Origin", "What it actually is", "How you can tell")
    assert tables["engine"][0].columns == ("Job", "Model or component", "Why")
    assert tables["blast-radius"][0].columns == ("Capability", "What it can do", "Constraint")
    assert tables["security"][0].columns == ("Concern", "How it is handled")
    notes = [
        b.title for s in trust.DOSSIER_SECTIONS for b in s.blocks if isinstance(b, trust.Note)
    ]
    assert "Things that are not happening" in notes
    assert "Still not convinced?" in notes
    still = next(
        b for s in trust.DOSSIER_SECTIONS for b in s.blocks
        if isinstance(b, trust.Note) and b.title == "Still not convinced?"
    )
    assert still.text.startswith("Good.")


def test_every_runtime_card_has_the_five_rows_in_order():
    cards = trust.runtime_cards()
    assert 12 <= len(cards) <= 24
    for card in cards:
        labels = tuple(label for label, _value in card.rows())
        assert labels == ("You do", "What runs", "What is sent", "AI involved", "Bounded by")
        assert all(value.strip() for _label, value in card.rows()), card.title


def test_ai_involved_says_none_wherever_nothing_reaches_anthropic():
    host = trust.FACTS["anthropic_host"]
    for card in trust.runtime_cards():
        sends_to_anthropic = host in card.what_is_sent
        if sends_to_anthropic:
            assert not card.ai_involved.startswith("None"), card.title
            assert any(
                trust.FACTS[key] in card.ai_involved
                for key in ("architect_model", "target_model_default")
            ), card.title
        else:
            assert card.ai_involved.startswith("None"), card.title


def test_boundary_accounts_for_every_network_destination():
    boundary = next(s for s in trust.DOSSIER_SECTIONS if s.anchor == "boundary")
    text = " ".join(trust._block_strings(boundary.blocks[1])) + " " + " ".join(
        trust._block_strings(boundary.blocks[2])
    )
    for needle in (
        trust.FACTS["anthropic_host"],
        "token-counting",
        trust.FACTS["update_host"],
        "Download & Install",
        "browser",
        trust.FACTS["anthropic_base_url_env"],
        trust.FACTS["update_url_env"],
    ):
        assert needle in text, needle


def _gui_controls() -> set[str]:
    """Every clickable control's label in gui.py, read from the source."""

    tree = ast.parse(_source("gui.py"))
    labels: set[str] = set()
    for node in ast.walk(tree):
        if not isinstance(node, ast.Call):
            continue
        name = getattr(node.func, "attr", "")
        if name in {"CTkButton", "CTkCheckBox", "CTkOptionMenu"}:
            for keyword in node.keywords:
                if keyword.arg == "text" and isinstance(keyword.value, ast.Constant):
                    labels.add(keyword.value.value)
        if name == "CTkLabel" and any(
            k.arg == "cursor" for k in node.keywords
        ):
            for keyword in node.keywords:
                if keyword.arg == "text" and isinstance(keyword.value, ast.Constant):
                    labels.add(keyword.value.value)
        if name == "_path_row" and isinstance(node.args[-1], ast.Constant):
            labels.add(node.args[-1].value)
    # The four help buttons are built in a loop from a tuple of labels.
    for node in ast.walk(tree):
        if isinstance(node, ast.Tuple) and len(node.elts) == 2:
            first = node.elts[0]
            if isinstance(first, ast.Constant) and first.value in {
                "How to Use", "How It Works", "Why Trust It?", "About"
            }:
                labels.add(first.value)
    return labels


#: Controls that need no card of their own, with the reason.
_CONTROLS_WITHOUT_CARDS = {
    "Close": "closes a help or trust window; nothing else happens",
}


def test_every_gui_control_has_a_runtime_card():
    controls = _gui_controls()
    assert {"FORMAT SPECS", "Cancel", "Check for Updates", "Download & Install"} <= controls
    copy = " ".join(
        f"{card.title} {card.you_do} {card.what_runs}" for card in trust.runtime_cards()
    )
    missing = sorted(
        label for label in controls - set(_CONTROLS_WITHOUT_CARDS) if label not in copy
    )
    assert not missing, f"controls with no runtime card: {missing}"


def test_external_links_are_only_the_further_reading_box():
    links = trust.external_links()
    assert links and all(url.startswith("https://") for url in links)
    shown_as_text = {
        trust.FACTS[key] for key in ("update_manifest_url", "releases_url", "console_keys_url")
    }
    for text in trust.iter_copy():
        for url in re.findall(r"https?://\S+", text):
            assert url.rstrip(".,)") in shown_as_text or url in links, url


def test_trust_components_load_no_external_asset():
    for relative in ("trust_ui.py", "spec_formatter/trust.py"):
        source = _source(relative)
        assert not _NETWORK_IMPORT.search(source.replace("import webbrowser", "")), relative
        for forbidden in (
            "PhotoImage", "CTkImage", "ImageTk", "font_file", "load_font",
            "FontManager", "iconbitmap", "iconphoto",
        ):
            assert forbidden not in source, (relative, forbidden)
        # No file is read at all: no open(), Path.read_* or resource lookup.
        assert not re.search(r"(?<![\w.])open\(|\.read_(?:text|bytes)\(|resource_path", source), relative
    ui = _source("trust_ui.py")
    # The further-reading links are the only thing the dialogs can open,
    # and only from a click or a Return on a focused link.
    opens = re.findall(r"webbrowser\.open\(([^)]*)\)", ui)
    assert opens == ["target", "target"]
    assert ui.count('"<Button-1>", lambda _event, target=url: webbrowser.open(target)') == 1


def test_docs_trust_md_is_current():
    committed = (ROOT / "docs" / "TRUST.md").read_text(encoding="utf-8")
    assert committed == trust.render_markdown(), (
        "docs/TRUST.md is stale: run `python -m spec_formatter.trust > docs/TRUST.md`"
    )


# ---------------------------------------------------------------------------
# Absolute claims the copy makes about the code (ledger N10-N12, M3, M5)
# ---------------------------------------------------------------------------

_SHIPPED_PYTHON = [
    path
    for path in [*ROOT.glob("*.py"), *ROOT.joinpath("spec_formatter").rglob("*.py")]
    if path.name != "phase1_smoke_test.py"
]

_NETWORK_IMPORT = re.compile(
    r"^\s*(?:import|from)\s+(urllib\.request|urllib3|httpx|requests|socket|http\.client|"
    r"anthropic|webbrowser|ftplib|smtplib|aiohttp)\b",
    re.MULTILINE,
)

#: The only shipped modules allowed to reach the network or open a browser.
_NETWORK_MODULES = {
    "spec_formatter/updates.py": {"urllib.request"},
    "llm_classifier.py": {"anthropic", "httpx"},
    "spec_formatter/style_application/core/llm_classifier.py": {"anthropic", "httpx"},
    "gui.py": {"webbrowser"},
    "trust_ui.py": {"webbrowser"},
}


def test_network_code_lives_only_in_known_modules():
    found: dict[str, set[str]] = {}
    for path in _SHIPPED_PYTHON:
        modules = set(_NETWORK_IMPORT.findall(path.read_text(encoding="utf-8")))
        if modules:
            found[path.relative_to(ROOT).as_posix()] = modules
    assert found == _NETWORK_MODULES


def test_no_server_socket_or_listener_in_shipped_code():
    for path in _SHIPPED_PYTHON:
        text = path.read_text(encoding="utf-8")
        for needle in ("socketserver", "HTTPServer", ".listen(", "serve_forever"):
            assert needle not in text, (path.name, needle)


def test_gui_opens_links_only_from_click_handlers():
    lines = _source("gui.py").splitlines()
    for number, line in enumerate(lines):
        if "webbrowser.open(" not in line:
            continue
        context = "\n".join(lines[max(0, number - 3): number + 1])
        assert (
            '"<Button-1>"' in context
            or "def _open_releases_page" in context
        ), context
    # _open_releases_page itself is only reached from a link's click binding.
    assert re.search(
        r'link\.bind\("<Button-1>", lambda _event: self\._open_releases_page\(\)\)',
        _source("gui.py"),
    )


def test_no_tools_are_offered_to_the_model():
    for source in (_architect_call_source(), _target_classify_source()):
        assert "tools" not in source and "tool_choice" not in source


def test_no_sampling_parameters_are_sent():
    for source in (_architect_call_source(), _target_classify_source()):
        for parameter in ("temperature", "top_p", "top_k"):
            assert parameter not in source


def test_no_classifier_sets_a_base_url():
    for source in (
        inspect.getsource(architect_classifier),
        inspect.getsource(target_classifier),
    ):
        assert "base_url" not in source


def test_installer_adds_no_autostart():
    script = _source("packaging/windows/installer.iss")
    sections = set(re.findall(r"(?m)^\[(\w+)\]", script))
    assert "Registry" not in sections
    assert "UninstallDelete" not in sections  # the copy says uninstall leaves data
    for needle in ("userstartup", "commonstartup", "schtasks", "DownloadTemporaryFile", "Service"):
        assert needle.lower() not in script.lower(), needle
    run = script.split("[Run]", 1)[1]
    assert "postinstall" in run and "skipifsilent" in run


def test_readme_trust_section_quotes_only_pinned_numbers():
    readme = _source("README.md")
    section = readme.split("## Why trust it?", 1)[1].split("\n## ", 1)[0]
    allowed = _numbers_allowed()
    stray = sorted(
        number
        for number in re.findall(r"\d+(?:[.,]\d+)*", section)
        if number not in allowed
    )
    assert not stray, f"README 'Why trust it?' quotes unpinned numbers: {stray}"
    for key in ("update_disable_env", "anthropic_host", "anthropic_base_url_env", "update_host"):
        assert str(trust.FACTS[key]) in section, key
