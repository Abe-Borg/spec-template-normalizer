# Trust claims ledger

This ledger backs the two trust surfaces: **Why Trust It?** (Level 1) and the
**"I'm not convinced"** dossier (Level 2). Both are written in
`spec_formatter/trust.py` and drawn by `trust_ui.py`. `docs/TRUST.md` is the
same text rendered as Markdown.

Every factual claim in either level is listed here with the code that makes
it true and how it is checked. Each number, model ID, limit, host name, path
and setting that the copy quotes is a key in `spec_formatter.trust.FACTS`.
`tests/test_trust_surface.py::test_fact_matches_code[<key>]` asserts that key
equals the constant, or the literal, in the source named below. A change to
that source therefore fails a test until the copy is updated.
`test_every_number_in_the_copy_is_pinned` also fails if a digit appears in the
copy without coming from `FACTS`.

"Manual" means a reviewer read the code path and confirmed it. A test cannot
hold that claim, so re-read the source whenever the named code changes.

The rule (also in `CLAUDE.md`): **if the implementation changes, the trust
surfaces and this ledger change in the same change.** A new user action, a
new network call, a new automatic behavior, a model or setting change, or a
new limit each require an update.

## Network destinations

| ID | Claim | Source | Verified by |
|---|---|---|---|
| N1 | The app's model requests go to `api.anthropic.com` by default (`anthropic_host`). | `anthropic` SDK 0.84.0 `Anthropic.__init__` base URL default; neither classifier passes `base_url` | `test_fact_matches_code[anthropic_host]` (reads the installed SDK source); `test_no_classifier_sets_a_base_url` |
| N2 | The SDK honours the `ANTHROPIC_BASE_URL` environment variable, which redirects those requests (`anthropic_base_url_env`). | `anthropic/_client.py` `os.environ.get("ANTHROPIC_BASE_URL")` | `test_fact_matches_code[anthropic_base_url_env]` |
| N3 | The model is contacted only during a run, and only when a template must be analyzed fresh or a target has paragraphs no rule resolved. | `pipeline.prepare_template_profile` (cache hit returns before any client is built); `batch_runner.process_single_file` (`classification` stage); `core/llm_classifier.classify_target_document` (returns before `import anthropic` when nothing is unresolved) | `tests/test_architect_free_modes.py` (no API key); manual |
| N4 | Each model request is preceded by a token-count request to the same host carrying the same content: one per architect analysis, and one per Haiku target attempt. Anthropic documents token counting as free. | `llm_classifier._count_input_tokens`; `core/llm_classifier._check_haiku_prompt_size`; platform.claude.com/docs/en/build-with-claude/token-counting ("Token counting is free to use"), checked 2026-10-22 | manual |
| N5 | The automatic update check runs `update_check_delay_s` s after the window opens, at most once per `update_check_interval_days` day, unless `SPEC_FORMATTER_DISABLE_UPDATE_CHECK` is set (`update_disable_env`). | `gui.AUTO_UPDATE_CHECK_DELAY_MS`, `gui.App._maybe_auto_check_for_updates`; `updates.DEFAULT_MIN_INTERVAL_DAYS`, `updates.should_auto_check`, `updates.update_check_disabled`, `updates.ENV_DISABLE` | `test_fact_matches_code[update_check_delay_s]`, `[update_check_interval_days]`, `[update_disable_env]`; existing `tests/test_updates.py` |
| N6 | The update check is one HTTPS GET of `update_manifest_url` (host `update_host`) with the user agent `update_user_agent`. It sends no version number, no document data and no key; the version comparison happens on your machine. | `updates.fetch_manifest`, `updates._DEFAULT_MANIFEST_URL`, `updates._USER_AGENT`, `updates.check_for_update` | `test_fact_matches_code[update_manifest_url]`, `[update_host]`, `[update_user_agent]`; manual |
| N7 | The manifest URL can be overridden with `SPEC_FORMATTER_UPDATE_URL` (`update_url_env`). Plain `http` is refused, both as requested and after redirects. | `updates.manifest_url`, `updates.fetch_manifest`, `updates._require_https_final_url` | `test_fact_matches_code[update_url_env]`; `tests/test_updates.py` |
| N8 | The manifest read is capped at `update_manifest_max_kib` KiB, with a timeout of `update_manifest_timeout_s` s. | `updates.MAX_MANIFEST_BYTES`, `updates.DEFAULT_MANIFEST_TIMEOUT` | `test_fact_matches_code[update_manifest_max_kib]`, `[update_manifest_timeout_s]` |
| N9 | An installer is downloaded only after you click Download & Install. It must be HTTPS, is capped at `installer_max_mib` MiB, uses a `installer_timeout_s` s socket timeout, and is SHA-256 checked before it can run. | `gui.App._start_update_download`; `updates.download_installer`, `updates.MAX_INSTALLER_BYTES`, `updates.DEFAULT_DOWNLOAD_TIMEOUT`, `updates.verify_sha256` | `test_fact_matches_code[installer_max_mib]`, `[installer_timeout_s]`; `tests/test_updates.py` |
| N10 | Links in the help, About, key-help and update dialogs open in your own browser or mail program only when clicked: the key console (`console_keys_url`), the releases page (`releases_url`), the author's email and LinkedIn. | `gui.py` `webbrowser.open` call sites (all inside click handlers) | `test_gui_opens_links_only_from_click_handlers`; `test_fact_matches_code[console_keys_url]`, `[releases_url]` |
| N11 | The app has no other network code: no telemetry, crash reporting, analytics, fonts or scripts fetched from the internet. | repository-wide search for `urllib`, `httpx`, `socket`, `requests`, `http.client` and `anthropic` imports | `test_network_code_lives_only_in_known_modules` |
| N13 | Both the Anthropic library (httpx, environment and system proxy settings) and the update check (`urllib`, system proxy settings) send their requests through a configured proxy. | httpx `trust_env` default; `urllib.request.urlopen`; `updates.fetch_manifest` docstring | manual |
| N12 | The Windows installer makes no network connections and adds no startup entry, service or scheduled task. | `packaging/windows/installer.iss` (`[Run]` launches the app only post-install, `skipifsilent`; no `[Registry]` section) | `test_installer_adds_no_autostart` |

## AI models and settings

| ID | Claim | Source | Verified by |
|---|---|---|---|
| M1 | Template analysis uses `architect_model` (`architect_model_label`). | `phase1_pipeline.DEFAULT_MODEL`; `gui.FormatWorker.run` passes no `template_model` | `test_fact_matches_code[architect_model]` |
| M2 | Target classification uses `target_model_default` (`target_model_default_label`) by default; `target_model_alt` (`target_model_alt_label`) can be chosen in the **Target AI model** menu. | `model_config.DEFAULT_TARGET_MODEL`, `model_config.TARGET_MODEL_CHOICES` | `test_fact_matches_code[target_model_default]`, `[target_model_alt]`, `[target_model_default_label]`, `[target_model_alt_label]` |
| M3 | Both use adaptive thinking (`thinking_type`) at effort `effort` and a maximum response of `max_output_tokens` tokens. No temperature is set. | `llm_classifier._call_api`; `core/llm_classifier.classify_target_document` request dict; `pipeline.format_specifications` `target_effort` default | `test_fact_matches_code[effort]`, `[max_output_tokens]`, `[thinking_type]`; `test_no_sampling_parameters_are_sent` |
| M4 | The target effort can be overridden with `SPEC_FORMATTER_TARGET_EFFORT` (`target_effort_env`). The desktop app does not set it. | `pipeline.format_specifications` | `test_fact_matches_code[target_effort_env]` |
| M5 | No tools are offered to either model. | both classifiers' request construction (no `tools` key) | `test_no_tools_are_offered_to_the_model` |
| M6 | Answers are forced into a JSON schema. The target model may answer only a paragraph index plus one of the template's role names, or the single ignore reason `target_ignore_reason`. | `core/llm_classifier._classification_output_config` | `test_fact_matches_code[target_ignore_reason]` |
| M7 | Target answers for paragraphs that were not sent, duplicates, and unknown roles are rejected. A model answer cannot override a rule's answer; trying fails the target (`classification_deterministic_override`). | `core/llm_classifier._validate_classifications`; `core/classification.coerce_to_final_classifications` | existing `tests/style_application_regression` coverage |
| M8 | The template model chooses roles, exemplar paragraphs, ignored paragraphs, the display names of the generated `CSI_` styles, and free-text notes. It never writes formatting values; formatting is copied from the exemplar paragraph, and `basedOn` is repaired from the source. | `llm_classifier._instruction_response_schema`; `docx_decomposer.validate_instructions` (formatting fields forbidden); `docx_decomposer.build_portable_styles_xml`; `llm_classifier._repair_style_based_on` | manual |
| M9 | Connect timeout `connect_timeout_s` s; read timeout `read_timeout_s` s (`read_timeout_min` minutes). SDK retries are off. | `httpx.Timeout(600.0, connect=5.0)`, `max_retries=0` in both classifiers | `test_fact_matches_code[connect_timeout_s]`, `[read_timeout_s]`, `[read_timeout_min]` |
| M10 | Target transport retries: `target_transport_retries` after the first try, for rate limits, connection errors and 5xx only. A rate-limit wait honours `Retry-After` up to `retry_after_cap_s` s; other waits are 2 s and then 4 s (`backoff_first_s`, `backoff_second_s`). | `core/llm_classifier._TRANSPORT_RETRIES`, `_MAX_RETRY_AFTER_SECONDS`, `_transport_retry_delay` | `test_fact_matches_code[target_transport_retries]`, `[retry_after_cap_s]`, `[backoff_first_s]`, `[backoff_second_s]` |
| M11 | Target unusable-answer regenerations: `target_regenerations` (so `target_total_attempts` attempts in all). | `classify_target_document` `max_regenerations` | `test_fact_matches_code[target_regenerations]` |
| M12 | Architect transport attempts: `architect_transport_attempts` in all; unusable-JSON attempts: `architect_response_attempts`; coverage follow-up requests: up to `architect_patch_attempts`. One extra retry without the schema if Anthropic cannot compile it. | `llm_classifier._call_api` (`while attempt < 3`), `DEFAULT_RESPONSE_ATTEMPTS`, `classify_document(max_patch_attempts=3)` | `test_fact_matches_code[architect_transport_attempts]`, `[architect_response_attempts]`, `[architect_patch_attempts]` |
| M13 | Large targets are split into chunks of about `chunk_tokens` estimated tokens or `chunk_paragraphs` paragraphs, overlapping by `chunk_overlap` paragraphs, with at most `chunk_workers` chunks at once. Chunks that disagree in an overlap are re-asked once. | `core/llm_classifier._MAX_BUNDLE_TOKENS`, `_split_bundle_into_chunks`, `_CHUNK_OVERLAP`, `classify_target_document` | `test_fact_matches_code[chunk_tokens]`, `[chunk_paragraphs]`, `[chunk_overlap]`, `[chunk_workers]` |
| M14 | At most `max_concurrent_requests` requests are in flight across the whole run (`max_concurrent_requests_env`, capped at `max_concurrent_requests_ceiling`). | `core/llm_classifier._DEFAULT_MAX_CONCURRENT_REQUESTS`, `_MAX_CONCURRENT_REQUESTS_CEILING`, `_MAX_CONCURRENT_REQUESTS_ENV` | `test_fact_matches_code[max_concurrent_requests]`, `[max_concurrent_requests_ceiling]`, `[max_concurrent_requests_env]` |
| M15 | A refusal is final and is reported as `classification_refused`; it is not retried. | both classifiers (`ClassificationRefused`) | existing `tests/test_llm_classifier_safety.py` |

## What is sent

| ID | Claim | Source | Verified by |
|---|---|---|---|
| S1 | Template analysis sends every paragraph of the template: up to `architect_text_chars` characters of its text, its style name, numbering and formatting hints, plus the style and numbering catalogs and the two prompt files. | `docx_decomposer.build_slim_bundle`; `llm_classifier.classify_document` (`user_message`) | `test_fact_matches_code[architect_text_chars]` |
| S2 | Target classification sends only the paragraphs no rule resolved: up to `target_text_chars` characters each, up to `neighbour_text_chars` characters of the paragraphs before and after, style ID, numbering and formatting hints, plus the role list and the shipped prompt. Table paragraphs are never sent. | `core/classification.build_phase2_slim_bundle`; `core/llm_classifier._project_request_bundle`, `_system_blocks` | `test_fact_matches_code[target_text_chars]`, `[neighbour_text_chars]` |
| S3 | Whole files, images, headers, footers, comments and file paths are not sent. | the two bundle builders above (paragraph rows only); manual |
| S4 | The key travels only in the request header to Anthropic. | SDK `auth_headers` (`X-Api-Key`) | manual |
| S5 | The SDK adds standard headers: SDK version, operating system, processor architecture and Python version. | `anthropic/_base_client.platform_headers` | manual |

## Bounds and failure behaviour

| ID | Claim | Source | Verified by |
|---|---|---|---|
| B1 | Template analysis is refused before any request above `architect_input_cap` input tokens. | `llm_classifier.MAX_SINGLE_PASS_INPUT_TOKENS`, `_check_input_cost_guard` | `test_fact_matches_code[architect_input_cap]` |
| B2 | A target with more than `target_paragraph_cap` unresolved paragraphs fails before any request (`target_too_large`). The cap can be changed with `SPEC_FORMATTER_MAX_TARGET_PARAGRAPHS` (`target_paragraph_cap_env`). | `batch_runner.DEFAULT_MAX_TARGET_PARAGRAPHS`, `_MAX_TARGET_PARAGRAPHS_ENV` | `test_fact_matches_code[target_paragraph_cap]`, `[target_paragraph_cap_env]` |
| B3 | Every Haiku request is checked against `haiku_prompt_budget` estimated tokens first. The code's stated reason is a price step at `haiku_price_step` tokens. A request over the budget fails (`classification_prompt_too_large`). | `core/llm_classifier._HAIKU_MAX_PROMPT_TOKENS`, `_check_haiku_prompt_size` | `test_fact_matches_code[haiku_prompt_budget]`, `[haiku_price_step]` |
| B4 | DOCX limits: `package_entries` entries, `package_total_mib` MiB uncompressed, `package_part_mib` MiB per part, compression ratio `compression_ratio`:1. Header/footer images: `hf_media_mib` MiB each, `hf_media_total_mib` MiB in all. | `style_application/docx_decomposer.MAX_*`; `core/registry.MAX_HEADER_FOOTER_MEDIA_*` | `test_fact_matches_code[...]` for each key |
| B5 | The desktop app formats `gui_workers` targets at once. The Python API allows up to `api_max_workers`. | `gui.DEFAULT_MAX_WORKERS`; `pipeline._MAX_WORKERS` | `test_fact_matches_code[gui_workers]`, `[api_max_workers]` |
| B6 | A failed target never publishes a DOCX; other targets keep their outputs. The run folder, `run.log`, `diagnostics.jsonl`, `run.json` and an audit for every target are still written. | `pipeline._format_one_target`, `_write_run_artifacts`, `_write_initialization_failure_artifacts` | existing `tests/test_unified_pipeline.py` |
| B7 | Cancel starts no new target or request, wakes retry waits, and closes an open response at its next streamed event. Outputs already published stay; nothing else is published. | `cancellation.check_cancelled`, `wait_for_retry`, `request_slot`; `llm_usage.streamed_stop_details`; `pipeline._publish_output` | existing `tests/test_cancellation.py` |
| B8 | Tokens from requests that finished before a failure or Cancel are counted and billed. A response cut off by Cancel is recorded as *unknown*, not zero. | `llm_usage.UsageCollector`, `attach_usage`, `usage_from_exception` | existing `tests/test_llm_usage.py` |

## Data at rest

| ID | Claim | Source | Verified by |
|---|---|---|---|
| D1 | Your template and targets are copied to a private snapshot and checked by size, modification time and SHA-256. Each target is checked again before its output is published. | `pipeline._snapshot_input`, `_stable_source_sha256`, `_format_one_target` | existing round-trip tests (source immutability) |
| D2 | Outputs go into a new `<UTC time>_<mode>_<id>` folder under the output folder you choose. Nothing in an older run folder is changed. | `pipeline._create_run_directory` | existing `tests/test_unified_pipeline.py` |
| D3 | Template analyses are cached in `cache_dir_windows`. Each holds the template's styles and settings parts, its header/footer parts, and the first `architect_text_chars` characters of every template paragraph. | `pipeline.default_template_cache_dir`; `phase1_bundle.build_classification_audit` (`text`); `arch_env_extractor` | `test_fact_matches_code[cache_dir_windows]`; manual |
| D4 | After a fresh analysis, only the newest `cache_keep` profiles of that template are kept. Profiles of other templates are never removed automatically. | `pipeline._prune_stale_profiles`, `_PROFILE_CACHE_KEEP` | `test_fact_matches_code[cache_keep]` |
| D5 | The update check stores its last-check time and any skipped version in `update_state_file` under `config_dir_windows`. Downloaded installers stay in its `updates` subfolder (`updates_subdir`). | `updates.STATE_FILENAME`, `app_paths.default_config_dir`, `gui.App._update_download_worker` | `test_fact_matches_code[update_state_file]`, `[config_dir_windows]`, `[updates_subdir]` |
| D6 | With **Remember** ticked, the key is stored in Windows Credential Manager under the service `keyring_service` and user name `keyring_user`. Unticking deletes it. Otherwise the key is held only in memory. | `secrets.KEYRING_SERVICE`, `KEYRING_USERNAME`, `save_api_key`, `clear_api_key`; `gui.App._on_remember_key_toggled` | `test_fact_matches_code[keyring_service]`, `[keyring_user]` |
| D7 | `ANTHROPIC_API_KEY` (`api_key_env`), when set, pre-fills the key field and wins over a remembered key. | `gui.App.__init__` | `test_fact_matches_code[api_key_env]` |
| D8 | Working copies live in short temporary folders that are deleted when the run ends. A hard kill can leave them behind. | `pipeline.format_specifications` (`TemporaryDirectory(prefix="sf-")`), `batch_runner.process_single_file` (`phase2_`), `phase1_pipeline.run_phase1` (`.phase1-work-`) | manual |
| D9 | Uninstalling leaves the cache, the update folder and any remembered key in place. | `installer.iss` (no `[UninstallDelete]`) | `test_installer_adds_no_autostart` (also asserts no `[UninstallDelete]`) |

## Records and redaction

| ID | Claim | Source | Verified by |
|---|---|---|---|
| R1 | `run.json`, every `audit.json`, `run.log` and `diagnostics.jsonl` never contain the API key; it is replaced wherever it could appear. | `pipeline._redact`, `_redact_json`, `_sanitize_run_text` | existing `tests/test_unified_pipeline.py` redaction tests |
| R2 | They hold no paragraph text: audit strings are restricted to known identifiers, diagnostics drop any value with spaces or a text-like key, and log lines must match known safe shapes. They do hold file names and paths. | `pipeline._normalize_audit_details`, `_sanitize_run_text`; `diagnostics.sanitize_fields` | existing `tests/test_diagnostics.py`, `tests/test_error_location.py` |
| R3 | Each audit lists every paragraph's final role or ignore reason, but not whether a rule or the AI decided it; `run.log` gives the counts (`unresolved` = sent to the AI, `deterministic` = decided by rules). | `batch_runner._classification_audit`, `_safe_application_audit`; `process_single_file` log line | manual (see finding F3) |
| R4 | `run.json` records model IDs, effort, prompt fingerprints, source and output SHA-256, and observed token counts (`diagnostics.usage`). | `pipeline._write_run_artifacts` | existing tests |

## Verification gates (deterministic)

| ID | Claim | Source | Verified by |
|---|---|---|---|
| V1 | Every visible target paragraph gets exactly one disposition (styled or ignored), or the target fails. | `core/classification` coverage checks; `batch_runner._apply_classified_target_impl` (`disposition_verification`) | existing tests |
| V2 | Format only: body text is compared run by run (tabs, breaks, spaces, hyphens, tracked deletions) and effective numbering is compared before and after. | `core/xml_helpers.paragraph_run_content_signature`; `classification._verify_format_only_body_invariants` | existing tests |
| V3 | Conversions: each converted paragraph is held to a prediction made before editing; every other paragraph must be unchanged. | `core/expected_changes`; `phase2_invariants.verify_phase2_invariants` | `tests/test_final_gate_text_identity.py` |
| V4 | Every package part outside the mode's remit must be byte-identical. | `phase2_invariants._package_member_census`, `_package_remit` | `tests/test_package_change_whitelist.py` |
| V5 | Tracked revisions are counted by author and kind before and after. | `core/revisions.revision_census` | `tests/test_revision_accounting.py` |
| V6 | Paragraph indentation is checked under both readings of Word's precedence rules. | `phase2_invariants` geometry check | `tests/test_geometry_invariant.py` |
| V7 | Every restyled paragraph's style is read back after the edit. | `core/classification` (`paragraph_style_not_applied`) | existing tests |
| V8 | Template profiles are reused only if source hash, engine fingerprint, model, prompts and contract version all match. | `pipeline._manifest_matches_current_engine`, `_find_cached_profile` | existing `tests/test_unified_pipeline.py` |
| V9 | `DOCTYPE`/`ENTITY` declarations are rejected before parsing; external relationship targets are never fetched. | `core/untrusted_xml.parse_untrusted_xml`; `core/opc_paths` | existing adversarial tests |

## Third-party policy claims (not code)

| ID | Claim | Source | Verified by |
|---|---|---|---|
| T1 | Anthropic's Commercial Terms say: "Anthropic may not train models on Customer Content from Services." | https://www.anthropic.com/legal/commercial-terms, read 2026-10-22 | manual; re-check when quoting |
| T2 | Anthropic documents token counting as free to use. | https://platform.claude.com/docs/en/build-with-claude/token-counting, read 2026-10-22 | manual |

## Findings (behaviour that limits what the copy may claim)

None of these is changed by this work. Each one is disclosed in the copy as it is.

| ID | Finding | Where | Disclosed in |
|---|---|---|---|
| F1 | The template cache keeps the first 200 characters of every template paragraph in plain text (`classification_audit.json`) with no expiry. Profiles of other templates are never pruned, and uninstalling does not remove them. | `phase1_bundle.build_classification_audit`; `_prune_stale_profiles`; `installer.iss` | Dossier: data at rest, Honesty |
| F2 | Downloaded installers are never deleted from `%LOCALAPPDATA%\SpecificationFormatter\updates`. | `updates.download_installer` | Dossier: data at rest |
| F3 | Audits do not record, per paragraph, whether a rule or the AI chose the role; only counts are logged. An AI decision therefore cannot be traced to its paragraph after the run. | `batch_runner._classification_audit` | Short answer, Provenance, Honesty |
| F4 | A model-authored string, the display name of each generated `CSI_` style, is written into the output's `styles.xml` in the two template modes. | `docx_decomposer.build_style_xml_block` | Provenance |
| F5 | The earlier Level 1 copy said a run "leaves nothing partial behind". Successful targets of a partly failed run *are* published (by design), and a hard kill can leave a `.staging` folder. Reworded. | `pipeline._publish_output`, `_remove_staging_dir` | Level 1 |
| F6 | `SPEC_FORMATTER_UPDATE_URL` redirects the update check, and `ANTHROPIC_BASE_URL` (honoured by the SDK) redirects all model traffic, document excerpts and key included, to whatever host it names. Neither is shown in the UI. | `updates.manifest_url`; Anthropic SDK | Boundary, Security |
| F7 | The app has a dark theme only, and Tk exposes no accessibility roles to screen readers; the trust dialogs inherit both limits. | `gui.main` (`set_appearance_mode("dark")`) | report only |
