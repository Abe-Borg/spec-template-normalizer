# Releasing the Windows desktop app

Specification Formatter ships as a downloadable, self-updating Windows app. There
is no server and nothing to host: the installer and a tiny `latest.json` update
manifest are published as **GitHub Release assets**, and the installed app checks
them directly. This runbook covers cutting a release.

## How updates work (the moving parts)

| Piece | File | Role |
|---|---|---|
| Version literal | `spec_formatter/__init__.py` → `__version__` | The single source of truth. The frozen app reports it; the updater compares it to the manifest. |
| Updater | `spec_formatter/updates.py` | Fetches `latest.json`, compares versions, downloads + SHA-256-verifies the installer, launches it. |
| GUI wiring | `gui.py` | Footer "Check for Updates" button, daily auto-check on launch, the update dialog. |
| Frozen entry | `packaging/windows/app_entry.py` | PyInstaller entry point; `--version` / `--selfcheck` flags for CI. |
| PyInstaller spec | `packaging/windows/specification-formatter.spec` | One-folder build → `dist/SpecificationFormatter/`. |
| Installer | `packaging/windows/installer.iss` | Inno Setup → `dist/installer/SpecificationFormatterSetup.exe`. |
| Manifest maker | `packaging/windows/make_manifest.py` | Computes the installer SHA-256 → `latest.json`. |
| Version guard | `packaging/windows/check_release_version.py` | Fails the build if the tag ≠ `__version__`. |
| Workflow | `.github/workflows/release.yml` | Builds on every relevant PR; builds **and publishes** on a `v*` tag. |

The updater reads `https://github.com/abe-borg/spec-template-normalizer/releases/latest/download/latest.json`.
GitHub only serves the newest **non-prerelease** release at `releases/latest`, so
release candidates never auto-offer themselves to stable installs.

## Cutting a release

1. **Bump the version.** Edit `spec_formatter/__init__.py`:

   ```python
   __version__ = "1.1.0"
   ```

   This is the only version literal — there is no `pyproject.toml`.

2. **Commit** the bump on `master` (or via a merged PR).

3. **Tag and push.** The tag must be `v` + the exact `__version__`:

   ```bash
   git tag v1.1.0
   git push origin v1.1.0
   ```

   For a pre-release, use an `rcN` suffix (e.g. `v1.1.0rc1`) — it publishes as a
   GitHub **pre-release** and is not offered to stable users.

4. **Let CI do the rest.** `release.yml` runs on the tag:
   - guards that the tag matches `__version__` (a half-bumped tag fails loudly
     here rather than shipping an installer stuck in a perpetual "update
     available" loop);
   - builds the one-folder app with PyInstaller and runs the frozen exe's
     `--selfcheck` (catches a missing hidden import before release);
   - compiles the Inno Setup installer;
   - generates `latest.json` (installer SHA-256 + the download URL);
   - the tag-only `publish` job attaches `SpecificationFormatterSetup.exe` and
     `latest.json` to the GitHub Release.

5. **Verify.** Download the installer from the Release page, run it, and confirm
   the app's footer shows the new version.

## Unsigned installer / SmartScreen

The app is **not code-signed** (no paid OS certificate). On first run Windows
SmartScreen shows *"Windows protected your PC"* — choose **More info → Run
anyway**. This is expected and documented for users in the README.

This is separate from the SHA-256 integrity check the updater performs: that is
free and always on. A tampered download can never be launched because its hash
won't match the manifest — the "signing" that actually protects users is the
hash check, not the (skipped) OS certificate.

## Testing the update flow without publishing

Point the app at a hand-made manifest with the `SPEC_FORMATTER_UPDATE_URL`
environment variable:

```json
{
  "schema": 1,
  "version": "9.9.9",
  "url": "https://github.com/abe-borg/spec-template-normalizer/releases/download/v9.9.9/SpecificationFormatterSetup.exe",
  "sha256": "<64 hex chars of a real installer>",
  "notes": "Test update.",
  "published_at": "2026-07-21"
}
```

```powershell
$env:SPEC_FORMATTER_UPDATE_URL = "https://.../your-test-latest.json"
```

Both the manifest URL and the installer `url` must be `https://` — the updater
refuses plaintext (the manifest is the root of trust for the installer hash).

## Environment variables

| Variable | Effect |
|---|---|
| `SPEC_FORMATTER_UPDATE_URL` | Override the manifest URL (testing / a fork's releases). Must be https. |
| `SPEC_FORMATTER_DISABLE_UPDATE_CHECK` | Set truthy to turn off update checks entirely. `0`/`false`/`no`/`off`/empty keep checks on. |
| `SPEC_FORMATTER_SELFCHECK_OUT` | Path the frozen `--selfcheck` writes its result to (used by CI, since the windowed exe has no stdout). |

## Building locally (optional)

On Windows, from the repo root:

```powershell
python -m pip install -r requirements.txt -r requirements-build.txt
pyinstaller packaging/windows/specification-formatter.spec --noconfirm --clean
& "${env:ProgramFiles(x86)}\Inno Setup 6\ISCC.exe" /DMyAppVersion=1.1.0 packaging\windows\installer.iss
```

Outputs land in `dist/SpecificationFormatter/` (the app folder) and
`dist/installer/SpecificationFormatterSetup.exe` (the installer).
