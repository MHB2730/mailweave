# Releasing MailWeave

A new MailWeave release is a single version bump plus an installer attached to a GitHub release. The running app auto-detects new releases via the GitHub Releases API and offers to download + install.

## One-time setup (already done)

- `version.py` holds `__version__` and `GITHUB_REPO = 'MHB2730/mailweave'`.
- `sync_version.py` regenerates `version_info.txt` (exe metadata) and `_version.iss` (installer) from `version.py`. Called automatically by `build.bat`.
- `installer.iss` produces `installer_output/MailWeave-Setup-<version>.exe` — the filename pattern the in-app updater looks for.
- `gh` (GitHub CLI) must be installed and authenticated (`gh auth login`) for the `gh release create` step. Alternative: upload the installer through the GitHub web UI.

## Cut a release

1. **Bump the version.** Edit `version.py` and change `__version__` (semantic versioning, e.g. `1.0.0` → `1.0.1` for a bug-fix release, `1.1.0` for new features).
2. **Build.** Run `build.bat`. This will:
   - Sync `version_info.txt` + `_version.iss` from `version.py`.
   - Produce `dist\MailWeave\MailWeave.exe`.
   - Produce `installer_output\MailWeave-Setup-<version>.exe`.
3. **Smoke-test the installer** on a clean directory before publishing. Install it, launch the app, confirm the About dialog shows the new version, do at least one bundle build.
4. **Commit + tag.**
   ```
   git add version.py
   git commit -m "release: v<version>"
   git tag v<version>
   git push origin main --tags
   ```
5. **Publish the GitHub release** with the installer attached:
   ```
   gh release create v<version> \
       installer_output/MailWeave-Setup-<version>.exe \
       --title "MailWeave v<version>" \
       --notes "Short summary of what's in this release."
   ```
   Or via the web UI: https://github.com/MHB2730/mailweave/releases/new — pick the tag, attach the installer, publish.

## Verifying the auto-update path

After publishing, on a machine running an older version of MailWeave:

- Open **Help → Check for Updates…**. The dialog should report the new version, show the installer name + size, and offer **Download & install**.
- Clicking that downloads the `.exe` to `%TEMP%`, launches it (UAC prompt — admin install), and exits the app so the installer can replace files.

If a release shows up but the updater says "no installer asset", you forgot to attach the `.exe` to the release. The asset filename must match `MailWeave-Setup-*.exe`.

## Hotfixes

For an urgent fix, bump the patch component (`1.0.0` → `1.0.1`), repeat the standard release flow. The auto-updater compares numeric version tuples so `1.0.10` is correctly treated as newer than `1.0.9`.

## A note on code signing

The installer is currently **unsigned**, so Windows SmartScreen will warn users on first install and on every auto-update download. To remove that:

1. Acquire an EV or OV code-signing certificate.
2. Set `MAILWEAVE_SIGN_THUMBPRINT` to the cert's SHA-1 thumbprint.
3. `build.bat` already calls `sign.ps1` against the exe and the installer if the cert is found.

Until then, the auto-update flow still works — users just need to click "Run anyway" once per update.
