# Changelog

All notable changes to Unitrans Portal are recorded here.

This file is the human-readable source of truth for release notes.
`changelog.json` holds the same content in the structured form the app displays
in its **What's New** dialog — update both when cutting a release.

The format follows [Keep a Changelog](https://keepachangelog.com/en/1.1.0/),
and the project uses [Semantic Versioning](https://semver.org/spec/v2.0.0.html).
While the version is `0.x.y` the release is **beta**: the interface and
behaviour can still change between versions.

## [Unreleased]

### Planned

- One-click in-app updating: download the new build, verify its SHA-256, and
  replace the installed executable with automatic rollback if the swap fails.
  Today the app tells you an update exists and opens the download page.
- Automatic update checks on launch. The setting is stored but the feature is
  not implemented; checking is manual.
- Inno Setup installer, per-user, so no administrator rights are needed.
- `progress_value` and `progress_text` are written but never displayed - there
  is no progress bar in the current UI. Either wire one up or remove them.

## [0.2.0] - 2026-10-10

### Added

- **Settings window**, reached from the new `Settings` button in the header. It
  shows the installed version and provides `Check for updates now`.
- **Manual update check.** The app reads the published release manifest, compares
  the version numerically, and either confirms you are current or shows what
  changed with a button to open the download page. Being offline, behind a
  blocked network, or receiving a malformed manifest all report a plain message
  rather than failing - a check never interrupts work.
- Release notes are shown in the user's chosen language, falling back to English.
- `--check-updates` command-line flag, which performs the same check with no GUI
  and returns an exit code: 0 up to date, 2 update available, 1 check failed.
- Settings are stored per-user in `%APPDATA%\UnitransPortal\settings.json`,
  written atomically, so a preference survives an update and a corrupt file
  degrades to defaults instead of stopping the app.

### Changed

- **Version numbering corrected to 0.2.0.** The previous release was published
  as 0.1.0 but contains no update checking at all, so 0.1.1 would misrepresent
  it: this adds a feature, and it is the release that makes the existing
  `0.1.0` installs able to see updates.

### Fixed

- Removed five methods that were defined twice in the application class
  (`append_log`, `clear_log`, `toggle_language`, `set_busy`,
  `update_action_state`) plus an orphaned `localize_log`. Python keeps only the
  last definition of a method, so the earlier copies were unreachable dead code
  left behind by an older generation of the UI. No behaviour changed - the
  removed code could not run - but editing it would have had no effect, which
  is a trap for whoever works on this next. A test now fails if it recurs.
- `save_settings` no longer raises on a malformed path; it reports failure.

## [0.1.0] - 2026-10-10

First beta release.

### Changed

- **Version numbering restarts at 0.1.0.** Earlier internal builds were
  labelled `1.0`. Nothing was removed — the beta simply gets an honest number,
  so expect the version to appear to go *down* when upgrading from a 1.0 build.
- **The application is now called Unitrans Portal** (previously "UniTrans
  Waybill Sorter").
- The version shown in the footer is now `Version 0.1.0 (Beta)`, generated from
  a single source of truth (`version.py`).

### Added

- A real application icon. The source logo is a wide 1.9:1 export carrying a
  transparent margin and a tagline; the build now trims it to the wordmark and
  emits a proper multi-resolution icon (16–256 px), so it is no longer the
  generic Windows placeholder.
- Windows version metadata (`FileVersion`, `ProductName`, `CompanyName`,
  copyright) so Explorer ▸ Properties identifies the build instead of showing
  blank fields.
- Release artifacts are named with their version
  (`Unitrans-Portal-0.1.0.exe`), so two builds can never share a filename.
- `version.py` — one place that defines the version, with numeric semantic
  version comparison covered by tests.
- `changelog.json` — release notes in English and Chinese, bundled into the
  app and ready for the What's New dialog.
- `make_assets.py` — regenerates the icon from the logo, run automatically by
  the build.

### Fixed

- Removed a hardcoded logo path pointing at a different developer's machine
  (`C:/Users/ChrisMicheal/...`), which could never resolve for anyone else.

### Existing features in this release

- Sort waybill PDFs into Excel order, with SHA-256 verification that each
  sorted file is a byte-for-byte copy of the original download.
- Create a single bulk PDF for printing, with optional separator pages and
  digital sequence marks on forms 1 and 4.
- Fill blank driver numbers from scanned waybill images.
- Unlock PDFs that open but are blocked from editing or printing by
  permissions-style encryption.
- English and Chinese interface.

[Unreleased]: https://github.com/ElvaraStudio/unitrans-portal/compare/v0.2.0...HEAD
[0.2.0]: https://github.com/ElvaraStudio/unitrans-portal/releases/tag/v0.2.0
[0.1.0]: https://github.com/ElvaraStudio/unitrans-portal/releases/tag/v0.1.0
