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

- In-app update checking and one-click updating from GitHub Releases.
- Inno Setup installer, per-user, so no administrator rights are needed.
- Chinese translation review of the new release-notes text.

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

[Unreleased]: https://github.com/ElvaraStudio/unitrans-portal/compare/v0.1.0...HEAD
[0.1.0]: https://github.com/ElvaraStudio/unitrans-portal/releases/tag/v0.1.0
