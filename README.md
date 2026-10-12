# Unitrans Portal

Windows tooling for preparing UniTrans shipping paperwork. It sorts downloaded
PDF waybills into the same order as the Excel workbook, creates a single bulk
PDF for printing, fills blank driver numbers from scanned waybill images, and
removes protection from PDFs and Excel workbooks.

**Version 0.4.3 (Beta)** — publisher Elvara Studio.

> **Beta:** the version is `0.x.y`, which means the interface and behaviour can
> still change between releases. Version numbering was restarted at 0.1.0; the
> earlier internal builds were labelled `1.0`, so the version number will appear
> to go *down* when upgrading from one of those. Nothing was removed.

## Use

1. Open `Unitrans Portal.exe`.
2. Choose `Sort Waybill PDFs`.
3. Click `Choose Excel` and select the Excel file.
4. Click `Choose PDF Folder` and select any PDF inside the folder containing the PDF waybills. The app will use that PDF's folder.
5. Confirm the downloaded PDF check matches the Excel supplementary count.
6. Click `Sort Waybills`.
7. Confirm the sort summary shows the copied PDFs were verified as exact copies.
8. Click `Create Bulk PDF`.
9. Open the bulk PDF in your preferred PDF viewer and print from there.

Optional PDF extras are hidden behind `Options` and are off by default:

- Separator pages between waybills.
- Digital sequence marks on forms 1 and 4.

The program creates a new folder inside the PDF folder:

`sorted_waybills_YYYYMMDD_HHMMSS`

Inside it:

- `Sorted` contains copied PDFs named like `001_XAPU1814169.pdf`.
- `Unmatched` contains PDFs that could not be matched.
- `Bulk PDF` contains the combined PDF when you create it.
- `sort_report.csv` shows every match, missing container, duplicate, and unmatched PDF.
- `sort_report.csv` also includes SHA-256 hashes proving each sorted PDF is byte-for-byte identical to the downloaded source PDF.
- `summary.txt` shows the copy amount detected from the Excel name and the final counts.

The original Excel and PDF files are not changed.

The files in `Sorted` are exact copies of the downloaded PDFs. The bulk PDF is a new combined PDF, but when optional extras are off it uses the original PDF pages directly; it does not turn pages into screenshots or change margins.

When separator pages are enabled, the bulk PDF order is: all pages for `001`, separator, all pages for `002`, separator, and so on.

The sequence mark option stamps the Excel sequence number on pages detected as form `1` and form `4`, using the form number printed at the top-left of each waybill page.

## Unlock files

Choose `Unlock Files` from the first screen. Select one or more files, or
choose a folder to process supported files in it and its subfolders. The app
supports PDFs and Excel `.xlsx`, `.xlsm`, `.xltx`, and `.xltm` workbooks.

Enter the password for any file that requires one before clicking `Unlock
Files`. The app can remove Excel workbook-structure and worksheet protection.
Legacy `.xls` files are not supported.

The program creates an `unlocked_files_YYYYMMDD_HHMMSS` folder beside the
selected folder or first selected file. Inside it:

- `Unlocked Files` contains unlocked/copied files.
- `unlock_report.csv` lists each source, output, status, file type, and note.
- `summary.txt` shows the totals.

Unlocked files are written as copies; originals are unchanged. Unprotected files
are copied unchanged. Files that need a missing or incorrect open password are
skipped and listed in the report.

## Version history

| Version | Date | Notes |
|---|---|---|
| 0.4.3 | 2026-10-12 | Seamless waybill sorter language/theme switching with stable control sizes. |
| 0.4.2 | 2026-10-12 | Language and theme changes refresh in place; controls keep fixed sizes. |
| 0.4.1 | 2026-10-12 | Smooth language/theme switching and fixed control sizes; includes PDF and Excel unlocking. |
| 0.4.0 | 2026-10-12 | Unlock PDFs and Excel workbooks; accept file selections or folders; add a Windows installer. |
| 0.3.0 | 2026-10-10 | Dark mode and theme improvements. |
| 0.2.3 | 2026-10-10 | Header spacing and layout improvements. |
| 0.2.2 | 2026-10-10 | Equal-sized buttons and layout improvements. |
| 0.2.1 | 2026-10-10 | Display selection and window placement. |
| 0.2.0 | 2026-10-10 | Settings window with a manual update check. |
| 0.1.0 | 2026-10-10 | First beta release. Versioning restarted; app renamed to Unitrans Portal. |

Full release notes are in [CHANGELOG.md](CHANGELOG.md).

The app shows its own version in the footer and in the Settings window, and
Windows Explorer ▸ Properties ▸ Details shows the file version of the installed
executable.

## Window and display

The app opens as a normal resizable window on the display chosen in
**Settings ▸ Display** (the second screen by default, falling back to the
primary if there is only one).

It is deliberately **not maximised**: the point of opening on a second monitor
is that it is waiting there when you look at it, without covering whatever you
are already working on. Placement uses `SetWindowPos` on the window frame
because Tk's `geometry()` positions the client area and cannot place a window
precisely on a secondary display — `winfo_screenwidth` also reports only the
primary monitor, so Tk alone cannot size a window for any other screen.

Switching language and theme updates the existing interface in place. The app
stays open and controls keep their size and position throughout the change.

## Checking for updates

Open **Settings** (top-right of the window) and choose **Check for updates now**.
The app reads `releases/latest/download/manifest.json` from GitHub, compares the
version, and either confirms you are current or lists what changed with a button
to open the download page.

**The app does not check for updates on its own.** There is no background or
launch-time check in this version — you start it manually. The
`auto_check_updates` setting is stored in `settings.json` so that enabling
automatic checking later needs no migration, but the behaviour is not
implemented yet.

From a terminal, the same check is available without the GUI:

```powershell
.\releases\Unitrans-Portal-0.4.3.exe --check-updates
```

Exit codes: `0` up to date, `2` update available, `1` the check failed. Being
offline or behind a blocked network is reported as a plain message, never a
crash — the check cannot interrupt work in progress.

## Building a release

Run PowerShell from this folder. It creates a portable executable and, when
Inno Setup is installed, a per-user installer in `releases` and Downloads:

```powershell
.\build.ps1
```

The build:

1. Creates `.venv` and installs `requirements.txt`.
2. Generates `assets/app.ico` from `UniTrans Logo.png` (`make_assets.py`).
3. Rebuilds the exe with the icon and the Windows version metadata.
4. Writes `releases/Unitrans-Portal-<version>.exe` and prints its SHA-256.

The intermediate build output is at `dist\Unitrans Portal.exe`.

The installer is named `Unitrans-Portal-<version>-Setup.exe`. It installs to
`%LOCALAPPDATA%\Programs\Unitrans Portal` and creates Start Menu shortcuts.
To build it manually, install Inno Setup and run
`iscc /DAppVersion=<version> installer.iss` from the project root.

Then, to create the update manifest and publish:

```powershell
.\.venv\Scripts\python.exe releasectl.py manifest --version 0.1.0
.\.venv\Scripts\python.exe releasectl.py publish  --version 0.1.0
```

`publish` creates the `v<version>` tag and GitHub release, attaches the exe and
`manifest.json`, and refuses to publish if the artifact's hash no longer matches
the manifest — so a release can never advertise a file it did not build.

### Changing the version

Edit `APP_VERSION` in [version.py](version.py). That is the only place a version
number is written down: the artifact filename, the exe metadata, the in-app
footer, and the manifest all read it from there.

### Python resolution

`build.ps1` looks for an interpreter in this order: `py -3` → `python3` /
`python` (skipping the Windows Store alias stub, which opens the Store instead
of running) → `%LOCALAPPDATA%\Python\bin\python.exe`. Pin one explicitly with
`$env:UNITRANS_PYTHON`.

Releases are currently built with Python 3.14.6. The original 1.0 build used
3.12; re-test after any change of interpreter.

## How updates will work

> **Partly implemented.** The check is live (see "Checking for updates" above).
> Downloading and installing a build from inside the app is not.

The app reads `releases/latest/download/manifest.json` on GitHub. The manifest
is a small document naming the newest version, where to download it, its size,
and its SHA-256. If a newer version exists, the app shows what changed and offers
to open the download page. Replacing the installed executable from inside the app
is the next step and needs a verified download plus a rollback path.

> **Two constraints the updater depends on — do not break these:**
>
> 1. **The repository must be public.** An unauthenticated `GET` is the only way
>    the app can read the manifest without shipping a credential inside the
>    executable. Anyone can extract a string from a distributed `.exe`, so an
>    embedded token is not a secret; it is a liability. A private repository
>    returns `404` to the app.
> 2. **Releases must not be GitHub pre-releases.** GitHub excludes pre-releases
>    from `/releases/latest`, so a beta published that way is invisible to the
>    updater. Beta status is carried by the version number (`0.2.0 (Beta)`) and
>    the manifest's `beta` field, not by a GitHub release channel.
>    `releasectl.py publish` therefore does not set `--prerelease` automatically;
>    the flag exists but warns when used.

Design decisions already made:

- **Notify only, never mid-job.** The app reports an available update and opens
  the download page; it does not interrupt work in progress.
- **Verified before running.** A future downloaded build will only be used if its
  SHA-256 matches the manifest, so a truncated or tampered download is refused.
- **Reversible.** The planned update will keep a backup of the current executable
  and restore it if the swap fails, so a failed update cannot leave the tool
  unusable.
- **Per-user install.** The intended installer places the app under
  `%LOCALAPPDATA%\Programs`, which needs no administrator rights — so updating
  needs none either.
- **Patch notes are offline-capable.** The release notes for the running version
  travel inside the executable in `changelog.json`, so the What's New content
  works without a network connection.

## Repository layout

| Path | Purpose |
|---|---|
| `waybill_sorter.py` | The application and its interface. |
| `version.py` | Version and product identity — the single source of truth. |
| `updater.py` | Update checking: fetches and validates the release manifest. |
| `settings.py` | Per-user settings in `%APPDATA%\UnitransPortal`. |
| `make_assets.py` | Generates the app icon from the logo. |
| `releasectl.py` | Generates the update manifest, publishes releases. |
| `build.ps1` | Builds the executable. |
| `CHANGELOG.md` | Human-readable release notes. |
| `changelog.json` | Release notes as shown in the app (English and Chinese). |
| `tests/` | Test suite, plus GUI and packaged-exe smoke tests. |
| `tools/` | One-off maintenance scripts. |
| `releases/` | Published artifacts and `manifest.json`. Binaries here are gitignored and live on the GitHub Release page. |
| `assets/` | Generated icon. |
| `pyinstaller_hooks/` | PyInstaller build hooks. |

## Tests

```powershell
.\.venv\Scripts\python.exe -m unittest discover -s tests
```

Covers version comparison, i18n key parity, update checking against a fake
transport (offline, timeouts, 404s, malformed manifests), settings persistence,
and source-health checks. The GUI and the packaged executable are verified by
running them:

```powershell
.\.venv\Scripts\python.exe tests\smoke_gui.py --settings --screenshot settings.png
```
