# Unitrans Portal

Windows tooling for preparing UniTrans shipping paperwork. It sorts downloaded
PDF waybills into the same order as the Excel workbook, creates a single bulk
PDF for printing, fills blank driver numbers from scanned waybill images, and
removes permissions-style locks from folders of PDFs.

**Version 0.1.0 (Beta)** — publisher Elvara Studio.

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

## Unlock PDFs

Choose `Unlock PDFs` from the first screen when a folder contains PDFs that open but cannot be edited, printed, or processed because of permissions-style encryption.

1. Click `Choose PDF Folder` and select the folder.
2. Click `Unlock PDFs`.
3. Use `Open Output Folder` to view the new timestamped output folder.

The program creates:

`unlocked_pdfs_YYYYMMDD_HHMMSS`

Inside it:

- `Unlocked PDFs` contains unlocked/copied PDF files.
- `unlock_report.csv` lists every source PDF, output PDF, status, page count, and note.
- `summary.txt` shows the totals.

PDFs that open with a blank password are rewritten without encryption. PDFs that are already unlocked are copied unchanged. PDFs that require an unknown open password are skipped and listed in the report.

## Version history

| Version | Date | Notes |
|---|---|---|
| 0.1.0 | 2026-10-10 | First beta release. Versioning restarted; app renamed to Unitrans Portal. |

Full release notes are in [CHANGELOG.md](CHANGELOG.md).

The app shows its own version in the footer, and Windows Explorer ▸ Properties ▸
Details shows the file version of the installed executable.

## Building a release

Run PowerShell from this folder:

```powershell
.\build.ps1
```

The build:

1. Creates `.venv` and installs `requirements.txt`.
2. Generates `assets/app.ico` from `UniTrans Logo.png` (`make_assets.py`).
3. Rebuilds the exe with the icon and the Windows version metadata.
4. Writes `releases/Unitrans-Portal-<version>.exe` and prints its SHA-256.

The intermediate build output is at `dist\Unitrans Portal.exe`.

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

> **Not implemented yet.** This describes the intended behaviour, so the
> release layout below can be set up correctly ahead of it.

The app checks `releases/latest/download/manifest.json` on GitHub. The manifest
is a small document naming the newest version, where to download it, its size,
and its SHA-256. If a newer version exists, the app shows what changed and can
download and install it in place.

Design decisions already made:

- **Notify first, never mid-job.** The app reports an available update but does
  not interrupt work in progress.
- **Verified before running.** A downloaded build is only used if its SHA-256
  matches the manifest, so a truncated or tampered download is refused.
- **Reversible.** The update keeps a backup of the current executable and
  restores it if the swap fails, so a failed update cannot leave the tool
  unusable.
- **Per-user install.** The intended installer places the app under
  `%LOCALAPPDATA%\Programs`, which needs no administrator rights — so updating
  needs none either.
- **Patch notes are offline-capable.** The release notes for the running version
  travel inside the executable in `changelog.json`, so the What's New dialog
  works without a network connection.

## Repository layout

| Path | Purpose |
|---|---|
| `waybill_sorter.py` | The application. |
| `version.py` | Version and product identity — the single source of truth. |
| `make_assets.py` | Generates the app icon from the logo. |
| `releasectl.py` | Generates the update manifest, publishes releases. |
| `build.ps1` | Builds the executable. |
| `CHANGELOG.md` | Human-readable release notes. |
| `changelog.json` | Release notes as shown in the app (English and Chinese). |
| `releases/` | Published artifacts and `manifest.json`. Binaries here are gitignored and live on the GitHub Release page. |
| `assets/` | Generated icon. |
| `pyinstaller_hooks/` | PyInstaller build hooks. |
