# Releases

Published artifacts for Unitrans Portal. The executables here are **not**
stored in git — they are attached to the corresponding
[GitHub Release](https://github.com/ElvaraStudio/unitrans-portal/releases) and
are gitignored locally.

Tracked in this folder:

| File | Purpose |
|---|---|
| `README.md` | This file. |
| `manifest.json` | Describes the newest published build: version, download URL, size and SHA-256. This is what the app's update check reads. |

Everything else here is a build artifact and can be deleted safely; run
`build.ps1` to regenerate.

## Current release

| Version | Artifact | Built | Size | SHA-256 |
|---|---|---|---|---|
| 0.2.0 | `Unitrans-Portal-0.2.0.exe` | 2026-10-10 | 23.3 MB | _(see `manifest.json`)_ |

The hash of record is always the one in `manifest.json`, so there is only one
place to check. `releasectl.py` refuses to publish if the artifact has changed
since the manifest was generated.

## Version history

| Version | Date | Notes |
|---|---|---|
| 0.2.0 | 2026-10-10 | Settings window with a manual update check. Version numbering corrected from 0.1.1: this release adds a feature, and it is what allows the 0.1.0 installs to see updates. |
| 0.1.0 | 2026-10-10 | First beta. Versioning restarted from the old `1.0` label; app renamed to Unitrans Portal. |

## Retired builds

Three pre-0.2.0 executables have been removed:

| File | Built | Why it was removed |
|---|---|---|
| `UniTrans Waybill Sorter 1.0.exe` | 2026-06-15 08:43 | Predated the current source by an hour; the source revision it came from no longer exists and it cannot be rebuilt. |
| `UniTrans Waybill Sorter.exe` | 2026-06-15 09:42 | Superseded by the numbered releases. Kept no version metadata, so it could not be identified from Explorer. |
| `Unitrans-Portal-0.1.0.exe` | 2026-10-10 | Superseded by 0.2.0. Still downloadable from its release page. |

The old `UniTrans-Waybill-Sorter-transfer-20260615.zip` hand-off bundle had
already been deleted before this cleanup; it contained the same executables plus
source that is now in git.

`build.ps1` retires superseded artifacts automatically, so this folder always
describes exactly one build.

## Rebuilding and publishing

From the project root:

```powershell
.\build.ps1
.\.venv\Scripts\python.exe releasectl.py manifest --version 0.1.0
.\.venv\Scripts\python.exe releasectl.py publish  --version 0.1.0
```

`build.ps1` recreates `.venv`, installs `requirements.txt`, regenerates the
icon, and writes `releases\Unitrans-Portal-<version>.exe` with its SHA-256.
`publish` creates the git tag and GitHub release and attaches the exe plus the
manifest.

Do **not** publish with `--prerelease`: GitHub hides pre-releases from
`/releases/latest`, which is the URL the app's update check reads. A beta is a
normal release here, labelled "Beta" by its version number.

Test a build before publishing it:

```powershell
.\.venv\Scripts\python.exe -m unittest discover -s tests   # 21 tests
.\.venv\Scripts\python.exe tests\smoke_gui.py --screenshot gui.png
```

The packaged executable can also be exercised without the GUI:

```powershell
.\releases\Unitrans-Portal-0.1.0.exe --excel book.xlsx --pdf-folder .\pdfs --make-print-pdf
```

Releases are built with Python 3.14.6. The original 1.0 build used 3.12, so
re-test after any change of interpreter.
