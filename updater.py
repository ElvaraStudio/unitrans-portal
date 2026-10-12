"""Verified update downloads and update checking against the release manifest.

Kept separate from the GUI so the logic that decides "is there a newer
version?" is testable without a display, and so a network problem can never
take the application down with it. Every function here returns a result object
instead of raising for expected conditions: being offline is normal, not a
crash.

How it works
------------
``releases/latest/download/manifest.json`` on GitHub is a small JSON document
naming the newest version, where to download it, its size and its SHA-256.
``/releases/latest`` is what makes this work without an API token and without
the app knowing any version numbers in advance.

Downloads are staged separately from the running executable, streamed to disk,
and checked against the manifest's SHA-256 before they can be applied. On
Windows a separate PowerShell helper waits for the app to exit, swaps the
verified executable, and restores the previous copy if the new app cannot
start.
"""

from __future__ import annotations

import json
import base64
import hashlib
import hmac
import os
import re
import shutil
import ssl
import subprocess
import sys
import tempfile
import threading
import time
import urllib.error
import urllib.request
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any
from urllib.parse import urlsplit

from version import APPDATA_DIR_NAME, APP_NAME, APP_VERSION, compare_versions

#: Where the manifest lives. `/releases/latest` resolves to the newest release
#: that is neither a draft nor a pre-release, so publishing a version as a
#: GitHub pre-release makes it invisible here.
MANIFEST_URL = "https://github.com/ElvaraStudio/unitrans-portal/releases/latest/download/manifest.json"

#: A hung connection must not leave the user staring at a spinner.
DEFAULT_TIMEOUT = 15.0

USER_AGENT = f"UnitransPortal/{APP_VERSION} (+https://github.com/ElvaraStudio/unitrans-portal)"


class UpdateCheckError(Exception):
    """Raised only for a malformed response - i.e. a bug or interference.

    Ordinary problems (offline, DNS failure, timeout, 404) are reported through
    ``UpdateCheckResult.error`` instead, because they are expected in normal
    use and the UI should say "couldn't check", not show a crash.
    """


class UpdateDownloadError(Exception):
    """A release could not be safely downloaded or verified."""


class UpdateCancelled(UpdateDownloadError):
    """The user cancelled an update download."""


@dataclass
class UpdateCheckResult:
    """Outcome of an update check, success or failure."""

    status: str                     # "up-to-date" | "update-available" | "error"
    current_version: str = APP_VERSION
    remote_version: str | None = None
    download_url: str | None = None
    release_page: str | None = None
    size_bytes: int | None = None
    sha256: str | None = None
    notes: dict[str, Any] = field(default_factory=dict)
    error: str | None = None

    @property
    def ok(self) -> bool:
        return self.status != "error"

    @property
    def update_available(self) -> bool:
        return self.status == "update-available"

    def notes_for(self, language: str) -> list[str]:
        """Release notes in ``language``, falling back to English."""
        block = self.notes.get(language) or self.notes.get("en") or {}
        if isinstance(block, dict):
            return list(block.get("notes") or [])
        return []

    def headline_for(self, language: str) -> str:
        block = self.notes.get(language) or self.notes.get("en") or {}
        if isinstance(block, dict):
            return str(block.get("headline") or "")
        return ""


def _validate_manifest(payload: Any) -> tuple[str, str, str, str]:
    """Check the manifest has the fields we depend on. Returns the key values."""
    if not isinstance(payload, dict):
        raise UpdateCheckError("manifest is not a JSON object")

    version = payload.get("version")
    if not isinstance(version, str) or not version.strip():
        raise UpdateCheckError("manifest has no usable 'version'")

    file_block = payload.get("file")
    if not isinstance(file_block, dict):
        raise UpdateCheckError("manifest has no 'file' block")

    download_url = payload.get("download_url")
    if not isinstance(download_url, str):
        raise UpdateCheckError("manifest has no usable 'download_url'")

    try:
        parts = urlsplit(download_url)
        port = parts.port
    except ValueError as exc:
        raise UpdateCheckError("manifest has no usable official GitHub release download URL") from exc
    expected_name = f"{APP_NAME.replace(' ', '-')}-{version}.exe"
    expected_path = f"/ElvaraStudio/unitrans-portal/releases/download/v{version}/{expected_name}"
    if (
        parts.scheme != "https"
        or parts.hostname != "github.com"
        or parts.username is not None
        or parts.password is not None
        or port not in (None, 443)
        or parts.query
        or parts.fragment
        or parts.path != expected_path
    ):
        raise UpdateCheckError("manifest has no usable official GitHub release download URL")

    sha256 = file_block.get("sha256")
    if not isinstance(sha256, str) or not re.fullmatch(r"[0-9a-fA-F]{64}", sha256):
        raise UpdateCheckError("manifest has no usable 'file.sha256'")

    release_page = str(payload.get("release_page") or "")
    try:
        release_parts = urlsplit(release_page)
        release_port = release_parts.port
    except ValueError:
        release_page = ""
    else:
        expected_release_path = f"/ElvaraStudio/unitrans-portal/releases/tag/v{version}"
        if not (
            release_parts.scheme == "https"
            and release_parts.hostname == "github.com"
            and release_parts.username is None
            and release_parts.password is None
            and release_port in (None, 443)
            and not release_parts.query
            and not release_parts.fragment
            and release_parts.path == expected_release_path
        ):
            release_page = ""

    return version.strip(), download_url, sha256.lower(), release_page


def check_for_update(
    manifest_url: str = MANIFEST_URL,
    current_version: str = APP_VERSION,
    timeout: float = DEFAULT_TIMEOUT,
    opener=None,
) -> UpdateCheckResult:
    """Fetch the manifest and compare it with the running version.

    Never raises for network conditions; inspect ``result.error`` instead.
    ``opener`` exists so tests can inject a fake transport.
    """
    result = UpdateCheckResult(status="error", current_version=current_version)

    fetch = opener or urllib.request.urlopen
    try:
        request = urllib.request.Request(manifest_url, headers={"User-Agent": USER_AGENT})
        with fetch(request, timeout=timeout) as response:
            raw = response.read()
    except urllib.error.HTTPError as exc:
        # 404 is the expected answer for a private repo or a pre-release-only
        # release, and is worth naming precisely because both are easy to
        # misconfigure.
        if exc.code == 404:
            result.error = "The update server returned 404. No published release was found."
        else:
            result.error = f"The update server returned HTTP {exc.code}."
        return result
    except (urllib.error.URLError, TimeoutError, OSError, ssl.SSLError) as exc:
        result.error = f"Could not reach the update server ({exc.__class__.__name__})."
        return result
    except Exception as exc:  # defensive: never let a check crash the app
        result.error = f"Update check failed: {exc.__class__.__name__}: {exc}"
        return result

    try:
        payload = json.loads(raw.decode("utf-8"))
        version, download_url, sha256, release_page = _validate_manifest(payload)
    except UpdateCheckError as exc:
        result.error = f"The update manifest is malformed: {exc}"
        return result
    except (json.JSONDecodeError, UnicodeDecodeError) as exc:
        result.error = f"The update manifest could not be read ({exc.__class__.__name__})."
        return result

    file_block = payload.get("file") or {}
    size = file_block.get("size")

    result.remote_version = version
    result.download_url = download_url
    result.release_page = release_page
    result.sha256 = sha256
    result.size_bytes = size if isinstance(size, int) else None
    result.notes = payload.get("notes") or {}

    try:
        comparison = compare_versions(version, current_version)
    except ValueError as exc:
        result.error = f"The update manifest has an unreadable version ({exc})."
        return result

    result.status = "update-available" if comparison > 0 else "up-to-date"
    return result


def format_size(num_bytes: int | None) -> str:
    """Human-readable download size for the update prompt."""
    if not num_bytes or num_bytes <= 0:
        return ""
    if num_bytes >= 1024 * 1024:
        return f"{num_bytes / (1024 * 1024):.1f} MB"
    if num_bytes >= 1024:
        return f"{num_bytes / 1024:.0f} KB"
    return f"{num_bytes} bytes"


DOWNLOAD_CHUNK_SIZE = 1024 * 1024
DOWNLOAD_TIMEOUT = 30.0
MAX_UPDATE_BYTES = 512 * 1024 * 1024
OFFICIAL_ASSET_HOSTS = {
    "github.com",
    "objects.githubusercontent.com",
    "release-assets.githubusercontent.com",
}


def supports_self_update() -> bool:
    """Whether this is a packaged Windows build that can replace its exe."""
    return sys.platform == "win32" and bool(getattr(sys, "frozen", False))


def download_update(
    result: UpdateCheckResult,
    destination_dir: Path | None = None,
    timeout: float = DOWNLOAD_TIMEOUT,
    opener=None,
    progress_callback=None,
    cancel_event: threading.Event | None = None,
) -> Path:
    """Stream a release exe to a staging folder and verify its size and hash.

    The staged file is never executed here. The caller can only ask the
    separate installer helper to use it after this function returns success.
    ``opener`` and ``progress_callback`` make the file/network work testable
    without GitHub or a GUI.
    """
    if not result.update_available:
        raise UpdateDownloadError("there is no available update to download")
    if not result.remote_version or not result.download_url:
        raise UpdateDownloadError("the update manifest is missing its version or download URL")
    try:
        if compare_versions(result.remote_version, result.current_version) <= 0:
            raise UpdateDownloadError("the manifest does not describe a newer version")
    except ValueError as exc:
        raise UpdateDownloadError("the update manifest has an invalid version") from exc
    expected_hash = str(result.sha256 or "").lower()
    if not re.fullmatch(r"[0-9a-f]{64}", expected_hash):
        raise UpdateDownloadError("the update manifest has an invalid SHA-256 hash")
    _validate_official_download_url(result.download_url, result.remote_version)

    expected_size = result.size_bytes
    if expected_size is not None and (expected_size <= 0 or expected_size > MAX_UPDATE_BYTES):
        raise UpdateDownloadError("the update manifest has an invalid file size")

    root = destination_dir or (Path(tempfile.gettempdir()) / APPDATA_DIR_NAME / "updates")
    try:
        root.mkdir(parents=True, exist_ok=True)
        handle, temp_name = tempfile.mkstemp(
            dir=str(root), prefix=f".update-{result.remote_version}-", suffix=".part"
        )
    except (OSError, ValueError) as exc:
        raise UpdateDownloadError(f"could not create a staging file ({exc.__class__.__name__})") from exc

    partial_path = Path(temp_name)
    final_path = root / f"{APP_NAME.replace(' ', '-')}-{result.remote_version}-verified.exe"
    fetch = opener or urllib.request.urlopen
    digest = hashlib.sha256()
    received = 0

    try:
        request = urllib.request.Request(
            result.download_url,
            headers={"User-Agent": USER_AGENT, "Accept": "application/octet-stream"},
        )
        with os.fdopen(handle, "wb") as output, fetch(request, timeout=timeout) as response:
            final_url = response.geturl() if hasattr(response, "geturl") else result.download_url
            final_parts = urlsplit(final_url)
            if final_parts.scheme != "https" or final_parts.hostname not in OFFICIAL_ASSET_HOSTS:
                raise UpdateDownloadError("GitHub redirected the download to an untrusted host")

            while True:
                if cancel_event is not None and cancel_event.is_set():
                    raise UpdateCancelled("download cancelled")
                chunk = response.read(DOWNLOAD_CHUNK_SIZE)
                if not chunk:
                    break
                received += len(chunk)
                if received > MAX_UPDATE_BYTES or (expected_size and received > expected_size):
                    raise UpdateDownloadError("the downloaded file is larger than expected")
                output.write(chunk)
                digest.update(chunk)
                if progress_callback is not None:
                    try:
                        progress_callback(received, expected_size)
                    except Exception:
                        pass

        if cancel_event is not None and cancel_event.is_set():
            raise UpdateCancelled("download cancelled")
        if expected_size is not None and received != expected_size:
            raise UpdateDownloadError(
                f"download size did not match the manifest ({received} of {expected_size} bytes)"
            )
        if not hmac.compare_digest(digest.hexdigest(), expected_hash):
            raise UpdateDownloadError("downloaded file failed its SHA-256 check")

        os.replace(partial_path, final_path)
        return final_path
    except (UpdateDownloadError, UpdateCancelled):
        _discard_partial(partial_path)
        raise
    except (urllib.error.URLError, TimeoutError, OSError, ssl.SSLError) as exc:
        _discard_partial(partial_path)
        raise UpdateDownloadError(f"download failed ({exc.__class__.__name__})") from exc
    except Exception:
        _discard_partial(partial_path)
        raise


def _validate_official_download_url(url: str, version: str) -> None:
    try:
        parts = urlsplit(url)
        port = parts.port
    except ValueError as exc:
        raise UpdateDownloadError("the update URL is malformed") from exc
    expected_name = f"{APP_NAME.replace(' ', '-')}-{version}.exe"
    expected_path = f"/ElvaraStudio/unitrans-portal/releases/download/v{version}/{expected_name}"
    if (
        parts.scheme != "https"
        or parts.hostname != "github.com"
        or parts.username is not None
        or parts.password is not None
        or port not in (None, 443)
        or parts.query
        or parts.fragment
        or parts.path != expected_path
    ):
        raise UpdateDownloadError("the update URL is not an official GitHub release asset")


def _discard_partial(path: Path) -> None:
    try:
        path.unlink(missing_ok=True)
    except OSError:
        pass


def _powershell_literal(value: str) -> str:
    """Quote one string as a PowerShell single-quoted literal."""
    return "'" + value.replace("'", "''") + "'"


def build_update_helper_script(
    executable: Path,
    staged_executable: Path,
    expected_sha256: str,
    parent_pid: int,
    timeout_seconds: int = 120,
) -> str:
    """Build the hidden PowerShell helper that swaps and rolls back the exe."""
    expected_hash = expected_sha256.lower()
    if not re.fullmatch(r"[0-9a-f]{64}", expected_hash):
        raise ValueError("expected_sha256 must contain 64 hexadecimal characters")
    if parent_pid <= 0:
        raise ValueError("parent_pid must be a positive process id")

    target = str(executable.resolve())
    staged = str(staged_executable.resolve())
    log_path = str(Path(tempfile.gettempdir()) / f"UnitransPortal-update-{parent_pid}.log")
    script = r"""
$ErrorActionPreference = 'Stop'
$target = __TARGET__
$staged = __STAGED__
$expectedHash = __HASH__
$parentProcessId = __PARENT_PID__
$waitSeconds = __WAIT_SECONDS__
$previous = "$target.previous"
$replacement = "$target.update"
$logPath = __LOG_PATH__
$swapped = $false
$appClosed = $false

function Write-UpdateLog([string]$message) {
    try { Add-Content -LiteralPath $logPath -Value "$(Get-Date -Format o) $message" -Encoding UTF8 } catch {}
}

function Get-UpdateFileHash([string]$path) {
    $algorithm = [System.Security.Cryptography.SHA256]::Create()
    try {
        $stream = [System.IO.File]::OpenRead($path)
        try {
            $bytes = $algorithm.ComputeHash($stream)
            return ([System.BitConverter]::ToString($bytes)).Replace('-', '').ToLowerInvariant()
        } finally { $stream.Dispose() }
    } finally { $algorithm.Dispose() }
}

try {
    Write-UpdateLog 'Waiting for the app to close.'
    $deadline = [DateTime]::UtcNow.AddSeconds($waitSeconds)
    while ((Get-Process -Id $parentProcessId -ErrorAction SilentlyContinue) -and [DateTime]::UtcNow -lt $deadline) {
        Start-Sleep -Milliseconds 250
    }
    if (Get-Process -Id $parentProcessId -ErrorAction SilentlyContinue) {
        throw 'The app did not close before the update timed out.'
    }
    $appClosed = $true

    $actualHash = Get-UpdateFileHash $staged
    if ($actualHash -ne $expectedHash) { throw 'The staged update failed its SHA-256 check.' }
    if (-not (Test-Path -LiteralPath $target -PathType Leaf)) { throw 'The current application file is missing.' }

    if (Test-Path -LiteralPath $replacement) { Remove-Item -LiteralPath $replacement -Force }
    Copy-Item -LiteralPath $staged -Destination $replacement
    $replacementHash = Get-UpdateFileHash $replacement
    if ($replacementHash -ne $expectedHash) { throw 'The copied update failed its SHA-256 check.' }
    if (Test-Path -LiteralPath $previous) { Remove-Item -LiteralPath $previous -Force }
    Move-Item -LiteralPath $target -Destination $previous
    $swapped = $true
    Move-Item -LiteralPath $replacement -Destination $target
    $newProcess = Start-Process -FilePath $target -WorkingDirectory (Split-Path -Parent $target) -PassThru
    Start-Sleep -Seconds 3
    $newProcess.Refresh()
    if ($newProcess.HasExited) { throw 'The updated app did not stay open.' }
    $swapped = $false
    try { Remove-Item -LiteralPath $staged -Force } catch {}
    Write-UpdateLog 'Update installed and the app restarted.'
    exit 0
} catch {
    Write-UpdateLog "Update failed: $($_.Exception.Message)"
    if ($swapped) {
        try {
            if (Test-Path -LiteralPath $target) { Remove-Item -LiteralPath $target -Force }
            if (Test-Path -LiteralPath $replacement) { Remove-Item -LiteralPath $replacement -Force }
            if (Test-Path -LiteralPath $previous) {
                Move-Item -LiteralPath $previous -Destination $target
                Write-UpdateLog 'Restored the previous app after the update failed.'
                Start-Process -FilePath $target -WorkingDirectory (Split-Path -Parent $target) | Out-Null
            }
        } catch {
            Write-UpdateLog "Rollback failed: $($_.Exception.Message)"
        }
    } elseif ($appClosed -and (Test-Path -LiteralPath $target)) {
        try { if (Test-Path -LiteralPath $replacement) { Remove-Item -LiteralPath $replacement -Force } } catch {}
        try { Start-Process -FilePath $target -WorkingDirectory (Split-Path -Parent $target) | Out-Null } catch {}
    }
    exit 1
}
"""
    replacements = {
        "__TARGET__": _powershell_literal(target),
        "__STAGED__": _powershell_literal(staged),
        "__HASH__": _powershell_literal(expected_hash),
        "__PARENT_PID__": str(int(parent_pid)),
        "__WAIT_SECONDS__": str(max(1, int(timeout_seconds))),
        "__LOG_PATH__": _powershell_literal(log_path),
    }
    for token, replacement in replacements.items():
        script = script.replace(token, replacement)
    return script


def launch_update_helper(
    staged_executable: Path,
    expected_sha256: str,
    executable: Path | None = None,
    parent_pid: int | None = None,
    popen=None,
):
    """Start a hidden helper to replace the frozen Windows executable."""
    if not supports_self_update():
        raise UpdateDownloadError("in-app installation is available in packaged Windows builds only")

    target = (executable or Path(sys.executable)).resolve()
    if not target.is_file():
        raise UpdateDownloadError("could not locate the running application executable")
    if not Path(staged_executable).is_file():
        raise UpdateDownloadError("the verified update file is missing")

    script = build_update_helper_script(
        target,
        Path(staged_executable),
        expected_sha256,
        parent_pid or os.getpid(),
    )
    encoded = base64.b64encode(script.encode("utf-16le")).decode("ascii")
    powershell = shutil.which("powershell.exe")
    if not powershell:
        system_root = Path(os.environ.get("SystemRoot", r"C:\Windows"))
        powershell = str(system_root / "System32" / "WindowsPowerShell" / "v1.0" / "powershell.exe")
    if not Path(powershell).is_file():
        raise UpdateDownloadError("Windows PowerShell was not found to finish the update")

    start = popen or subprocess.Popen
    creation_flags = getattr(subprocess, "CREATE_NO_WINDOW", 0x08000000) | getattr(
        subprocess, "CREATE_NEW_PROCESS_GROUP", 0x00000200
    )
    try:
        process = start(
            [
                powershell,
                "-NoProfile",
                "-NonInteractive",
                "-ExecutionPolicy",
                "Bypass",
                "-EncodedCommand",
                encoded,
            ],
            cwd=str(target.parent),
            stdin=subprocess.DEVNULL,
            stdout=subprocess.DEVNULL,
            stderr=subprocess.DEVNULL,
            close_fds=True,
            creationflags=creation_flags,
        )
        # Give PowerShell a moment to start and report immediate launch errors
        # before the app closes itself.
        time.sleep(0.2)
        exit_code = process.poll()
        if exit_code not in (None, 0):
            raise UpdateDownloadError(f"the update helper exited immediately (code {exit_code})")
        return process
    except UpdateDownloadError:
        raise
    except Exception as exc:
        raise UpdateDownloadError(f"could not start the update helper ({exc.__class__.__name__})") from exc
