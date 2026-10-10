"""Update checking against the published release manifest.

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

Nothing here installs anything. Downloading and replacing the running
executable is a later, separate step with its own failure handling; this module
only answers "is there something newer, and what changed?".
"""

from __future__ import annotations

import json
import ssl
import urllib.error
import urllib.request
from dataclasses import dataclass, field
from typing import Any

from version import APP_VERSION, compare_versions

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
    if not isinstance(download_url, str) or not download_url.startswith(("https://", "http://")):
        raise UpdateCheckError("manifest has no usable 'download_url'")

    sha256 = file_block.get("sha256")
    if not isinstance(sha256, str) or len(sha256) != 64:
        raise UpdateCheckError("manifest has no usable 'file.sha256'")

    return version.strip(), download_url, sha256.lower(), str(payload.get("release_page") or "")


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
