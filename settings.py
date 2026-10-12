"""Per-user settings, stored outside the application folder.

The app is installed per-user and may be replaced in place by an update, so
preferences and update state must live somewhere the update does not touch:

    %APPDATA%\\UnitransPortal\\settings.json

Keeping this separate also means an update never silently resets a user's
choices. Writes are atomic (temp file + replace) so an interrupted save cannot
leave a half-written settings file that fails to parse on next launch.
"""

from __future__ import annotations

import json
import os
import tempfile
from pathlib import Path
from typing import Any

from version import APPDATA_DIR_NAME

SETTINGS_FILENAME = "settings.json"

#: Defaults. Anything not listed here is preserved when settings are saved, so
#: an older build does not discard keys written by a newer one.
DEFAULTS: dict[str, Any] = {
    # Automatic update checks are opt-in and not yet implemented; the flag is
    # stored now so enabling the feature later needs no migration.
    "auto_check_updates": False,
    # ISO-8601 timestamp of the last successful check, for "last checked" text.
    "last_update_check": "",
    # The version the user was last told about, so a prompt is not repeated on
    # every launch for the same release.
    "last_seen_version": "",
    # Colour scheme: "light" or "dark". Read by theme.py.
    "theme": "light",
}


def config_dir() -> Path:
    """Per-user configuration directory, created on demand."""
    if os.name == "nt":
        base = os.environ.get("APPDATA") or os.environ.get("LOCALAPPDATA")
        root = Path(base) if base else Path.home() / "AppData" / "Roaming"
    else:
        root = Path(os.environ.get("XDG_CONFIG_HOME") or Path.home() / ".config")
    return root / APPDATA_DIR_NAME


def settings_path() -> Path:
    return config_dir() / SETTINGS_FILENAME


def load_settings(path: Path | None = None) -> dict[str, Any]:
    """Read settings, falling back to defaults.

    A corrupt or unreadable file is treated as "no settings yet" rather than an
    error: losing a preference is trivial, refusing to start is not.
    """
    target = path or settings_path()
    settings = dict(DEFAULTS)
    try:
        raw = target.read_text(encoding="utf-8")
    except (FileNotFoundError, NotADirectoryError):
        return settings
    except OSError:
        return settings

    try:
        stored = json.loads(raw)
    except (json.JSONDecodeError, UnicodeDecodeError):
        return settings

    if isinstance(stored, dict):
        settings.update(stored)
    return settings


def save_settings(settings: dict[str, Any], path: Path | None = None) -> bool:
    """Write settings atomically. Returns True on success.

    Failure is reported, not raised - a read-only or full disk should not stop
    the user from working.
    """
    target = path or settings_path()
    merged = dict(DEFAULTS)
    merged.update(settings)

    try:
        target.parent.mkdir(parents=True, exist_ok=True)
        # Write beside the target so os.replace stays on the same filesystem
        # and is therefore atomic.
        handle, temp_name = tempfile.mkstemp(
            dir=str(target.parent), prefix=".settings-", suffix=".tmp"
        )
        try:
            with os.fdopen(handle, "w", encoding="utf-8") as stream:
                json.dump(merged, stream, indent=2, ensure_ascii=False)
                stream.write("\n")
            os.replace(temp_name, target)
        except BaseException:
            # Do not leave a stray temp file behind on any failure path.
            try:
                os.unlink(temp_name)
            except OSError:
                pass
            raise
    except (OSError, ValueError):
        # ValueError covers a malformed path (e.g. an embedded null), which
        # pathlib raises instead of OSError on some platforms. Reporting the
        # failure is enough; the caller decides whether to tell the user.
        return False

    return True


def update_settings(path: Path | None = None, **changes: Any) -> dict[str, Any]:
    """Load, merge ``changes``, save, and return the resulting settings."""
    settings = load_settings(path)
    settings.update(changes)
    save_settings(settings, path)
    return settings
