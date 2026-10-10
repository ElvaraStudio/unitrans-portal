"""Single source of truth for the product version and product identity.

Everything that needs a version number reads it from here:

* ``waybill_sorter.py``  - the in-app footer and the "What's New" dialog.
* ``build.ps1``          - artifact file names, the exe version resource.
* the future updater     - comparing the running version against a manifest.

``build.ps1`` imports this module through the project venv instead of parsing
it with a regex, so this file stays the only place a version is ever edited.

Versioning
----------
This project uses semantic versioning. ``0.x.y`` means beta: the customer
should expect the interface and behaviour to still change. ``1.0.0`` is the
first release we are willing to call stable.

Version comparison is numeric, component by component, so ``0.10.0`` is
correctly newer than ``0.9.0``. A naive string comparison gets that backwards,
which would silently stop every future update from being detected. That is why
``parse_version`` and ``compare_versions`` live here with a test suite rather
than being inlined at the call site.
"""

from __future__ import annotations

import json
import re
from typing import Iterable, Sequence

# --------------------------------------------------------------------------
# Product identity - shown in the UI, the exe metadata, and the installer.
# --------------------------------------------------------------------------

APP_NAME = "Unitrans Portal"
APP_VERSION = "0.3.0"

#: True while the version is still 0.x. Drives the "(Beta)" suffix in the UI.
BETA = True

PUBLISHER = "Elvara Studio"
COPYRIGHT = "\u00a9 2026 Unitrans / Elvara Studio"

#: Used by the exe version resource and the installer. Kept separate from
#: APP_NAME because Windows also wants a short, filename-safe description.
FILE_DESCRIPTION = "Unitrans Portal - waybill sorting and PDF preparation"

#: Folder name created under %LOCALAPPDATA%\Programs by the installer.
INSTALL_DIR_NAME = APP_NAME

#: Stable per-user data location (settings, update state, logs).
APPDATA_DIR_NAME = "UnitransPortal"


# --------------------------------------------------------------------------
# Version parsing and comparison
# --------------------------------------------------------------------------

#: Accepts "1.2.3", "v1.2.3", "0.1.0-beta.1", "1.2.3+build.7".
_VERSION_RE = re.compile(
    r"""
    ^
    v?
    (?P<release>\d+(?:\.\d+)*)          # numeric release: 0.1.0
    (?:-(?P<prerelease>[0-9A-Za-z.\-]+))?   # optional -beta.1
    (?:\+(?P<build>[0-9A-Za-z.\-]+))?       # optional +build.7
    $
    """,
    re.VERBOSE,
)


def parse_version(text: str) -> tuple[tuple[int, ...], tuple[tuple[int, object], ...]]:
    """Split a version string into ``(release, prerelease)``.

    ``release`` is always a tuple of ints. ``prerelease`` is a tuple of
    ``(rank, identifier)`` pairs compared component-wise; an empty tuple means
    a final release, which sorts *after* any prerelease of the same number.

    Raises ``ValueError`` on anything unparseable, so a malformed remote
    manifest fails loudly instead of silently comparing as equal.
    """
    if not isinstance(text, str):
        raise ValueError(f"version must be a string, got {type(text).__name__}")

    match = _VERSION_RE.match(text.strip())
    if not match:
        raise ValueError(f"unparseable version: {text!r}")

    release = tuple(int(part) for part in match.group("release").split("."))

    prerelease: list[tuple[int, object]] = []
    raw_pre = match.group("prerelease")
    if raw_pre:
        for part in raw_pre.split("."):
            if part.isdigit():
                # Numeric identifiers sort before alphabetic ones (SemVer rule).
                prerelease.append((0, int(part)))
            else:
                prerelease.append((1, part.lower()))

    return release, tuple(prerelease)


def _pad(left: Sequence[int], right: Sequence[int]) -> tuple[tuple[int, ...], tuple[int, ...]]:
    width = max(len(left), len(right))
    return (
        tuple(left) + (0,) * (width - len(left)),
        tuple(right) + (0,) * (width - len(right)),
    )


def compare_versions(left: str, right: str) -> int:
    """Return -1 if ``left < right``, 0 if equal, 1 if ``left > right``.

    Release numbers are compared numerically first; only when they are equal
    does the prerelease decide, so ``0.1.0`` is newer than ``0.1.0-beta.1``.
    """
    left_release, left_pre = parse_version(left)
    right_release, right_pre = parse_version(right)

    padded_left, padded_right = _pad(left_release, right_release)
    if padded_left != padded_right:
        return -1 if padded_left < padded_right else 1

    # Same release number: a final release beats any prerelease.
    if not left_pre and not right_pre:
        return 0
    if not left_pre:
        return 1
    if not right_pre:
        return -1

    if left_pre == right_pre:
        return 0
    return -1 if left_pre < right_pre else 1


def is_newer(candidate: str, current: str = APP_VERSION) -> bool:
    """True when ``candidate`` is a strictly newer version than ``current``."""
    return compare_versions(candidate, current) > 0


def sorted_versions(versions: Iterable[str], reverse: bool = True) -> list[str]:
    """Sort version strings newest-first (default) using numeric comparison."""
    from functools import cmp_to_key

    return sorted(versions, key=cmp_to_key(compare_versions), reverse=reverse)


# --------------------------------------------------------------------------
# Display helpers
# --------------------------------------------------------------------------

def display_version(version: str = APP_VERSION) -> str:
    """``0.1.0 (Beta)`` while the release line is still 0.x."""
    release, _ = parse_version(version)
    if release and release[0] == 0:
        return f"{version} (Beta)"
    return version


def version_label(version: str = APP_VERSION) -> str:
    """``Version 0.1.0 (Beta)`` - the text shown in the UI footer."""
    return f"Version {display_version(version)}"


def file_version_tuple(version: str = APP_VERSION) -> tuple[int, int, int, int]:
    """Windows version resource fields are exactly four 16-bit integers."""
    release, _ = parse_version(version)
    padded = tuple(release[:4]) + (0,) * (4 - len(release[:4]))
    return padded  # type: ignore[return-value]


def file_version_string(version: str = APP_VERSION) -> str:
    """``0.1.0.0`` - the ``FileVersion`` string Windows Explorer displays."""
    return ".".join(str(part) for part in file_version_tuple(version))


def build_info(version: str = APP_VERSION) -> dict:
    """Machine-readable product identity, consumed by ``build.ps1``."""
    return {
        "app_name": APP_NAME,
        "version": version,
        "display_version": display_version(version),
        "file_version": file_version_string(version),
        "publisher": PUBLISHER,
        "copyright": COPYRIGHT,
        "file_description": FILE_DESCRIPTION,
        "install_dir_name": INSTALL_DIR_NAME,
        "appdata_dir_name": APPDATA_DIR_NAME,
        "beta": BETA,
    }


def pyinstaller_version_file(version: str = APP_VERSION) -> str:
    """Render a PyInstaller ``--version-file`` payload.

    Without this the built exe has completely blank version metadata, so
    right-clicking it in Explorer cannot answer "which build is this?". With
    it, ``FileVersion`` / ``ProductName`` / ``CompanyName`` are populated and
    two builds can be told apart without launching either of them.
    """
    quad = file_version_tuple(version)
    return f"""# Generated by version.py - do not edit by hand.
# Consumed by PyInstaller --version-file to populate the exe properties.
VSVersionInfo(
  ffi=FixedFileInfo(
    filevers={quad},
    prodvers={quad},
    mask=0x3f,
    flags=0x0,
    OS=0x40004,
    fileType=0x1,
    subtype=0x0,
    date=(0, 0)
  ),
  kids=[
    StringFileInfo([
      StringTable(
        '040904B0',
        [StringStruct('CompanyName', {PUBLISHER!r}),
         StringStruct('FileDescription', {FILE_DESCRIPTION!r}),
         StringStruct('FileVersion', {file_version_string(version)!r}),
         StringStruct('InternalName', 'waybill_sorter'),
         StringStruct('LegalCopyright', {COPYRIGHT!r}),
         StringStruct('OriginalFilename', 'Unitrans Portal.exe'),
         StringStruct('ProductName', {APP_NAME!r}),
         StringStruct('ProductVersion', {version!r})])
    ]),
    VarFileInfo([VarStruct('Translation', [1033, 1200])])
  ]
)
"""


def _main(argv: list[str]) -> int:
    """``python version.py [--json]`` - read by build.ps1, no dependencies."""
    import argparse

    parser = argparse.ArgumentParser(description="Print product version metadata.")
    parser.add_argument("--json", action="store_true", help="Print the full metadata as JSON.")
    parser.add_argument("--version", action="store_true", help="Print just the version number.")
    parser.add_argument("--file-version", action="store_true", help="Print the 4-part Windows file version.")
    parser.add_argument("--label", action="store_true", help="Print the UI footer label.")
    parser.add_argument(
        "--pyinstaller-version-file",
        action="store_true",
        help="Print the PyInstaller VSVersionInfo payload (redirect into a file).",
    )
    args = parser.parse_args(argv)

    if args.version:
        print(APP_VERSION)
    elif args.file_version:
        print(file_version_string())
    elif args.label:
        print(version_label())
    elif args.pyinstaller_version_file:
        # Print exactly the payload so PowerShell can redirect stdout into a file.
        import sys

        sys.stdout.write(pyinstaller_version_file())
    else:
        print(json.dumps(build_info(), indent=2))
    return 0


if __name__ == "__main__":
    import sys

    raise SystemExit(_main(sys.argv[1:]))
