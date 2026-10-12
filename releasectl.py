"""Release helper: generate the update manifest and publish a GitHub release.

Two jobs, both run from the project root through the build venv:

    python releasectl.py manifest --version 0.1.0
        Writes releases/manifest.json describing the built artifact. Run this
        after build.ps1 so the hash and size match the real file.

    python releasectl.py publish --version 0.1.0 --repo ElvaraStudio/unitrans-portal
        Creates the git tag and GitHub release, attaches the exe and the
        manifest, and takes the release notes from changelog.json.

Why a manifest exists
---------------------
The updater needs one small, stable document it can fetch to answer "is there a
newer version, where is it, and is it intact?". Fetching and parsing the whole
GitHub releases API from the app would couple the app to GitHub's response
shape; this manifest is ours and changes only when we change it.

The hash is what makes the download safe: the updater refuses to run an
artifact whose SHA-256 does not match the manifest.
"""

from __future__ import annotations

import argparse
import hashlib
import json
import subprocess
import sys
from datetime import date
from pathlib import Path

ROOT = Path(__file__).resolve().parent
RELEASES_DIR = ROOT / "releases"
CHANGELOG = ROOT / "changelog.json"

DEFAULT_REPO = "ElvaraStudio/unitrans-portal"

#: Where the app looks for manifest.json. The `/releases/latest` URL is what
#: makes this work without an API token or any knowledge of version numbers -
#: but GitHub only resolves it to a release that is neither a draft nor a
#: pre-release. A pre-release is invisible here, so do NOT publish beta
#: versions as GitHub pre-releases; the app carries its own beta signal in the
#: version number ("0.1.0 (Beta)") and in the manifest's `beta` field.
MANIFEST_URL = f"https://github.com/{DEFAULT_REPO}/releases/latest/download/manifest.json"


def sha256_of(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as handle:
        for chunk in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def load_release_notes(version: str) -> dict | None:
    """Return the changelog entry for ``version``, or None if it is not listed."""
    if not CHANGELOG.exists():
        return None
    data = json.loads(CHANGELOG.read_text(encoding="utf-8"))
    for release in data.get("releases", []):
        if release.get("version") == version:
            return release
    return None


def artifact_path(version: str) -> Path:
    from version import APP_NAME

    return RELEASES_DIR / f"{APP_NAME.replace(' ', '-')}-{version}.exe"


def installer_path(version: str) -> Path:
    from version import APP_NAME

    return RELEASES_DIR / f"{APP_NAME.replace(' ', '-')}-{version}-Setup.exe"


def build_manifest(version: str, repo: str = DEFAULT_REPO) -> dict:
    from version import APP_NAME, PUBLISHER, build_info

    exe = artifact_path(version)
    if not exe.exists():
        raise SystemExit(
            f"Artifact not found: {exe}\n"
            f"Run build.ps1 first - the manifest must describe a real file."
        )

    info = build_info(version)
    entry = load_release_notes(version)
    download_url = f"https://github.com/{repo}/releases/download/v{version}/{exe.name}"

    manifest = {
        "app": APP_NAME,
        "publisher": PUBLISHER,
        "version": version,
        "display_version": info["display_version"],
        "published": (entry or {}).get("date", date.today().isoformat()),
        "beta": info["beta"],
        "download_url": download_url,
        "release_page": f"https://github.com/{repo}/releases/tag/v{version}",
        "file": {
            "name": exe.name,
            "size": exe.stat().st_size,
            "sha256": sha256_of(exe),
        },
        "notes": {
            "en": (entry or {}).get("en", {}),
            "zh": (entry or {}).get("zh", {}),
        },
    }
    return manifest


def cmd_manifest(args: argparse.Namespace) -> int:
    manifest = build_manifest(args.version, args.repo)
    RELEASES_DIR.mkdir(parents=True, exist_ok=True)
    out = RELEASES_DIR / "manifest.json"
    out.write_text(json.dumps(manifest, indent=2, ensure_ascii=False) + "\n", encoding="utf-8")

    file_info = manifest["file"]
    print(f"Wrote {out}")
    print(f"  version  : {manifest['version']} ({manifest['display_version']})")
    print(f"  file     : {file_info['name']}  {file_info['size']:,} bytes")
    print(f"  sha256   : {file_info['sha256']}")
    print(f"  download : {manifest['download_url']}")
    if not manifest["notes"]["en"]:
        print("  WARNING  : no changelog entry for this version - notes will be empty.")
    return 0


def cmd_publish(args: argparse.Namespace) -> int:
    from version import APP_NAME

    version = args.version
    tag = f"v{version}"
    exe = artifact_path(version)
    manifest_path = RELEASES_DIR / "manifest.json"

    if not exe.exists():
        raise SystemExit(f"Artifact not found: {exe}\nRun build.ps1 first.")
    if not manifest_path.exists():
        raise SystemExit("releases/manifest.json missing. Run: python releasectl.py manifest")

    # Guard against publishing a manifest that describes a different build.
    manifest = json.loads(manifest_path.read_text(encoding="utf-8"))
    if manifest.get("version") != version:
        raise SystemExit(
            f"manifest.json describes {manifest.get('version')}, not {version}. "
            "Regenerate it with: python releasectl.py manifest"
        )
    actual = sha256_of(exe)
    if manifest["file"]["sha256"] != actual:
        raise SystemExit(
            "The artifact changed after the manifest was generated.\n"
            f"  manifest: {manifest['file']['sha256']}\n"
            f"  actual  : {actual}\n"
            "Regenerate the manifest and re-run publish."
        )

    entry = load_release_notes(version) or {}
    notes_path = RELEASES_DIR / f"notes-{version}.md"
    notes_path.write_text(render_release_notes(version, entry), encoding="utf-8")

    title = f"{APP_NAME} {version}"
    if entry.get("en", {}).get("headline"):
        title = f"{APP_NAME} {version} - {entry['en']['headline']}"

    command = [
        "gh", "release", "create", tag,
        str(exe),
        str(manifest_path),
        "--repo", args.repo,
        "--title", title,
        "--notes-file", str(notes_path),
    ]
    installer = installer_path(version)
    if installer.exists():
        command.insert(command.index(str(manifest_path)), str(installer))
    # Deliberately NOT setting --prerelease from the manifest's `beta` flag.
    # A GitHub pre-release is excluded from /releases/latest, which would make
    # this version invisible to the updater. `beta` describes the app's own
    # maturity; it is not a GitHub release channel.
    if args.prerelease:
        command.append("--prerelease")
    if args.draft:
        command.append("--draft")

    print("Running:", " ".join(command))
    result = subprocess.run(command, cwd=ROOT)
    if result.returncode != 0:
        raise SystemExit(f"gh release create failed with exit code {result.returncode}")

    print()
    print(f"Published {tag} to {args.repo}")
    print(f"  release page : {manifest['release_page']}")
    print(f"  manifest     : {MANIFEST_URL}")
    return 0


def render_release_notes(version: str, entry: dict) -> str:
    """Markdown for the GitHub release body."""
    lines = [f"## Unitrans Portal {version}", ""]

    en = entry.get("en", {})
    zh = entry.get("zh", {})

    if en.get("headline"):
        lines += [f"**{en['headline']}**", ""]
    for note in en.get("notes", []):
        lines.append(f"- {note}")

    if zh.get("notes"):
        lines += ["", "---", "", "### 中文", ""]
        if zh.get("headline"):
            lines += [f"**{zh['headline']}**", ""]
        for note in zh["notes"]:
            lines.append(f"- {note}")

    manifest = build_manifest(version)
    lines += [
        "",
        "---",
        "",
        "### Download",
        "",
        f"`{manifest['file']['name']}` - {manifest['file']['size']:,} bytes",
        "",
        "SHA-256:",
        "",
        f"```\n{manifest['file']['sha256']}\n```",
        "",
        "The `.exe` runs on its own; no installation or Python is required.",
        "",
    ]
    installer = installer_path(version)
    if installer.exists():
        lines += [
            f"`{installer.name}` - Windows per-user installer ({installer.stat().st_size:,} bytes)",
            "",
            "SHA-256:",
            "",
            f"```\n{sha256_of(installer)}\n```",
            "",
        ]
    return "\n".join(lines)


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="Generate the update manifest and publish releases.")
    parser.add_argument("--repo", default=DEFAULT_REPO, help=f"GitHub repo (default: {DEFAULT_REPO})")
    sub = parser.add_subparsers(dest="command", required=True)

    p_manifest = sub.add_parser("manifest", help="Write releases/manifest.json for a built version.")
    p_manifest.add_argument("--version", required=True, help="Version to describe, e.g. 0.1.0")
    p_manifest.set_defaults(func=cmd_manifest)

    p_publish = sub.add_parser("publish", help="Create the GitHub release and attach the artifacts.")
    p_publish.add_argument("--version", required=True, help="Version to publish, e.g. 0.1.0")
    p_publish.add_argument("--draft", action="store_true", help="Create as a draft instead of publishing.")
    p_publish.add_argument(
        "--prerelease",
        action="store_true",
        help=(
            "Mark as a GitHub pre-release. WARNING: pre-releases are excluded from "
            "/releases/latest, so the updater will not see this version. Only use this "
            "for a build you do not intend users to pick up automatically."
        ),
    )
    p_publish.set_defaults(func=cmd_publish)

    args = parser.parse_args(argv)
    return args.func(args)


if __name__ == "__main__":
    sys.exit(main())
