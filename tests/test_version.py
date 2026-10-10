"""Tests for version identity and update-manifest integrity.

Run with:

    .\\.venv\\Scripts\\python.exe -m pytest -q
    .\\.venv\\Scripts\\python.exe -m unittest discover -s tests -v

These cover the parts that fail *silently* and expensively: a version
comparison that stops detecting updates, or a manifest whose hash does not
match the artifact it describes. The GUI is verified by running it.
"""

from __future__ import annotations

import hashlib
import json
import sys
import tempfile
import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))

import version  # noqa: E402
import releasectl  # noqa: E402


class TestVersionComparison(unittest.TestCase):
    """Numeric, component-wise comparison. The 0.10.0 case is the reason."""

    def test_numeric_not_lexicographic(self):
        # A naive string comparison says "0.10.0" < "0.9.0", which would stop
        # every future update from ever being detected.
        self.assertEqual(version.compare_versions("0.10.0", "0.9.0"), 1)
        self.assertEqual(version.compare_versions("0.9.0", "0.10.0"), -1)
        self.assertTrue(version.is_newer("0.10.0", "0.9.0"))
        self.assertFalse(version.is_newer("0.9.0", "0.10.0"))

    def test_ordering_across_magnitudes(self):
        ordered = ["0.1.0", "0.2.0", "0.9.0", "0.10.0", "0.11.0", "0.20.0", "1.0.0", "1.0.1", "2.0.0"]
        for older, newer in zip(ordered, ordered[1:]):
            self.assertEqual(version.compare_versions(older, newer), -1, f"{older} < {newer}")
            self.assertEqual(version.compare_versions(newer, older), 1, f"{newer} > {older}")

    def test_equal_versions(self):
        for left, right in [("0.1.0", "0.1.0"), ("v0.1.0", "0.1.0"), ("0.1.0", "0.1"), ("0.1.0", "0.1.0.0")]:
            self.assertEqual(version.compare_versions(left, right), 0, f"{left} == {right}")

    def test_prerelease_sorts_before_final(self):
        self.assertEqual(version.compare_versions("0.1.0-beta.1", "0.1.0"), -1)
        self.assertEqual(version.compare_versions("0.1.0", "0.1.0-beta.1"), 1)
        self.assertEqual(version.compare_versions("0.1.0-beta.1", "0.1.0-beta.2"), -1)
        self.assertEqual(version.compare_versions("0.1.0-alpha", "0.1.0-beta"), -1)

    def test_build_metadata_is_ignored_for_ordering(self):
        self.assertEqual(version.compare_versions("0.1.1+hotfix", "0.1.1"), 0)

    def test_unparseable_version_raises(self):
        # A malformed remote manifest must fail loudly, not compare as equal.
        for bad in ["", "abc", "1.2.x", "1..2", None, 123]:
            with self.assertRaises(ValueError):
                version.parse_version(bad)  # type: ignore[arg-type]

    def test_sorted_versions_newest_first(self):
        result = version.sorted_versions(["0.9.0", "0.1.0", "0.10.0", "0.2.0"])
        self.assertEqual(result, ["0.10.0", "0.9.0", "0.2.0", "0.1.0"])


class TestVersionMetadata(unittest.TestCase):
    def test_file_version_is_four_ints(self):
        quad = version.file_version_tuple()
        self.assertEqual(len(quad), 4)
        self.assertTrue(all(isinstance(part, int) for part in quad))
        self.assertEqual(version.file_version_string("1.2.3"), "1.2.3.0")
        self.assertEqual(version.file_version_string("1.2.3.4"), "1.2.3.4")

    def test_beta_suffix_reflects_major_zero(self):
        self.assertEqual(version.display_version("0.1.0"), "0.1.0 (Beta)")
        self.assertEqual(version.display_version("0.9.9"), "0.9.9 (Beta)")
        # Once the release line reaches 1.x the beta marker disappears by itself.
        self.assertEqual(version.display_version("1.0.0"), "1.0.0")

    def test_version_label_format(self):
        self.assertEqual(version.version_label("0.1.0"), "Version 0.1.0 (Beta)")

    def test_build_info_is_json_serialisable(self):
        info = version.build_info()
        json.dumps(info)  # must not raise
        for key in ("app_name", "version", "file_version", "publisher", "copyright"):
            self.assertIn(key, info)
            self.assertTrue(info[key], f"{key} must not be empty")

    def test_pyinstaller_version_file_is_valid_python(self):
        import ast

        payload = version.pyinstaller_version_file()
        ast.parse(payload)  # PyInstaller evals this file
        self.assertIn("VSVersionInfo", payload)
        self.assertIn(version.APP_NAME, payload)
        self.assertIn(version.PUBLISHER, payload)

    def test_current_version_matches_changelog(self):
        """The running version must be documented, or What's New opens empty."""
        changelog = json.loads((ROOT / "changelog.json").read_text(encoding="utf-8"))
        versions = [release["version"] for release in changelog["releases"]]
        self.assertIn(
            version.APP_VERSION,
            versions,
            f"changelog.json has no entry for {version.APP_VERSION}; add one before releasing.",
        )

    def test_changelog_translations_are_in_step(self):
        changelog = json.loads((ROOT / "changelog.json").read_text(encoding="utf-8"))
        for release in changelog["releases"]:
            with self.subTest(version=release["version"]):
                english = release.get("en", {}).get("notes", [])
                chinese = release.get("zh", {}).get("notes", [])
                self.assertTrue(english, "English notes must not be empty")
                self.assertEqual(
                    len(english),
                    len(chinese),
                    f"{release['version']}: {len(english)} English vs {len(chinese)} Chinese notes",
                )


class TestTranslations(unittest.TestCase):
    """The English and Chinese dictionaries must offer the same keys.

    ``WaybillSorterApp.tr`` falls back to the English string when a key is
    missing, so a gap does not crash - a Chinese user just silently gets an
    English dialog. That is easy to miss by eye, which is why it is tested.

    The dictionaries are read from the source, not imported, so this test does
    not construct a Tk root or need a display.
    """

    @classmethod
    def setUpClass(cls):
        import ast

        source = (ROOT / "waybill_sorter.py").read_text(encoding="utf-8")
        cls.languages: dict[str, set[str]] = {}
        for node in ast.walk(ast.parse(source)):
            if not isinstance(node, ast.Assign):
                continue
            if not any(
                isinstance(target, ast.Name) and target.id == "I18N"
                for target in node.targets
            ):
                continue
            for key, value in zip(node.value.keys, node.value.values):
                language = key.value  # type: ignore[attr-defined]
                cls.languages[language] = {
                    entry.value for entry in value.keys  # type: ignore[attr-defined]
                }

    def test_both_languages_present(self):
        self.assertEqual(set(self.languages), {"en", "zh"})

    def test_key_sets_match(self):
        english = self.languages.get("en", set())
        chinese = self.languages.get("zh", set())
        self.assertTrue(english, "English dictionary appears to be empty")
        missing_chinese = sorted(english - chinese)
        missing_english = sorted(chinese - english)
        self.assertEqual(
            missing_chinese,
            [],
            f"These keys exist in English but not Chinese: {missing_chinese}",
        )
        self.assertEqual(
            missing_english,
            [],
            f"These keys exist in Chinese but not English: {missing_english}",
        )

    def test_no_empty_translations(self):
        import ast

        source = (ROOT / "waybill_sorter.py").read_text(encoding="utf-8")
        for node in ast.walk(ast.parse(source)):
            if isinstance(node, ast.Assign) and any(
                isinstance(target, ast.Name) and target.id == "I18N"
                for target in node.targets
            ):
                for key, value in zip(node.value.keys, node.value.values):
                    language = key.value  # type: ignore[attr-defined]
                    for entry in value.keys:  # type: ignore[attr-defined]
                        self.assertTrue(
                            entry.value.strip(),  # type: ignore[attr-defined]
                            f"empty translation in {language}: {entry.value}",  # type: ignore[attr-defined]
                        )


class TestReleaseManifest(unittest.TestCase):
    """The manifest is what the updater trusts, so its integrity checks matter."""

    def test_sha256_matches_hashlib(self):
        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / "sample.bin"
            payload = b"unitrans" * 5000
            path.write_bytes(payload)
            self.assertEqual(
                releasectl.sha256_of(path),
                hashlib.sha256(payload).hexdigest(),
            )

    def test_manifest_url_points_at_the_latest_release(self):
        # This exact URL shape is what the updater will fetch; /releases/latest
        # resolves without an API token.
        self.assertTrue(releasectl.MANIFEST_URL.endswith("/releases/latest/download/manifest.json"))
        self.assertIn(releasectl.DEFAULT_REPO, releasectl.MANIFEST_URL)

    def test_missing_artifact_is_refused(self):
        # A manifest must never be written for a build that does not exist.
        with self.assertRaises(SystemExit):
            releasectl.build_manifest("99.99.99")

    def test_artifact_path_is_versioned(self):
        path = releasectl.artifact_path("0.1.0")
        self.assertEqual(path.name, "Unitrans-Portal-0.1.0.exe")
        self.assertEqual(path.parent.name, "releases")


if __name__ == "__main__":
    unittest.main(verbosity=2)
