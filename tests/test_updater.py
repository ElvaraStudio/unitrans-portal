"""Tests for update checking and settings persistence.

The update check is the part of this app that talks to the outside world, so it
gets tested against a fake transport rather than the real GitHub: that keeps the
suite offline-safe, fast, and able to exercise responses we cannot provoke on
demand - a newer version, a truncated manifest, a 500, a TLS failure.

Run with:

    .\\.venv\\Scripts\\python.exe -m unittest discover -s tests -v
"""

from __future__ import annotations

import io
import hashlib
import json
import sys
import tempfile
import threading
import unittest
import urllib.error
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))

import settings as app_settings  # noqa: E402
import updater  # noqa: E402
from version import APP_VERSION  # noqa: E402


def next_version(current: str = APP_VERSION) -> str:
    """A version guaranteed to be newer than the one under test.

    Deriving this instead of writing "0.2.0" keeps the suite correct across a
    version bump - a hardcoded number silently stops being "newer" the moment
    the app catches up with it, and the test then fails for the wrong reason.
    """
    parts = [int(piece) for piece in current.split("-")[0].split(".")]
    while len(parts) < 3:
        parts.append(0)
    parts[1] += 1
    parts[2] = 0
    return ".".join(str(piece) for piece in parts[:3])


#: A release clearly ahead of the running version, for "update available" cases.
NEWER = next_version()


def manifest_bytes(
    version: str,
    sha256: str = "a" * 64,
    size: int = 24_403_427,
    notes: bool = True,
) -> bytes:
    payload = {
        "app": "Unitrans Portal",
        "version": version,
        "display_version": version,
        "download_url": (
            f"https://github.com/ElvaraStudio/unitrans-portal/releases/download/"
            f"v{version}/Unitrans-Portal-{version}.exe"
        ),
        "release_page": (
            f"https://github.com/ElvaraStudio/unitrans-portal/releases/tag/v{version}"
        ),
        "file": {
            "name": f"Unitrans-Portal-{version}.exe",
            "size": size,
            "sha256": sha256,
        },
    }
    if notes:
        payload["notes"] = {
            "en": {"headline": "Test release", "notes": ["English note one", "English note two"]},
            "zh": {"headline": "\u6d4b\u8bd5", "notes": ["\u4e2d\u6587\u8bf4\u660e\u4e00", "\u4e2d\u6587\u8bf4\u660e\u4e8c"]},
        }
    return json.dumps(payload).encode("utf-8")


class FakeResponse(io.BytesIO):
    """Minimal stand-in for the object urlopen returns."""

    def __enter__(self):
        return self

    def __exit__(self, *exc_info):
        self.close()
        return False


def opener_returning(body: bytes):
    def _open(request, timeout=None):
        return FakeResponse(body)

    return _open


def opener_raising(exc: BaseException):
    def _open(request, timeout=None):
        # HTTPError subclasses a file wrapper, so an unclosed instance emits a
        # ResourceWarning when it is garbage collected. Close it explicitly to
        # keep the suite's output clean.
        if isinstance(exc, urllib.error.HTTPError):
            try:
                exc.close()
            except Exception:
                pass
        raise exc

    return _open


class TestUpdateCheck(unittest.TestCase):
    def test_current_version_reports_up_to_date(self):
        result = updater.check_for_update(opener=opener_returning(manifest_bytes(APP_VERSION)))
        self.assertTrue(result.ok)
        self.assertFalse(result.update_available)
        self.assertEqual(result.status, "up-to-date")
        self.assertIsNone(result.error)

    def test_newer_version_reports_update_available(self):
        # The case that cannot be provoked on demand: the released manifest
        # always describes a version at or below the one being developed, so
        # the newer side is synthesised.
        result = updater.check_for_update(
            current_version="0.1.0", opener=opener_returning(manifest_bytes("0.2.0"))
        )
        self.assertTrue(result.update_available)
        self.assertEqual(result.remote_version, "0.2.0")
        self.assertEqual(result.size_bytes, 24_403_427)
        self.assertEqual(len(result.sha256 or ""), 64)

    def test_release_ahead_of_this_build_is_offered(self):
        # Guards the suite itself: NEWER must stay ahead of APP_VERSION across
        # version bumps, or every "update available" case silently inverts.
        result = updater.check_for_update(opener=opener_returning(manifest_bytes(NEWER)))
        self.assertTrue(
            result.update_available,
            f"{NEWER} should be newer than the running {APP_VERSION}",
        )

    def test_numeric_comparison_not_lexicographic(self):
        # 0.10.0 is newer than 0.9.0; a string compare would say otherwise and
        # the update would never be offered.
        result = updater.check_for_update(
            current_version="0.9.0", opener=opener_returning(manifest_bytes("0.10.0"))
        )
        self.assertTrue(result.update_available)

    def test_older_remote_version_is_not_an_update(self):
        # A user on a newer build than the published manifest (e.g. a rollback
        # on the server) must not be told to "update" downwards.
        result = updater.check_for_update(
            current_version="0.5.0", opener=opener_returning(manifest_bytes("0.1.0"))
        )
        self.assertTrue(result.ok)
        self.assertFalse(result.update_available)

    def test_prerelease_ordering(self):
        result = updater.check_for_update(
            current_version="0.1.0-beta.1", opener=opener_returning(manifest_bytes("0.1.0"))
        )
        self.assertTrue(result.update_available)

    def test_notes_are_language_aware(self):
        result = updater.check_for_update(opener=opener_returning(manifest_bytes(NEWER)))
        self.assertEqual(len(result.notes_for("en")), 2)
        self.assertEqual(len(result.notes_for("zh")), 2)
        self.assertNotEqual(result.notes_for("en"), result.notes_for("zh"))
        # An unknown language falls back to English rather than showing nothing.
        self.assertEqual(result.notes_for("de"), result.notes_for("en"))
        self.assertEqual(result.headline_for("de"), result.headline_for("en"))

    def test_notes_absent_is_not_an_error(self):
        result = updater.check_for_update(
            opener=opener_returning(manifest_bytes(NEWER, notes=False))
        )
        self.assertTrue(result.update_available)
        self.assertEqual(result.notes_for("en"), [])

    # --- failure handling -------------------------------------------------
    # Being offline is normal use, not a crash. Every one of these must come
    # back as a reported error, never an exception.

    def test_offline_is_reported_not_raised(self):
        result = updater.check_for_update(
            opener=opener_raising(urllib.error.URLError("no route to host"))
        )
        self.assertFalse(result.ok)
        self.assertIn("Could not reach", result.error or "")

    def test_timeout_is_reported_not_raised(self):
        result = updater.check_for_update(opener=opener_raising(TimeoutError("timed out")))
        self.assertFalse(result.ok)
        self.assertIn("Could not reach", result.error or "")

    def test_404_names_the_likely_cause(self):
        # 404 is what a private repo or a pre-release-only release returns, and
        # both are easy to misconfigure - so the message should be specific.
        error = urllib.error.HTTPError("url", 404, "Not Found", {}, None)
        result = updater.check_for_update(opener=opener_raising(error))
        self.assertFalse(result.ok)
        self.assertIn("404", result.error or "")
        self.assertIn("No published release", result.error or "")

    def test_http_500_is_reported_with_the_code(self):
        error = urllib.error.HTTPError("url", 500, "Server Error", {}, None)
        result = updater.check_for_update(opener=opener_raising(error))
        self.assertFalse(result.ok)
        self.assertIn("500", result.error or "")

    def test_unexpected_exception_does_not_escape(self):
        result = updater.check_for_update(opener=opener_raising(RuntimeError("boom")))
        self.assertFalse(result.ok)
        self.assertIn("RuntimeError", result.error or "")

    def test_malformed_json_is_reported(self):
        result = updater.check_for_update(opener=opener_returning(b"{not json"))
        self.assertFalse(result.ok)
        self.assertIn("could not be read", (result.error or "").lower())

    def test_non_utf8_body_is_reported(self):
        result = updater.check_for_update(opener=opener_returning(b"\xff\xfe\x00garbage"))
        self.assertFalse(result.ok)

    def test_manifest_missing_fields_is_reported(self):
        for payload in [
            {},
            {"version": "0.2.0"},
            {"version": "0.2.0", "file": {}},
            {"version": "0.2.0", "file": {"sha256": "a" * 64}},          # no download_url
            {"version": "", "file": {"sha256": "a" * 64}, "download_url": "https://x.invalid/a"},
            {"version": "0.2.0", "file": {"sha256": "short"}, "download_url": "https://x.invalid/a"},
            {"version": "0.2.0", "file": {"sha256": "a" * 64}, "download_url": "javascript:alert(1)"},
        ]:
            with self.subTest(payload=payload):
                body = json.dumps(payload).encode("utf-8")
                result = updater.check_for_update(opener=opener_returning(body))
                self.assertFalse(result.ok, f"should have been rejected: {payload}")
                self.assertIn("malformed", (result.error or "").lower())

    def test_unparseable_remote_version_is_reported(self):
        body = json.dumps({
            "version": "not-a-version",
            "download_url": (
                "https://github.com/ElvaraStudio/unitrans-portal/releases/download/"
                "vnot-a-version/Unitrans-Portal-not-a-version.exe"
            ),
            "file": {"sha256": "a" * 64},
        }).encode("utf-8")
        result = updater.check_for_update(opener=opener_returning(body))
        self.assertFalse(result.ok)
        self.assertIn("unreadable version", result.error or "")

    def test_manifest_url_is_the_latest_release_endpoint(self):
        self.assertEqual(updater.MANIFEST_URL, updater.MANIFEST_URL.strip())
        self.assertIn("/releases/latest/download/manifest.json", updater.MANIFEST_URL)

    def test_manifest_rejects_plain_http_downloads(self):
        payload = json.loads(manifest_bytes(NEWER))
        payload["download_url"] = payload["download_url"].replace("https://", "http://")
        result = updater.check_for_update(opener=opener_returning(json.dumps(payload).encode()))
        self.assertFalse(result.ok)
        self.assertIn("official GitHub", result.error or "")


class TestUpdateDownload(unittest.TestCase):
    def setUp(self):
        self._tmp = tempfile.TemporaryDirectory()
        self.destination = Path(self._tmp.name)
        self.payload = b"verified updater payload" * 1700
        self.result = updater.UpdateCheckResult(
            status="update-available",
            current_version="0.4.3",
            remote_version="0.4.4",
            download_url=(
                "https://github.com/ElvaraStudio/unitrans-portal/releases/download/"
                "v0.4.4/Unitrans-Portal-0.4.4.exe"
            ),
            size_bytes=len(self.payload),
            sha256=hashlib.sha256(self.payload).hexdigest(),
        )

    def tearDown(self):
        self._tmp.cleanup()

    def test_download_streams_and_verifies_sha256(self):
        progress = []
        path = updater.download_update(
            self.result,
            destination_dir=self.destination,
            opener=opener_returning(self.payload),
            progress_callback=lambda received, total: progress.append((received, total)),
        )
        self.assertEqual(path.read_bytes(), self.payload)
        self.assertEqual(hashlib.sha256(path.read_bytes()).hexdigest(), self.result.sha256)
        self.assertEqual(progress[-1], (len(self.payload), len(self.payload)))
        self.assertEqual(list(self.destination.glob("*.part")), [])

    def test_hash_mismatch_discards_partial_download(self):
        self.result.sha256 = "b" * 64
        with self.assertRaisesRegex(updater.UpdateDownloadError, "SHA-256"):
            updater.download_update(
                self.result,
                destination_dir=self.destination,
                opener=opener_returning(self.payload),
            )
        self.assertEqual(list(self.destination.iterdir()), [])

    def test_incomplete_download_is_rejected_and_discarded(self):
        self.result.size_bytes += 1
        with self.assertRaisesRegex(updater.UpdateDownloadError, "size did not match"):
            updater.download_update(
                self.result,
                destination_dir=self.destination,
                opener=opener_returning(self.payload),
            )
        self.assertEqual(list(self.destination.iterdir()), [])

    def test_cancelled_download_discards_partial_file(self):
        cancelled = threading.Event()
        cancelled.set()
        with self.assertRaises(updater.UpdateCancelled):
            updater.download_update(
                self.result,
                destination_dir=self.destination,
                opener=opener_returning(self.payload),
                cancel_event=cancelled,
            )
        self.assertEqual(list(self.destination.iterdir()), [])

    def test_download_rejects_untrusted_redirect(self):
        class RedirectedResponse(FakeResponse):
            def geturl(self):
                return "https://evil.example/payload.exe"

        with self.assertRaisesRegex(updater.UpdateDownloadError, "untrusted host"):
            updater.download_update(
                self.result,
                destination_dir=self.destination,
                opener=lambda request, timeout=None: RedirectedResponse(self.payload),
            )
        self.assertEqual(list(self.destination.iterdir()), [])

    def test_helper_script_waits_verifies_and_has_a_rollback(self):
        target = self.destination / "Chris's Unitrans Portal.exe"
        staged = self.destination / "verified update.exe"
        script = updater.build_update_helper_script(target, staged, "a" * 64, 321)
        self.assertIn("Chris''s Unitrans Portal.exe", script)
        self.assertIn("Waiting for the app to close", script)
        self.assertIn("Get-UpdateFileHash", script)
        self.assertIn("$target.previous", script)
        self.assertIn("Rollback failed", script)
        self.assertIn("321", script)


class TestSizeFormatting(unittest.TestCase):
    def test_human_readable(self):
        self.assertEqual(updater.format_size(24_403_427), "23.3 MB")
        self.assertEqual(updater.format_size(2048), "2 KB")
        self.assertEqual(updater.format_size(512), "512 bytes")

    def test_missing_size_is_blank_not_zero(self):
        for value in (None, 0, -1):
            self.assertEqual(updater.format_size(value), "")


class TestSettings(unittest.TestCase):
    def setUp(self):
        self._tmp = tempfile.TemporaryDirectory()
        self.path = Path(self._tmp.name) / "settings.json"

    def tearDown(self):
        self._tmp.cleanup()

    def test_defaults_when_file_absent(self):
        loaded = app_settings.load_settings(self.path)
        self.assertEqual(loaded["auto_check_updates"], False)
        self.assertEqual(loaded["last_update_check"], "")

    def test_round_trip(self):
        app_settings.save_settings({"auto_check_updates": True, "last_seen_version": "0.2.0"}, self.path)
        loaded = app_settings.load_settings(self.path)
        self.assertTrue(loaded["auto_check_updates"])
        self.assertEqual(loaded["last_seen_version"], "0.2.0")

    def test_corrupt_file_falls_back_to_defaults(self):
        # Losing a preference is trivial; refusing to start is not.
        self.path.write_text("{{{ not json", encoding="utf-8")
        loaded = app_settings.load_settings(self.path)
        self.assertEqual(loaded["auto_check_updates"], False)

    def test_unknown_keys_are_preserved(self):
        # An older build must not discard settings written by a newer one.
        app_settings.save_settings({"future_option": "keep me"}, self.path)
        self.assertEqual(app_settings.load_settings(self.path)["future_option"], "keep me")

    def test_update_settings_merges(self):
        app_settings.save_settings({"auto_check_updates": True}, self.path)
        merged = app_settings.update_settings(self.path, last_seen_version="0.3.0")
        self.assertTrue(merged["auto_check_updates"])
        self.assertEqual(merged["last_seen_version"], "0.3.0")

    def test_save_creates_missing_directories(self):
        nested = Path(self._tmp.name) / "a" / "b" / "settings.json"
        self.assertTrue(app_settings.save_settings({"auto_check_updates": True}, nested))
        self.assertTrue(nested.exists())

    def test_save_leaves_no_temp_files(self):
        app_settings.save_settings({"auto_check_updates": True}, self.path)
        leftovers = [p.name for p in self.path.parent.iterdir() if p.name != "settings.json"]
        self.assertEqual(leftovers, [], f"temp files left behind: {leftovers}")

    def test_unwritable_location_reports_failure_without_raising(self):
        # A read-only or full disk must not stop the user from working.
        result = app_settings.save_settings({"auto_check_updates": True}, Path("\0invalid/settings.json"))
        self.assertFalse(result)

    def test_settings_live_outside_the_install_folder(self):
        # The app can be replaced in place by an update, so settings must not
        # sit next to the executable.
        location = str(app_settings.settings_path())
        self.assertNotIn("unitrans-portal", location.lower())
        self.assertIn("UnitransPortal", location)


if __name__ == "__main__":
    unittest.main(verbosity=2)
