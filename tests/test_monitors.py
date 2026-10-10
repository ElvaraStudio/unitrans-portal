"""Tests for monitor detection and window placement.

Placement is verified against synthetic monitor lists rather than the machine's
real displays, so the suite passes on a single-monitor build machine while
still proving the multi-monitor arithmetic - and, critically, that a saved
preference for a second screen degrades safely when there is no second screen.
"""

from __future__ import annotations

import sys
import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))

import monitors  # noqa: E402
from monitors import Monitor  # noqa: E402


def make_monitor(
    index: int,
    left: int,
    top: int,
    width: int,
    height: int,
    primary: bool = False,
    taskbar: int = 0,
) -> Monitor:
    return Monitor(
        index=index,
        left=left,
        top=top,
        right=left + width,
        bottom=top + height,
        primary=primary,
        work_left=left,
        work_top=top,
        work_right=left + width,
        work_bottom=top + height - taskbar,
    )


#: The developer's actual desk layout: primary 1682x1051 at the origin, a
#: 1920x1080 screen starting at x=2944, both with a taskbar.
TWO_SCREENS = [
    make_monitor(0, 0, 0, 1682, 1051, primary=True, taskbar=48),
    make_monitor(1, 2944, 0, 1920, 1080, taskbar=48),
]

ONE_SCREEN = [make_monitor(0, 0, 0, 1920, 1080, primary=True, taskbar=48)]


class TestPlacement(unittest.TestCase):
    def test_second_monitor_is_used_by_default(self):
        _w, _h, x, _y, monitor = monitors.placement_for(
            monitors.MONITOR_DEFAULT, 1400, 900, TWO_SCREENS
        )
        self.assertEqual(monitor.index, 1)
        self.assertGreaterEqual(x, 2944, "window should sit on the second screen")

    def test_primary_monitor_when_requested(self):
        _w, _h, x, _y, monitor = monitors.placement_for(0, 1400, 900, TWO_SCREENS)
        self.assertEqual(monitor.index, 0)
        # Centred within the primary work area, so the window starts inside it
        # rather than at the origin.
        self.assertGreaterEqual(x, monitor.work_left)
        self.assertLess(x, monitor.work_width)

    def test_missing_second_monitor_degrades_to_primary(self):
        # Somebody running this on a laptop must not get a window placed at
        # x=2944 on a screen that does not exist.
        _w, _h, x, _y, monitor = monitors.placement_for(
            monitors.MONITOR_DEFAULT, 1400, 900, ONE_SCREEN
        )
        self.assertEqual(monitor.index, 0)
        self.assertTrue(monitor.primary)
        self.assertLess(x, 1920, "window must land on the only screen available")

    def test_absurd_preference_degrades_to_primary(self):
        for preference in (7, 99, -3):
            with self.subTest(preference=preference):
                _w, _h, _x, _y, monitor = monitors.placement_for(
                    preference, 1400, 900, TWO_SCREENS
                )
                self.assertTrue(monitor.primary)

    def test_fill_occupies_the_whole_work_area(self):
        width, height, x, y, monitor = monitors.placement_for(
            monitors.MONITOR_DEFAULT, 100, 100, TWO_SCREENS, fill=True
        )
        self.assertEqual((x, y), (monitor.work_left, monitor.work_top))
        self.assertEqual(width, monitor.work_width)
        self.assertEqual(height, monitor.work_height)
        # 1080 tall monitor with a 48px taskbar.
        self.assertEqual(height, 1032)

    def test_window_is_clamped_to_the_work_area(self):
        # Asking for a window larger than the screen must not push it off the
        # bottom, where the taskbar and screen edge would hide it.
        width, height, x, y, monitor = monitors.placement_for(
            0, 99999, 99999, TWO_SCREENS
        )
        self.assertLessEqual(x + width, monitor.work_right)
        self.assertLessEqual(y + height, monitor.work_bottom)
        self.assertLess(height, monitor.height, "must leave room for the taskbar")

    def test_result_always_fits_inside_the_work_area(self):
        for preference in (0, 1, 99):
            for size in ((800, 600), (1400, 900), (1920, 1080), (4000, 3000)):
                with self.subTest(preference=preference, size=size):
                    width, height, x, y, monitor = monitors.placement_for(
                        preference, size[0], size[1], TWO_SCREENS
                    )
                    self.assertGreaterEqual(x, monitor.work_left)
                    self.assertGreaterEqual(y, monitor.work_top)
                    self.assertLessEqual(x + width, monitor.work_right)
                    self.assertLessEqual(y + height, monitor.work_bottom)

    def test_negative_coordinates_are_supported(self):
        # A monitor placed to the left of the primary has negative x. Placement
        # must not clamp that to zero.
        left_screen = [
            make_monitor(0, 0, 0, 1920, 1080, primary=True),
            make_monitor(1, -1680, 0, 1680, 1050),
        ]
        _w, _h, x, _y, monitor = monitors.placement_for(1, 1200, 800, left_screen)
        self.assertEqual(monitor.index, 1)
        self.assertLess(x, 0)


class TestWindowPlacementHelpers(unittest.TestCase):
    """The Win32 helpers behind startup placement.

    These need a real Tk window, so they run against one - skipped rather than
    failed if Tk is unavailable.
    """

    def setUp(self):
        try:
            import tkinter as tk
        except ImportError:  # pragma: no cover
            self.skipTest("tkinter unavailable")
        try:
            self.window = tk.Tk()
        except Exception as exc:  # pragma: no cover - headless environment
            self.skipTest(f"no display available: {exc}")
        self.window.geometry("300x200+40+40")
        self.window.update_idletasks()

    def tearDown(self):
        try:
            self.window.destroy()
        except Exception:
            pass

    def test_toplevel_handle_differs_from_the_inner_client_window(self):
        # ShowWindow/SetWindowPos against the inner client handle are ignored,
        # which is what made maximising silently do nothing.
        inner = int(self.window.winfo_id())
        frame = monitors.toplevel_handle(self.window)
        self.assertGreater(frame, 0)
        self.assertNotEqual(frame, inner, "must resolve to the frame, not the client")

    def test_move_window_positions_the_window(self):
        moved = monitors.move_window(self.window, 111, 222, 640, 480)
        self.assertTrue(moved)
        self.window.update_idletasks()
        import ctypes
        import ctypes.wintypes as wintypes

        rect = wintypes.RECT()
        ctypes.windll.user32.GetWindowRect(
            monitors.toplevel_handle(self.window), ctypes.byref(rect)
        )
        # Allow a pixel of slack for frame rounding.
        self.assertLessEqual(abs(rect.left - 111), 2, f"x was {rect.left}")
        self.assertLessEqual(abs(rect.top - 222), 2, f"y was {rect.top}")

    def test_startup_placement_returns_the_target_monitor(self):
        monitor = monitors.apply_startup_placement(
            self.window, 1, maximise=False, monitors=TWO_SCREENS
        )
        self.assertIsNotNone(monitor)
        self.assertEqual(monitor.index, 1)

    def test_startup_placement_windowed_lands_on_the_chosen_screen(self):
        monitor = monitors.apply_startup_placement(
            self.window, 1, maximise=False, monitors=TWO_SCREENS,
            window_width=1200, window_height=800,
        )
        self.window.update_idletasks()
        import ctypes
        import ctypes.wintypes as wintypes

        rect = wintypes.RECT()
        ctypes.windll.user32.GetWindowRect(
            monitors.toplevel_handle(self.window), ctypes.byref(rect)
        )
        self.assertGreaterEqual(rect.left, monitor.left)
        self.assertGreaterEqual(rect.top, monitor.top)
        self.assertLessEqual(rect.right, monitor.right)
        self.assertLessEqual(rect.bottom, monitor.bottom)

    def test_startup_placement_survives_a_single_monitor(self):
        # The saved preference says "second screen" but there is only one.
        monitor = monitors.apply_startup_placement(
            self.window, monitors.MONITOR_DEFAULT, maximise=False, monitors=ONE_SCREEN
        )
        self.assertEqual(monitor.index, 0)

    def test_monitor_of_point(self):
        self.assertEqual(monitors.monitor_of_point(100, 100, TWO_SCREENS).index, 0)
        self.assertEqual(monitors.monitor_of_point(3000, 100, TWO_SCREENS).index, 1)
        self.assertIsNone(monitors.monitor_of_point(-5000, -5000, TWO_SCREENS))


class TestMonitorList(unittest.TestCase):
    def test_real_detection_returns_at_least_one(self):
        found = monitors.list_monitors()
        self.assertGreaterEqual(len(found), 1)
        self.assertTrue(all(isinstance(m, Monitor) for m in found))

    def test_exactly_one_monitor_is_marked_primary(self):
        found = monitors.list_monitors()
        self.assertEqual(
            sum(1 for monitor in found if monitor.primary),
            1,
            "exactly one display should be reported as primary",
        )

    def test_indices_are_contiguous_and_ordered(self):
        found = monitors.list_monitors()
        self.assertEqual([m.index for m in found], list(range(len(found))))
        self.assertEqual(
            [m.left for m in found],
            sorted(m.left for m in found),
            "monitors should be ordered left to right",
        )

    def test_work_area_is_never_larger_than_the_monitor(self):
        for monitor in monitors.list_monitors():
            with self.subTest(index=monitor.index):
                self.assertLessEqual(monitor.work_width, monitor.width)
                self.assertLessEqual(monitor.work_height, monitor.height)
                self.assertGreater(monitor.work_height, 200)

    def test_describe_mentions_usable_area_when_a_taskbar_exists(self):
        monitor = make_monitor(0, 0, 0, 1920, 1080, primary=True, taskbar=48)
        self.assertIn("1920x1080", monitor.describe())
        self.assertIn("1920x1032 usable", monitor.describe())

    def test_describe_omits_usable_area_when_there_is_no_taskbar(self):
        monitor = make_monitor(0, 0, 0, 1920, 1080, primary=True)
        self.assertNotIn("usable", monitor.describe())


if __name__ == "__main__":
    unittest.main(verbosity=2)
