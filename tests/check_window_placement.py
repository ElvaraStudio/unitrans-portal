"""Launch a built executable and report where its window ends up.

The onefile build unpacks ~23 MB to a temporary folder before the window
appears, and then maximises shortly afterwards, so a fixed sleep measures the
window mid-startup. This polls until the window's rectangle stops changing and
only then reports - which is what a user actually ends up looking at.

    python tests/check_window_placement.py "C:\\path\\to\\app.exe"
    python tests/check_window_placement.py path.exe --expect-monitor 1
"""

from __future__ import annotations

import argparse
import ctypes
import ctypes.wintypes as wintypes
import subprocess
import sys
import time
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))

from monitors import list_monitors  # noqa: E402

user32 = ctypes.windll.user32
WINDOW_TEXT_LIMIT = 512


class RECT(ctypes.Structure):
    _fields_ = [("left", ctypes.c_long), ("top", ctypes.c_long),
                ("right", ctypes.c_long), ("bottom", ctypes.c_long)]


def find_app_window(deadline: float) -> int:
    """Handle of the first visible window titled 'Unitrans Portal'."""
    buffer = ctypes.create_unicode_buffer(WINDOW_TEXT_LIMIT)

    while time.time() < deadline:
        found: list[int] = []

        @ctypes.WINFUNCTYPE(ctypes.c_bool, ctypes.c_void_p, ctypes.c_void_p)
        def callback(hwnd, _lparam):
            if not user32.IsWindowVisible(hwnd):
                return True
            user32.GetWindowTextW(hwnd, buffer, WINDOW_TEXT_LIMIT)
            if buffer.value.startswith("Unitrans Portal"):
                found.append(hwnd)
                return False
            return True

        user32.EnumWindows(callback, 0)
        if found:
            return found[0]
        time.sleep(0.4)
    return 0


def window_rect(hwnd: int) -> tuple[int, int, int, int]:
    rect = RECT()
    user32.GetWindowRect(hwnd, ctypes.byref(rect))
    return (rect.left, rect.top, rect.right - rect.left, rect.bottom - rect.top)


def settle(hwnd: int, stable_for: float = 2.0, timeout: float = 30.0) -> tuple[int, int, int, int]:
    """Wait until the window rectangle stops changing, then return it."""
    deadline = time.time() + timeout
    last = window_rect(hwnd)
    stable_since = time.time()

    while time.time() < deadline:
        time.sleep(0.25)
        current = window_rect(hwnd)
        if current != last:
            last = current
            stable_since = time.time()
        elif time.time() - stable_since >= stable_for:
            return current
    return last


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="Check where a built app opens.")
    parser.add_argument("exe", nargs="?", help="Path to the executable to launch.")
    parser.add_argument("--expect-monitor", type=int, default=None,
                        help="Monitor index the window should land on.")
    parser.add_argument("--startup-timeout", type=float, default=90.0)
    parser.add_argument("--source", action="store_true",
                        help="Launch waybill_sorter.py instead of an executable.")
    parser.add_argument("--maximised", action="store_true",
                        help="Require the window to fill the monitor. The app opens "
                             "windowed by default so it does not cover the target screen.")
    parser.add_argument("--min-size", type=int, default=800,
                        help="Smallest acceptable window width in windowed mode.")
    args = parser.parse_args(argv)

    monitors = list_monitors()
    print("monitors detected:")
    for monitor in monitors:
        print(f"  [{monitor.index}] {monitor.describe()}")

    if args.source:
        command = [sys.executable, str(ROOT / "waybill_sorter.py")]
        label = "waybill_sorter.py (from source)"
    else:
        if not args.exe:
            print("Provide an executable path, or use --source.")
            return 2
        exe = Path(args.exe)
        if not exe.exists():
            print(f"Not found: {exe}")
            return 1
        command = [str(exe)]
        label = exe.name

    print(f"\nlaunching {label} ...")
    process = subprocess.Popen(command)

    try:
        hwnd = find_app_window(time.time() + args.startup_timeout)
        if not hwnd:
            print(f"FAIL: no window appeared within {args.startup_timeout:.0f}s")
            return 1

        x, y, width, height = settle(hwnd)
        print(f"window settled at {width}x{height} at ({x},{y})")

        # Which monitor contains the window's centre?
        centre_x, centre_y = x + width // 2, y + height // 2
        containing = [
            monitor for monitor in monitors
            if monitor.left <= centre_x < monitor.right and monitor.top <= centre_y < monitor.bottom
        ]

        failures: list[str] = []
        if not containing:
            failures.append("window centre is not on any detected monitor")
            monitor = None
        else:
            monitor = containing[0]
            print(f"landed on monitor [{monitor.index}] ({monitor.describe()})")
            # A maximised window is inset by the frame and sits inside the work
            # area, so the tolerance here is deliberately small: a window whose
            # outer rect starts above y=0 has its title bar under the screen
            # edge, and one that ends below the work area has its bottom under
            # the taskbar.
            if x < monitor.left:
                failures.append(f"window starts left of the monitor (x={x} < {monitor.left})")
            if y < monitor.top:
                failures.append(f"window starts above the monitor (y={y} < {monitor.top})")
            if x + width > monitor.right:
                failures.append(
                    f"window overflows the right edge (right={x + width} > {monitor.right})"
                )
            if y + height > monitor.bottom:
                failures.append(
                    f"window overflows the bottom edge (bottom={y + height} > {monitor.bottom})"
                )
            if args.maximised:
                if width < monitor.width * 0.9:
                    failures.append(
                        f"window is not maximised ({width} of {monitor.width}px wide)"
                    )
            else:
                # Windowed mode: it must be usable, not a sliver, and must not
                # be larger than the screen it is on.
                if width < args.min_size:
                    failures.append(
                        f"window is too small to use ({width}px wide, want >= {args.min_size})"
                    )
                if width > monitor.width or height > monitor.height:
                    failures.append(
                        f"window is larger than the monitor ({width}x{height} on "
                        f"{monitor.width}x{monitor.height})"
                    )
            if args.expect_monitor is not None and monitor.index != args.expect_monitor:
                failures.append(
                    f"expected monitor {args.expect_monitor}, got {monitor.index}"
                )

        print()
        if failures:
            for failure in failures:
                print(f"FAIL: {failure}")
            return 1

        mode = "maximised" if args.maximised else "windowed"
        print(f"PASS: window opened {mode} on the expected monitor and fits within it.")
        return 0
    finally:
        process.terminate()
        try:
            process.wait(timeout=10)
        except subprocess.TimeoutExpired:
            process.kill()
        # The onefile child can outlive the bootloader process.
        time.sleep(1.0)
        subprocess.run(
            ["taskkill", "/F", "/IM", exe.name],
            capture_output=True,
            check=False,
        )


if __name__ == "__main__":
    sys.exit(main())
