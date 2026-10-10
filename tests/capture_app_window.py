"""Launch a built executable and capture its window, optionally after a click.

Verifies the *packaged* build visually: whether a change that works from source
also works once PyInstaller has frozen it.

The window is captured by asking Tk's own handle rather than searching by title,
because "Settings" is also the name of a Windows app and a title search will
happily capture the wrong window.

    python tests/capture_app_window.py app.exe --out shot.png
    python tests/capture_app_window.py app.exe --out dark.png --click-dark-toggle
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

user32 = ctypes.windll.user32
WINDOW_TEXT_LIMIT = 512


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


def settle(hwnd: int, stable_for: float = 1.5, timeout: float = 20.0) -> tuple[int, int, int, int]:
    """Wait until the window rectangle stops changing."""
    class RECT(ctypes.Structure):
        _fields_ = [("left", ctypes.c_long), ("top", ctypes.c_long),
                    ("right", ctypes.c_long), ("bottom", ctypes.c_long)]

    def rect() -> tuple[int, int, int, int]:
        value = RECT()
        user32.GetWindowRect(hwnd, ctypes.byref(value))
        return (value.left, value.top, value.right - value.left, value.bottom - value.top)

    deadline = time.time() + timeout
    last = rect()
    stable_since = time.time()
    while time.time() < deadline:
        time.sleep(0.3)
        current = rect()
        if current != last:
            last, stable_since = current, time.time()
        elif time.time() - stable_since >= stable_for:
            return current
    return last


def average_brightness(hwnd: int) -> float:
    """Mean luminance of the window, 0-255. Distinguishes light from dark."""
    from PIL import Image

    x, y, width, height = settle(hwnd, stable_for=0.6, timeout=10.0)
    if width < 10 or height < 10:
        return -1.0

    hdc = user32.GetWindowDC(hwnd)
    gdi32 = ctypes.windll.gdi32
    memory_dc = gdi32.CreateCompatibleDC(hdc)
    bitmap = gdi32.CreateCompatibleBitmap(hdc, width, height)
    gdi32.SelectObject(memory_dc, bitmap)
    # PW_RENDERFULLCONTENT so a composited window captures correctly.
    user32.PrintWindow(hwnd, memory_dc, 2)

    class BITMAPINFOHEADER(ctypes.Structure):
        _fields_ = [
            ("biSize", wintypes.DWORD), ("biWidth", wintypes.LONG),
            ("biHeight", wintypes.LONG), ("biPlanes", wintypes.WORD),
            ("biBitCount", wintypes.WORD), ("biCompression", wintypes.DWORD),
            ("biSizeImage", wintypes.DWORD), ("biXPelsPerMeter", wintypes.LONG),
            ("biYPelsPerMeter", wintypes.LONG), ("biClrUsed", wintypes.DWORD),
            ("biClrImportant", wintypes.DWORD),
        ]

    header = BITMAPINFOHEADER()
    header.biSize = ctypes.sizeof(BITMAPINFOHEADER)
    header.biWidth = width
    header.biHeight = -height
    header.biPlanes = 1
    header.biBitCount = 32

    buffer = ctypes.create_string_buffer(width * height * 4)
    gdi32.GetDIBits(memory_dc, bitmap, 0, height, buffer, ctypes.byref(header), 0)
    image = Image.frombuffer("RGBA", (width, height), buffer, "raw", "BGRA", 0, 1).convert("L")

    gdi32.DeleteObject(bitmap)
    gdi32.DeleteDC(memory_dc)
    user32.ReleaseDC(hwnd, hdc)

    pixels = list(image.getdata())
    return sum(pixels) / len(pixels)


def save_screenshot(hwnd: int, destination: Path) -> bool:
    from PIL import Image

    x, y, width, height = settle(hwnd, stable_for=0.6, timeout=10.0)
    hdc = user32.GetWindowDC(hwnd)
    gdi32 = ctypes.windll.gdi32
    memory_dc = gdi32.CreateCompatibleDC(hdc)
    bitmap = gdi32.CreateCompatibleBitmap(hdc, width, height)
    gdi32.SelectObject(memory_dc, bitmap)
    user32.PrintWindow(hwnd, memory_dc, 2)

    class BITMAPINFOHEADER(ctypes.Structure):
        _fields_ = [
            ("biSize", wintypes.DWORD), ("biWidth", wintypes.LONG),
            ("biHeight", wintypes.LONG), ("biPlanes", wintypes.WORD),
            ("biBitCount", wintypes.WORD), ("biCompression", wintypes.DWORD),
            ("biSizeImage", wintypes.DWORD), ("biXPelsPerMeter", wintypes.LONG),
            ("biYPelsPerMeter", wintypes.LONG), ("biClrUsed", wintypes.DWORD),
            ("biClrImportant", wintypes.DWORD),
        ]

    header = BITMAPINFOHEADER()
    header.biSize = ctypes.sizeof(BITMAPINFOHEADER)
    header.biWidth = width
    header.biHeight = -height
    header.biPlanes = 1
    header.biBitCount = 32

    buffer = ctypes.create_string_buffer(width * height * 4)
    gdi32.GetDIBits(memory_dc, bitmap, 0, height, buffer, ctypes.byref(header), 0)
    image = Image.frombuffer("RGBA", (width, height), buffer, "raw", "BGRA", 0, 1)

    destination.parent.mkdir(parents=True, exist_ok=True)
    image.convert("RGB").save(destination)

    gdi32.DeleteObject(bitmap)
    gdi32.DeleteDC(memory_dc)
    user32.ReleaseDC(hwnd, hdc)
    return True


def click(hwnd: int, client_x: int, client_y: int) -> bool:
    """Click at a point in the window's client area, as a real mouse would.

    ``SendMessage(WM_LBUTTONDOWN)`` is not enough: a Tk button tracks the press
    and release across messages and only fires when it sees a consistent
    press/release pair at the same position, so posting messages straight to the
    window is silently ignored. ``SendInput`` injects a genuine mouse event
    through the normal input queue instead.

    This requires the window to actually be in the foreground; a background
    window will not receive the synthetic click. Returns False if the setup
    failed, so a caller can tell "the app ignored the click" apart from "the
    click never happened".
    """
    import ctypes.wintypes as wintypes

    point = wintypes.POINT(client_x, client_y)
    if not user32.ClientToScreen(hwnd, ctypes.byref(point)):
        return False

    screen_width = user32.GetSystemMetrics(0)
    screen_height = user32.GetSystemMetrics(1)
    MOUSEEVENTF_MOVE = 0x0001
    MOUSEEVENTF_ABSOLUTE = 0x8000
    MOUSEEVENTF_LEFTDOWN = 0x0002
    MOUSEEVENTF_LEFTUP = 0x0004

    class MOUSEINPUT(ctypes.Structure):
        _fields_ = [
            ("dx", ctypes.c_long),
            ("dy", ctypes.c_long),
            ("mouseData", ctypes.c_ulong),
            ("dwFlags", ctypes.c_ulong),
            ("time", ctypes.c_ulong),
            ("dwExtraInfo", ctypes.POINTER(ctypes.c_ulong)),
        ]

    class _INPUTUNION(ctypes.Union):
        _fields_ = [("mi", MOUSEINPUT)]

    class INPUT(ctypes.Structure):
        _fields_ = [("type", ctypes.c_ulong), ("union", _INPUTUNION)]

    absolute_x = int(point.x * 65535 / max(1, screen_width - 1))
    absolute_y = int(point.y * 65535 / max(1, screen_height - 1))

    sent = []

    def send(flags: int) -> None:
        item = INPUT(
            type=0,
            union=_INPUTUNION(
                mi=MOUSEINPUT(absolute_x, absolute_y, 0, flags | MOUSEEVENTF_ABSOLUTE, 0, None)
            ),
        )
        sent.append(user32.SendInput(1, ctypes.byref(item), ctypes.sizeof(INPUT)))

    user32.SetForegroundWindow(hwnd)
    time.sleep(0.4)
    send(MOUSEEVENTF_MOVE)
    time.sleep(0.15)
    send(MOUSEEVENTF_LEFTDOWN)
    time.sleep(0.1)
    send(MOUSEEVENTF_LEFTUP)
    # SendInput returns the number of events inserted; anything else means the
    # input was blocked (a background or elevated window) rather than ignored.
    return all(count == 1 for count in sent)


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("exe")
    parser.add_argument("--out", help="Where to write the screenshot.")
    parser.add_argument("--startup-timeout", type=float, default=90.0)
    parser.add_argument(
        "--click-at",
        help="Client x,y to click before capturing, e.g. '1000,45'. Used to "
             "exercise a button in the frozen build, where the widget tree is "
             "not reachable from outside.",
    )
    parser.add_argument(
        "--expect",
        choices=["light", "dark"],
        help="Require the window to be predominantly that brightness.",
    )
    args = parser.parse_args(argv)

    exe = Path(args.exe)
    if not exe.exists():
        print(f"Not found: {exe}")
        return 1

    print(f"launching {exe.name} ...")
    process = subprocess.Popen([str(exe)])
    failures: list[str] = []

    try:
        hwnd = find_app_window(time.time() + args.startup_timeout)
        if not hwnd:
            print("FAIL: no window appeared")
            return 1
        settle(hwnd, stable_for=2.0)

        before = average_brightness(hwnd)
        print(f"window brightness before: {before:.1f} / 255")

        if args.click_at:
            client_x, client_y = (int(part) for part in args.click_at.split(","))
            print(f"clicking client ({client_x}, {client_y})")
            delivered = click(hwnd, client_x, client_y)
            if not delivered:
                failures.append(
                    "SendInput was blocked, so the click never reached the app - "
                    "this says nothing about the app itself"
                )
            time.sleep(2.5)
            settle(hwnd, stable_for=1.0)
            after = average_brightness(hwnd)
            print(f"window brightness after : {after:.1f} / 255")
            if delivered and abs(after - before) < 15:
                failures.append(
                    f"clicking did not visibly change the window "
                    f"({before:.1f} -> {after:.1f})"
                )
            before = after

        if args.expect:
            dark = before < 128
            if args.expect == "dark" and not dark:
                failures.append(f"expected a dark window, brightness {before:.1f}")
            if args.expect == "light" and dark:
                failures.append(f"expected a light window, brightness {before:.1f}")
            else:
                print(f"window is {'dark' if dark else 'light'}, as expected")

        if args.out:
            save_screenshot(hwnd, Path(args.out))
            print(f"screenshot: {args.out}")

        if failures:
            for failure in failures:
                print(f"FAIL: {failure}")
            return 1
        print("PASS: packaged window verified.")
        return 0
    finally:
        process.terminate()
        try:
            process.wait(timeout=10)
        except subprocess.TimeoutExpired:
            process.kill()
        time.sleep(0.8)
        subprocess.run(["taskkill", "/F", "/IM", exe.name], capture_output=True, check=False)


if __name__ == "__main__":
    sys.exit(main())
