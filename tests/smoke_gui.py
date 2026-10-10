"""Launch the real GUI, confirm it rendered, and optionally save a screenshot.

    python tests/smoke_gui.py --screenshot out.png

Run from the project root with the build venv. This drives the actual
``WaybillSorterApp`` rather than inspecting the source, so it catches failures
that only appear once the widget tree is built - a missing i18n key, a bad
geometry string, an exception during layout.

If the window is not found within the timeout the script exits non-zero and
prints the Tk window that *was* created, so a title mismatch is obvious.
"""

from __future__ import annotations

import argparse
import ctypes
import sys
import time
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))

import tkinter as tk  # noqa: E402

from version import version_label  # noqa: E402


def find_window_by_title(title: str, timeout: float = 20.0) -> int:
    """Return the first top-level window whose text starts with ``title``."""
    user32 = ctypes.windll.user32
    buffer = ctypes.create_unicode_buffer(512)
    deadline = time.time() + timeout

    while time.time() < deadline:
        found: list[int] = []

        @ctypes.WINFUNCTYPE(ctypes.c_bool, ctypes.c_void_p, ctypes.c_void_p)
        def callback(hwnd, _lparam):
            if not user32.IsWindowVisible(hwnd):
                return True
            user32.GetWindowTextW(hwnd, buffer, 512)
            text = buffer.value
            if text and (text == title or text.startswith(title)):
                # Skip tool windows and hidden helpers.
                if user32.GetWindow(hwnd, 4) == 0:  # GW_OWNER
                    found.append(hwnd)
                    return False
            return True

        user32.EnumWindows(callback, 0)
        if found:
            return found[0]
        time.sleep(0.25)

    return 0


def capture_window(hwnd: int, destination: Path) -> bool:
    """Screenshot a single window by handle."""
    try:
        from PIL import Image
    except ImportError:
        print("Pillow not available; skipping screenshot.")
        return False

    import ctypes.wintypes as wintypes

    user32 = ctypes.windll.user32
    gdi32 = ctypes.windll.gdi32

    # Bring the window forward so the capture is not occlusion.
    user32.SetForegroundWindow(hwnd)
    time.sleep(1.0)

    rect = wintypes.RECT()
    if not user32.GetWindowRect(hwnd, ctypes.byref(rect)):
        print("GetWindowRect failed.")
        return False
    width = rect.right - rect.left
    height = rect.bottom - rect.top

    hdc_window = user32.GetWindowDC(hwnd)
    hdc_mem = gdi32.CreateCompatibleDC(hdc_window)
    bitmap = gdi32.CreateCompatibleBitmap(hdc_window, width, height)
    gdi32.SelectObject(hdc_mem, bitmap)

    # PW_RENDERFULLCONTENT (2) captures composited content, so it works even
    # when the window is not the top-most one.
    if not user32.PrintWindow(hwnd, hdc_mem, 2):
        print("PrintWindow failed; falling back to BitBlt.")
        gdi32.BitBlt(hdc_mem, 0, 0, width, height, hdc_window, 0, 0, 0x00CC0020)

    class BITMAPINFOHEADER(ctypes.Structure):
        _fields_ = [
            ("biSize", wintypes.DWORD),
            ("biWidth", wintypes.LONG),
            ("biHeight", wintypes.LONG),
            ("biPlanes", wintypes.WORD),
            ("biBitCount", wintypes.WORD),
            ("biCompression", wintypes.DWORD),
            ("biSizeImage", wintypes.DWORD),
            ("biXPelsPerMeter", wintypes.LONG),
            ("biYPelsPerMeter", wintypes.LONG),
            ("biClrUsed", wintypes.DWORD),
            ("biClrImportant", wintypes.DWORD),
        ]

    header = BITMAPINFOHEADER()
    header.biSize = ctypes.sizeof(BITMAPINFOHEADER)
    header.biWidth = width
    header.biHeight = -height  # negative = top-down rows
    header.biPlanes = 1
    header.biBitCount = 32
    header.biCompression = 0

    buffer_size = width * height * 4
    buffer = ctypes.create_string_buffer(buffer_size)
    gdi32.GetDIBits(hdc_mem, bitmap, 0, height, buffer, ctypes.byref(header), 0)

    image = Image.frombuffer("RGBA", (width, height), buffer, "raw", "BGRA", 0, 1)
    destination.parent.mkdir(parents=True, exist_ok=True)
    image.convert("RGB").save(destination)

    gdi32.DeleteObject(bitmap)
    gdi32.DeleteDC(hdc_mem)
    user32.ReleaseDC(hwnd, hdc_window)

    print(f"Screenshot saved: {destination}")
    return True


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="Verify the GUI launches and renders.")
    parser.add_argument("--screenshot", help="Save a PNG of the main window here.")
    parser.add_argument("--keep-open", type=float, default=6.0, help="Seconds to leave the window up.")
    args = parser.parse_args(argv)

    from waybill_sorter import WaybillSorterApp

    app = WaybillSorterApp()
    expected_title = app.title()
    print(f"Tk reports window title: {expected_title!r}")

    if not expected_title or expected_title == "window_title":
        print("FAIL: the window title is empty or an unresolved i18n key.")
        app.destroy()
        return 1

    # Render one frame so layout errors surface here rather than on screen.
    app.update_idletasks()
    app.update()

    footer_expected = version_label()
    print(f"Expected footer text    : {footer_expected!r}")

    # Confirm the footer label widget actually carries the version string.
    footer_found = False
    def walk(widget):
        nonlocal footer_found
        for child in widget.winfo_children():
            if isinstance(child, tk.Label):
                try:
                    if footer_expected in str(child.cget("text")):
                        footer_found = True
                except tk.TclError:
                    pass
            walk(child)
    walk(app)
    if not footer_found:
        print("FAIL: no label in the widget tree shows the version footer.")
        app.destroy()
        return 1
    print("PASS: version footer present in the widget tree.")

    app.deiconify()
    app.lift()
    app.update()

    hwnd = find_window_by_title(expected_title)
    if not hwnd:
        print(f"FAIL: no visible top-level window titled {expected_title!r} appeared.")
        app.destroy()
        return 1
    print(f"PASS: window handle 0x{hwnd:x} found and visible.")

    if args.screenshot:
        capture_window(hwnd, Path(args.screenshot))

    # Let it render, then close through Tk to prove clean shutdown.
    deadline = time.time() + args.keep_open
    while time.time() < deadline:
        app.update()
        time.sleep(0.05)

    app.destroy()
    print("PASS: GUI built, rendered, and shut down cleanly.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
